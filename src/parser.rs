#![forbid(unsafe_code)]

use crate::{
    Information, Module, ModuleType, Reference, ReferenceControl, ReferenceOriginal,
    ReferenceProject, ReferenceRegistered, SysKind,
};
use codepage::to_encoding;
use encoding_rs::{CoderResult};
use nom::{
    bytes::complete::{tag, take},
    combinator::opt,
    error::{ErrorKind, ParseError},
    multi::length_data,
    number::complete::{le_u16, le_u32, le_u8},
    sequence::{preceded, tuple},
    Err::Error,
    IResult,
};

// This used to be part of the public interface prior to flattening this out into the
// [`Project`] struct.
// TODO: Re-evaluate whether this struct is strictly necessary, or can be removed.
/// Specifies information for the VBA project, including project information, project
/// references, and modules.
#[derive(Debug)]
pub(crate) struct ProjectInformation {
    /// Specifies version-independent information for the VBA project.
    pub information: Information,
    /// Specifies the external references of the VBA project.
    pub references: Vec<Reference>,
    /// Specifies the modules in the project.
    pub modules: Vec<Module>,
}

// TODO: Make this error private by translating to a crate-level error type
//       at the public parser interface.
#[derive(Debug, PartialEq)]
pub(crate) enum FormatError<I> {
    UnexpectedValue,
    Nom(I, ErrorKind),
}

impl<I> ParseError<I> for FormatError<I> {
    fn from_error_kind(input: I, kind: ErrorKind) -> Self {
        FormatError::Nom(input, kind)
    }
    fn append(_: I, _: ErrorKind, other: Self) -> Self {
        other
    }
}

fn uncompressed_chunk_parser(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
    Ok((&[], i.to_vec()))
}

fn compressed_chunk_parser(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
    // Initialize output storage; Chunks are at most 4096 decompressed bytes
    let mut result = Vec::<u8>::with_capacity(4096);
    // Loop until `i` is depleted
    let mut input = i;
    while !input.is_empty() {
        // Read FlagByte
        let (i_after_flag, flag_byte) = le_u8(input)?;
        input = i_after_flag; 

        // Loop over bits
        for flag_bit_index in 0..=7 {
            if input.is_empty() {
                return Ok((input, result));
            }
            let is_copy_token = (flag_byte & (1 << flag_bit_index)) != 0;
            if is_copy_token {
                if input.len() < 2 { 
                    return Err(nom::Err::Error(FormatError::Nom(input, ErrorKind::Eof)));
                }
                let (i_after_token_data, copy_token_raw) = le_u16(input)?;
                
                // --- Logic with dynamic bit_count (compatible with original file reading) ---
                let diff = result.len(); 
                let mut bit_count = 4_usize; 
                if diff > 0 { 
                    while (1_usize).checked_shl(bit_count as u32).map_or(false, |val| val < diff) && bit_count < 16 {
                        bit_count += 1;
                    }
                }
                // if diff is 0, bit_count remains 4.
                
                let length_mask = 0xffff_u16 >> bit_count; 
                let offset_mask = !length_mask;          
                
                let length_val = (copy_token_raw & length_mask) + 3; 
                let offset_val = ((copy_token_raw & offset_mask) >> (16 - bit_count)) + 1; 
                // --- End of dynamic bit_count logic ---

                let current_copy_length = length_val as usize;
                let lookback_distance = offset_val as usize;

                // This check was problematic. Original files passed when this was less strict or different.
                // For a CopyToken to be valid, lookback_distance must be > 0 and <= result.len().
                if lookback_distance == 0 || lookback_distance > result.len() {
                    return Err(nom::Err::Error(FormatError::Nom(input, ErrorKind::Verify)));
                }

                input = i_after_token_data; 

                let copy_start_index_in_result = result.len() - lookback_distance;
                for k in 0..current_copy_length {
                    result.push(result[copy_start_index_in_result + k]);
                }
            } else {
                // LiteralToken
                if input.is_empty() { 
                    return Err(nom::Err::Error(FormatError::Nom(input, ErrorKind::Eof)));
                }
                let (i_after_literal, byte) = le_u8(input)?;
                input = i_after_literal;
                result.push(byte);
            }
        }
    }
    Ok((input, result))
}

fn chunk_parser(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
    // CompressedChunkHeader (12 bits: size minus 3; 3 bits: 0b110; 1 bit: flag)
    // Delegate to specific parser (compressed/uncompressed) depending on the `flag`
    let (i, header_raw) = le_u16(i)?;
    // Check header magic (0b110) in bit positions 12..=14
    if (header_raw >> 12) & 0b111 != 0b011 {
        return Err(Error(FormatError::UnexpectedValue));
    }
    // Extract compressed/uncompressed flag
    let flag = ((header_raw >> 15) & 0b1) != 0;
    // Extract length
    let length = (header_raw & 0xfff) as usize + 1;

    let (chunk, remainder) = i.split_at(length);
    if flag {
        Ok((remainder, compressed_chunk_parser(chunk)?.1))
    } else {
        Ok((remainder, uncompressed_chunk_parser(chunk)?.1))
    }
}

/// Decompress a CompressedContainer.
pub(crate) fn decompress(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
    const COMPRESSED_CONTAINER_SIGNATURE: &[u8] = &[0x01];
    let (i, _) = tag(COMPRESSED_CONTAINER_SIGNATURE)(i)?;

    // This is the main `Chunk` parser:
    // * It parses 1 or more chunks, returning a `Vec<u8>` with decoded content.
    // * It appends the contents of the most recent `Chunk` to the existing decoded stream.
    // * If all data has been consumed, return an `Ok()` value.
    nom::combinator::all_consuming(nom::multi::fold_many1(
        chunk_parser,
        Vec::new,
        |mut acc: Vec<_>, data| {
            acc.extend(data);
            acc
        },
    ))(i)
}

// -------------------------------------------------------------------------
// -------------------------------------------------------------------------

// Several size fields in the binary format have fixed values.
const U32_FIXED_SIZE_4: &[u8] = &[0x04, 0x00, 0x00, 0x00];
const U32_FIXED_SIZE_2: &[u8] = &[0x02, 0x00, 0x00, 0x00];

fn parse_syskind(i: &[u8]) -> IResult<&[u8], SysKind, FormatError<&[u8]>> {
    const SYS_KIND_SIGNATURE: &[u8] = &[0x01, 0x00];
    let (i, sys_kind) = preceded(
        tuple((tag(SYS_KIND_SIGNATURE), tag(U32_FIXED_SIZE_4))),
        le_u32,
    )(i)?;
    match sys_kind {
        0x0000_0000 => Ok((i, SysKind::Win16)),
        0x0000_0001 => Ok((i, SysKind::Win32)),
        0x0000_0002 => Ok((i, SysKind::MacOs)),
        0x0000_0003 => Ok((i, SysKind::Win64)),
        _ => Err(Error(FormatError::UnexpectedValue)),
    }
}

fn parse_compat(i: &[u8]) -> IResult<&[u8], Option<u32>, FormatError<&[u8]>> {
    const COMPAT_SIGNATURE: &[u8] = &[0x4A, 0x00];
    let (i, compat) = opt(preceded(
        tuple((tag(COMPAT_SIGNATURE), tag(U32_FIXED_SIZE_4))),
        le_u32,
    ))(i)?;
    Ok((i, compat))
}

fn parse_lcid(i: &[u8]) -> IResult<&[u8], u32, FormatError<&[u8]>> {
    const LCID_SIGNATURE: &[u8] = &[0x02, 0x00];
    let (i, lcid) = preceded(tuple((tag(LCID_SIGNATURE), tag(U32_FIXED_SIZE_4))), le_u32)(i)?;
    Ok((i, lcid))
}

fn parse_lcid_invoke(i: &[u8]) -> IResult<&[u8], u32, FormatError<&[u8]>> {
    const LCID_INVOKE_SIGNATURE: &[u8] = &[0x14, 0x00];
    let (i, lcid_invoke) = preceded(
        tuple((tag(LCID_INVOKE_SIGNATURE), tag(U32_FIXED_SIZE_4))),
        le_u32,
    )(i)?;
    Ok((i, lcid_invoke))
}

fn parse_code_page(i: &[u8]) -> IResult<&[u8], u16, FormatError<&[u8]>> {
    const CODE_PAGE_SIGNATURE: &[u8] = &[0x03, 0x00];
    let (i, code_page) = preceded(
        tuple((tag(CODE_PAGE_SIGNATURE), tag(U32_FIXED_SIZE_2))),
        le_u16,
    )(i)?;
    Ok((i, code_page))
}

fn parse_name(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
    const NAME_SIGNATURE: &[u8] = &[0x04, 0x00];
    let (i, name) = preceded(tag(NAME_SIGNATURE), length_data(le_u32))(i)?;
    Ok((i, name.to_vec()))
}

fn parse_doc_string(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
    const DOC_STRING_SIGNATURE: &[u8] = &[0x05, 0x00];
    let (i, doc_string) = preceded(tag(DOC_STRING_SIGNATURE), length_data(le_u32))(i)?;
    Ok((i, doc_string.to_vec()))
}

fn parse_doc_string_unicode(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
    const DOC_STRING_UNICODE_SIGNATURE: &[u8] = &[0x40, 0x00];
    let (i, doc_string_unicode) =
        preceded(tag(DOC_STRING_UNICODE_SIGNATURE), length_data(le_u32))(i)?;
    // `doc_string_unicode` represents a sequence of UTF-16 code units. If its length is uneven,
    // the input is malformed.
    if (doc_string_unicode.len() & 1_usize) != 0 {
        Err(Error(FormatError::UnexpectedValue))
    } else {
        Ok((i, doc_string_unicode.to_vec()))
    }
}

fn parse_help_file_1(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
    const HELP_FILE_1_SIGNATURE: &[u8] = &[0x06, 0x00];
    let (i, help_file_1) = preceded(tag(HELP_FILE_1_SIGNATURE), length_data(le_u32))(i)?;
    Ok((i, help_file_1.to_vec()))
}

fn parse_help_file_2(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
    const HELP_FILE_2_SIGNATURE: &[u8] = &[0x3d, 0x00];
    let (i, help_file_2) = preceded(tag(HELP_FILE_2_SIGNATURE), length_data(le_u32))(i)?;
    Ok((i, help_file_2.to_vec()))
}

fn parse_help_context(i: &[u8]) -> IResult<&[u8], u32, FormatError<&[u8]>> {
    const HELP_CONTEXT_SIGNATURE: &[u8] = &[0x07, 0x00];
    let (i, help_context) = preceded(
        tuple((tag(HELP_CONTEXT_SIGNATURE), tag(U32_FIXED_SIZE_4))),
        le_u32,
    )(i)?;
    Ok((i, help_context))
}

fn parse_lib_flags(i: &[u8]) -> IResult<&[u8], u32, FormatError<&[u8]>> {
    const LIB_FLAGS_SIGNATURE: &[u8] = &[0x08, 0x00];
    let (i, lib_flags) = preceded(
        tuple((tag(LIB_FLAGS_SIGNATURE), tag(U32_FIXED_SIZE_4))),
        le_u32,
    )(i)?;
    Ok((i, lib_flags))
}

fn parse_version(i: &[u8]) -> IResult<&[u8], (u32, u16), FormatError<&[u8]>> {
    const VERSION_SIGNATURE: &[u8] = &[0x09, 0x00];
    let (i, version) = preceded(
        tuple((tag(VERSION_SIGNATURE), tag(U32_FIXED_SIZE_4))),
        tuple((le_u32, le_u16)),
    )(i)?;
    Ok((i, version))
}

fn parse_constants(i: &[u8]) -> IResult<&[u8], Option<Vec<u8>>, FormatError<&[u8]>> {
    const CONSTANTS_SIGNATURE: &[u8] = &[0x0c, 0x00];
    let (i, constants) = opt(preceded(tag(CONSTANTS_SIGNATURE), length_data(le_u32)))(i)?;
    let constants = constants.map(|slice| slice.to_vec());
    Ok((i, constants))
}

fn parse_constants_unicode(i: &[u8]) -> IResult<&[u8], Option<Vec<u8>>, FormatError<&[u8]>> {
    const CONSTANTS_UNICODE_SIGNATURE: &[u8] = &[0x3c, 0x00];
    let (i, constants_unicode) = opt(preceded(
        tag(CONSTANTS_UNICODE_SIGNATURE),
        length_data(le_u32),
    ))(i)?;
    let constants_unicode = constants_unicode.map(|slice| slice.to_vec());
    Ok((i, constants_unicode))
}

// -------------------------------------------------------------------------
// -------------------------------------------------------------------------

#[allow(clippy::type_complexity)]
fn parse_reference_name(
    i: &[u8],
    code_page: u16,
) -> IResult<&[u8], Option<String>, FormatError<&[u8]>> {
    const NAME_SIGNATURE: &[u8] = &[0x16, 0x00];
    const NAME_UNICODE_SIGNATURE: &[u8] = &[0x3e, 0x00];
    let (i, name) = opt(tuple((
        preceded(tag(NAME_SIGNATURE), length_data(le_u32)),
        preceded(tag(NAME_UNICODE_SIGNATURE), length_data(le_u32)),
    )))(i)?;
    // name_unicode MUST contain the UTF-16 encoding of name. Can be dropped without
    // loss of information.
    if let Some((name, _name_unicode)) = name {
        let name = cp_to_string(name, code_page);
        Ok((i, Some(name)))
    } else {
        Ok((i, None))
    }
}

fn parse_reference_original(
    i: &[u8],
    code_page: u16,
) -> IResult<&[u8], String, FormatError<&[u8]>> {
    const ORIGINAL_SIGNATURE: &[u8] = &[0x33, 0x00];
    let (i, libid_original) = preceded(tag(ORIGINAL_SIGNATURE), length_data(le_u32))(i)?;
    let libid_original = cp_to_string(libid_original, code_page);
    Ok((i, libid_original))
}

fn parse_reference_control(
    i: &[u8],
    code_page: u16,
) -> IResult<&[u8], ReferenceControl, FormatError<&[u8]>> {
    // REFERENCEORIGINAL Record is optional here
    let (_, id) = le_u16(i)?;
    let (i, libid_original) = match id {
        0x0033_u16 => {
            let (i, libid_original) = parse_reference_original(i, code_page)?;
            (i, Some(libid_original))
        }
        _ => (i, None),
    };

    const CONTROL_SIGNATURE: &[u8] = &[0x2f, 0x00];
    let (i, libid_twiddled) =
        preceded(tuple((tag(CONTROL_SIGNATURE), le_u32)), length_data(le_u32))(i)?;
    let libid_twiddled = cp_to_string(libid_twiddled, code_page);

    const RESERVED_1: &[u8] = &[0x00, 0x00, 0x00, 0x00];
    const RESERVED_2: &[u8] = &[0x00, 0x00];
    let (i, _) = tuple((tag(RESERVED_1), tag(RESERVED_2)))(i)?;

    let (i, name_extended) = parse_reference_name(i, code_page)?;

    const RESERVED_3: &[u8] = &[0x30, 0x00];
    let (i, libid_extended) = preceded(tuple((tag(RESERVED_3), le_u32)), length_data(le_u32))(i)?;
    let libid_extended = cp_to_string(libid_extended, code_page);

    const RESERVED_4: &[u8] = &[0x00, 0x00, 0x00, 0x00];
    const RESERVED_5: &[u8] = &[0x00, 0x00];
    let (i, _) = tuple((tag(RESERVED_4), tag(RESERVED_5)))(i)?;

    let (i, guid) = take(16_usize)(i)?;
    let guid = guid.to_vec();

    let (i, cookie) = le_u32(i)?;

    Ok((
        i,
        ReferenceControl {
            name: None,
            libid_original,
            libid_twiddled,
            name_extended,
            libid_extended,
            guid,
            cookie,
        },
    ))
}

fn parse_reference_registered(
    i: &[u8],
    code_page: u16,
) -> IResult<&[u8], ReferenceRegistered, FormatError<&[u8]>> {
    const REGISTERED_SIGNATURE: &[u8] = &[0x0d, 0x00];
    let (i, libid) = preceded(
        tuple((tag(REGISTERED_SIGNATURE), le_u32)),
        length_data(le_u32),
    )(i)?;
    let libid = cp_to_string(libid, code_page);

    const RESERVED_1: &[u8] = &[0x00, 0x00, 0x00, 0x00];
    const RESERVED_2: &[u8] = &[0x00, 0x00];
    let (i, _) = tuple((tag(RESERVED_1), tag(RESERVED_2)))(i)?;

    Ok((i, ReferenceRegistered { name: None, libid }))
}

fn parse_reference_project(
    i: &[u8],
    code_page: u16,
) -> IResult<&[u8], ReferenceProject, FormatError<&[u8]>> {
    let (i, (libid_absolute, libid_relative, major_version, minor_version)) = tuple((
        preceded(tuple((tag(&[0x0e, 0x00]), le_u32)), length_data(le_u32)),
        length_data(le_u32),
        le_u32,
        le_u16,
    ))(i)?;
    let libid_absolute = cp_to_string(libid_absolute, code_page);
    let libid_relative = cp_to_string(libid_relative, code_page);

    Ok((
        i,
        ReferenceProject {
            name: None,
            libid_absolute,
            libid_relative,
            major_version,
            minor_version,
        },
    ))
}

/// Parses a single REFERENCE Record.
///
/// There are several tricky bits to this:
/// * The first entry (NameRecord) is optional.
/// * The REFERENCE Record can be one of 4 variants.
/// * The length is implied through a terminator (0x000F) that starts a PROJECTMODULES Record.
///
/// Returns `Some(reference)` if a variant was found, `None` if the end of the array was
/// reached, or an error.
fn parse_reference(
    i: &[u8],
    code_page: u16,
) -> IResult<&[u8], Option<Reference>, FormatError<&[u8]>> {
    let (i, name) = parse_reference_name(i, code_page)?;
    // Determine REFERENCE Record variant (or end of array)
    let (_, id) = le_u16(i)?;
    match id {
        0x002f_u16 => {
            let (i, mut value) = parse_reference_control(i, code_page)?;
            value.name = name;
            Ok((i, Some(Reference::Control(value))))
        }
        0x0033_u16 => {
            let (i, libid_original) = parse_reference_original(i, code_page)?;
            let original = ReferenceOriginal {
                name,
                libid_original,
            };
            Ok((i, Some(Reference::Original(original))))
        }
        0x000d_u16 => {
            let (i, mut value) = parse_reference_registered(i, code_page)?;
            value.name = name;
            Ok((i, Some(Reference::Registered(value))))
        }
        0x000e_u16 => {
            let (i, mut value) = parse_reference_project(i, code_page)?;
            value.name = name;
            Ok((i, Some(Reference::Project(value))))
        }
        0x000f_u16 => Ok((i, None)),
        _ => Err(Error(FormatError::UnexpectedValue)),
    }
}

fn parse_references(
    i: &[u8],
    code_page: u16,
) -> IResult<&[u8], Vec<Reference>, FormatError<&[u8]>> {
    let mut result = Vec::new();
    let mut i = i;
    loop {
        let (remainder, value) = parse_reference(i, code_page)?;
        i = remainder;
        if let Some(reference) = value {
            result.push(reference);
        } else {
            return Ok((i, result));
        }
    }
}

// -------------------------------------------------------------------------
// -------------------------------------------------------------------------

fn parse_module(i: &[u8], code_page: u16) -> IResult<&[u8], Module, FormatError<&[u8]>> {
    // MODULENAME Record
    let (i, name) = preceded(tag(&[0x19, 0x00]), length_data(le_u32))(i)?;
    let name = cp_to_string(name, code_page);

    // (Optional) MODULENAMEUNICODE Record
    // If present it MUST be the UTF-16 encoding of MODULENAME. It can safely be dropped.
    let (i, _name_unicode) = opt(preceded(tag(&[0x47, 0x00]), length_data(le_u32)))(i)?;

    // MODULESTREAMNAME Record
    // stream_name_unicode MUST be the UTF-16 encoding of stream_name. It can safely be dropped.
    let (i, (stream_name, _stream_name_unicode)) = tuple((
        preceded(tag(&[0x1a, 0x00]), length_data(le_u32)),
        preceded(tag(&[0x32, 0x00]), length_data(le_u32)),
    ))(i)?;
    let stream_name = cp_to_string(stream_name, code_page);

    // MODULEDOCSTRING Record
    // doc_string_unicode MUST be the UTF-16 encoding of doc_string. It can safely be dropped.
    let (i, (doc_string, _doc_string_unicode)) = tuple((
        preceded(tag(&[0x1c, 0x00]), length_data(le_u32)),
        preceded(tag(&[0x48, 0x00]), length_data(le_u32)),
    ))(i)?;
    let doc_string = cp_to_string(doc_string, code_page);

    // MODULEOFFSET Record
    let (i, text_offset) = preceded(tuple((tag(&[0x31, 0x00]), tag(U32_FIXED_SIZE_4))), le_u32)(i)?;
    let text_offset = text_offset as _;

    // MODULEHELPCONTEXT Record
    let (i, help_context) =
        preceded(tuple((tag(&[0x1e, 0x00]), tag(U32_FIXED_SIZE_4))), le_u32)(i)?;

    // MODULECOOKIE Record
    // Cookie MUST be ignored on read.
    let (i, _cookie) = preceded(tuple((tag(&[0x2c, 0x00]), tag(U32_FIXED_SIZE_2))), le_u16)(i)?;

    // MODULETYPE Record
    let (i, id) = le_u16(i)?;
    let module_type = match id {
        0x0021_u16 => ModuleType::Procedural,
        0x0022_u16 => ModuleType::DocClsDesigner,
        _ => return Err(Error(FormatError::UnexpectedValue)),
    };
    let (i, _) = tag(&[0x00, 0x00, 0x00, 0x00])(i)?;

    // MODULEREADONLY Record
    let (i, read_only) = opt(tag(&[0x25, 0x00, 0x00, 0x00, 0x00, 0x00]))(i)?;
    let read_only = read_only.is_some();

    // MODULEPRIVATE Record
    let (i, private) = opt(tag(&[0x28, 0x00, 0x00, 0x00, 0x00, 0x00]))(i)?;
    let private = private.is_some();

    // Terminator
    let (i, _) = tag(&[0x2b, 0x00])(i)?;

    // Reserved
    let (i, _) = tag(&[0x00, 0x00, 0x00, 0x00])(i)?;

    Ok((
        i,
        Module {
            name,
            stream_name,
            doc_string,
            text_offset,
            help_context,
            module_type,
            read_only,
            private,
        },
    ))
}

fn parse_modules(i: &[u8], code_page: u16) -> IResult<&[u8], Vec<Module>, FormatError<&[u8]>> {
    let (i, count) = preceded(tuple((tag(&[0x0f, 0x00]), tag(U32_FIXED_SIZE_2))), le_u16)(i)?;
    // Cookie MUST be ignored on read.
    let (i, _cookie) = preceded(tuple((tag(&[0x13, 0x00]), tag(U32_FIXED_SIZE_2))), le_u16)(i)?;

    let mut modules = Vec::new();
    let mut i = i;
    for _ in 0..count {
        let (remainder, module) = parse_module(i, code_page)?;
        i = remainder;
        modules.push(module);
    }

    Ok((i, modules))
}

// -------------------------------------------------------------------------
// -------------------------------------------------------------------------

/// *dir* stream parser.
pub(crate) fn parse_project_information(
    i: &[u8],
) -> IResult<&[u8], ProjectInformation, FormatError<&[u8]>> {
    let (i, sys_kind) = parse_syskind(i)?;
    let (i, compat) = parse_compat(i)?;
    let (i, lcid) = parse_lcid(i)?;
    let (i, lcid_invoke) = parse_lcid_invoke(i)?;
    let (i, code_page) = parse_code_page(i)?;

    let (i, name) = parse_name(i)?;
    let name = cp_to_string(&name, code_page);

    let (i, doc_string) = parse_doc_string(i)?;
    let doc_string = cp_to_string(&doc_string, code_page);

    // doc_string_unicode MUST contain the UTF-16 encoding of doc_string. Can safely be dropped.
    let (i, _doc_string_unicode) = parse_doc_string_unicode(i)?;

    let (i, help_file_1) = parse_help_file_1(i)?;
    let help_file_1 = cp_to_string(&help_file_1, code_page);

    // help_file_2 MUST contain the same bytes as help_file_1. Can safely be dropped.
    let (i, _help_file_2) = parse_help_file_2(i)?;

    let (i, help_context) = parse_help_context(i)?;
    let (i, lib_flags) = parse_lib_flags(i)?;
    let (i, (version_major, version_minor)) = parse_version(i)?;

    // The `PROJECTCONSTANTS` record is optional (as a whole); make sure to only parse the
    // Unicode portion if `parse_constants` returned `Some`.
    //
    // TODO: Consider consolidating CP and Unicode parsing into a single function. This
    // would avoid having to subsequently deal with the outcome of this function.
    let (i, constants) = parse_constants(i)?;
    let constants = constants.map(|constants| cp_to_string(&constants, code_page));

    let i = if constants.is_some() {
        // constants_unicode MUST contain the UTF-16 encoding of constants. Can safely be
        // dropped.
        let (i, _constants_unicode) = parse_constants_unicode(i)?;
        i
    } else {
        i
    };

    let (i, references) = parse_references(i, code_page)?;

    let (i, modules) = parse_modules(i, code_page)?;

    // Terminator
    let (i, _) = tag(&[0x10, 0x00])(i)?;

    // Reserved
    let (i, _) = tag(&[0x00, 0x00, 0x00, 0x00])(i)?;

    debug_assert_eq!(i.len(), 0, "Input not fully read");

    Ok((
        i,
        ProjectInformation {
            information: Information {
                sys_kind,
                compat,
                lcid,
                lcid_invoke,
                code_page,
                name,
                doc_string,
                help_file_1,
                help_context,
                lib_flags,
                version_major,
                version_minor,
                constants,
            },
            references,
            modules,
        },
    ))
}

// -------------------------------------------------------------------------
// -------------------------------------------------------------------------

/// # Panics
///
/// This function panics, if:
/// * the passed in code page cannot be mapped to an encoding.
/// * the maximum length of the output would overflow a `usize`.
/// * part of the input could not be decoded into the allocated output `String`.
///
/// This is a temporary solution that allows me to postpone implementing error reporting
/// to a later time, when the set of expected errors and the overall error handling strategy
/// are better understood.
pub(crate) fn cp_to_string(data: &[u8], code_page: u16) -> String {
    let encoding = to_encoding(code_page).expect("Failed to map code page to an encoding.");
    let mut decoder = encoding.new_decoder_without_bom_handling();
    // The following returns `None` on overflow. That case is only expected with malformed document
    // input, so let's just panic in this case.
    let max_length = decoder.max_utf8_buffer_length(data.len()).unwrap();
    let mut result = String::with_capacity(max_length);
    let (decoder_result, _, _) = decoder.decode_to_string(data, &mut result, true);
    assert_eq!(
        decoder_result,
        CoderResult::InputEmpty,
        "Failed to decode full MBCS sequence."
    );

    result
}

/// Converts a string to a byte vector using the specified code page.
///
/// # Panics
///
/// This function panics, if:
/// * the passed in code page cannot be mapped to an encoding.
/// * the maximum length of the output would overflow a `usize`.
/// * part of the input could not be encoded into the allocated output `Vec<u8>`.
// TODO: Consider returning a Result instead of panicking.
pub(crate) fn string_to_cp(text: &str, code_page: u16) -> Vec<u8> {
    let encoding = to_encoding(code_page).expect("Failed to map code page to an encoding.");
    let mut encoder = encoding.new_encoder();
    // Estimate max length; this might be conservative for some encodings.
    let max_length = encoder.max_buffer_length_from_utf8_if_no_unmappables(text.len()).unwrap_or(text.len() * 2);
    let mut result = vec![0u8; max_length];
    let (encoder_result, _read, written, _had_errors) = encoder.encode_from_utf8(text, &mut result, true);

    match encoder_result {
        CoderResult::OutputFull => {
            // This case might happen if our initial buffer estimation was too small.
            // For simplicity now, we panic. A more robust solution would reallocate and retry.
            panic!("Output buffer too small during string_to_cp encoding. Text len: {}, buffer len: {}, written: {}", text.len(), max_length, written);
        }
        CoderResult::InputEmpty => {
            result.truncate(written);
        }
        // CoderResult::Unmappable or other errors
        _ => { // This now covers CoderResult::Unmappable and any other variants
             panic!("Failed to encode full string to code page. Result: {:?}", encoder_result);
        }
    }
    result
}

// Helper function to find the longest match for RLE compression
// Searches in `data` at `current_pos` for a sequence that matches
// a sequence in `data[0..current_pos-1]` (the history buffer).
fn find_longest_match(
    data: &[u8],
    current_pos: usize,
    min_len: usize,
    max_len: usize,       // Max length of a copy token (e.g., 18 for MS-OVBA)
    max_offset: usize,    // Max lookback offset (e.g., 4096 for MS-OVBA)
) -> (usize, usize) { // Returns (length, offset)
    let mut best_len = 0;
    let mut best_offset = 0;
    let input_len = data.len();

    if current_pos == 0 || current_pos + min_len > input_len {
        return (0, 0); // Cannot find matches if at the beginning or not enough data left for a min_len match
    }

    // Iterate over possible offsets (1 to max_offset)
    // offset_val = 1 means look at `data[current_pos - 1]`
    // The history buffer is `data[0 .. current_pos -1]`
    // The actual start of search in history: `current_pos - offset_val`
    for offset_val in 1..=std::cmp::min(max_offset, current_pos) {
        let history_match_ptr = current_pos - offset_val;

        let mut current_match_len = 0;
        for l in 0..max_len { // Potential length of match, up to MAX_COPY_LEN
            if current_pos + l >= input_len || data[history_match_ptr + l] != data[current_pos + l] {
                break; // Out of bounds for current string, or mismatch
            }
            // Ensure that history_match_ptr + l does not read past current_pos if strict non-overlapping is needed for history.
            // However, standard LZ77 allows overlapping copies (e.g. "aaaaa" -> a + copy(len=4,offset=1))
            // For MS-OVBA, overlap is generally allowed and expected.
            current_match_len += 1;
        }

        if current_match_len >= min_len {
            // If current_match_len is longer, it's a better match.
            // If current_match_len is same as best_len, prefer the one with smaller offset.
            // Since offset_val iterates from 1 (smallest offset) upwards, the first one
            // we encounter for a given length will have the smallest offset.
            // So, we only update if strictly longer.
            if current_match_len > best_len {
                best_len = current_match_len;
                best_offset = offset_val;
            }
        }
    }

    // Ensure the found length is at least min_len, otherwise it's not a valid copy token candidate
    if best_len < min_len {
        (0, 0)
    } else {
        (best_len, best_offset)
    }
}

/// Compresses data using a simple RLE algorithm as per MS-OVBA.
/// This function produces the `CompressedData` part of a `CompressedChunk`.
/// The `CompressedChunkHeader` (Size and Flag 0xB0) needs to be prepended separately.
pub(crate) fn compress(input: &[u8]) -> Result<Vec<u8>, crate::Error> {
    const MIN_COPY_LEN: usize = 3;
    const MAX_COPY_LEN: usize = 18;    // (0x0F) + 3
    // const MAX_LITERAL_LEN: usize = 128; // (0x7F) + 1 // This constant is currently unused due to LiteralToken being 1 byte
    const MAX_COPY_OFFSET: usize = 4096;

    let mut output_buffer: Vec<u8> = Vec::with_capacity(input.len());
    let mut current_pos: usize = 0;
    let input_len = input.len();

    let mut token_data_buffer: Vec<Vec<u8>> = Vec::with_capacity(8);
    let mut token_flags: u8 = 0;
    let mut token_count_in_flagbyte = 0;

    while current_pos < input_len {
        // 1. Try to find the best match (M1) starting at current_pos
        let (len1, offset1) = find_longest_match(
            input,
            current_pos,
            MIN_COPY_LEN,
            MAX_COPY_LEN,
            MAX_COPY_OFFSET,
        );

        let mut emit_copy_token = false;
        if len1 >= MIN_COPY_LEN {
            let mut len2 = 0;
            if current_pos + 1 < input_len {
                let (l2, _) = find_longest_match(
                    input,
                    current_pos + 1,
                    MIN_COPY_LEN,
                    MAX_COPY_LEN,
                    MAX_COPY_OFFSET,
                );
                len2 = l2;
            }
            if len1 >= len2 {
                emit_copy_token = true;
            }
        }

        if emit_copy_token {
            // Prepare CopyToken data
            let length_to_encode = (len1 - MIN_COPY_LEN) as u16;
            let offset_to_encode = (offset1 - 1) as u16;
            
            // Simulate bit_count calculation from decompressor
            // `current_pos` in compressor is equivalent to `result.len()` in decompressor
            // for the purpose of determining how many bytes have been (conceptually) decompressed.
            let diff = current_pos; 
            let mut bit_count = 4_usize;
            if diff > 0 {
                // This loop determines the number of bits for the offset field.
                // It stops when 2^bit_count is no longer less than diff, or bit_count reaches 16.
                while (1_usize).checked_shl(bit_count as u32).map_or(false, |val| val < diff) && bit_count < 16 {
                    bit_count += 1;
                }
            }
            // Now, bit_count holds the number of bits for the offset.
            // The remaining (16 - bit_count) bits are for the length.

            // Check if the values fit into the allocated bits.
            // Max length_to_encode (18-3 = 15) needs up to 4 bits (fits in 16-bit_count if bit_count <= 12).
            // Max offset_to_encode (4096-1 = 4095) needs up to 12 bits (fits in bit_count if bit_count >= 12).
            let max_val_for_length_field = if (16 - bit_count) == 0 { 0 } else { (1u16 << (16 - bit_count)) - 1 };
            let max_val_for_offset_field = if bit_count == 0 { 0 } else { (1u16 << bit_count) - 1 };

            let token_data_word: u16;
            if length_to_encode > max_val_for_length_field || offset_to_encode > max_val_for_offset_field {
                // This condition means the current (len1, offset1) from find_longest_match
                // cannot be encoded with the bit allocation determined by `bit_count`.
                // For example, if bit_count is 4 (expecting small offset), but offset1 is large.
                // Or if bit_count is 15 (expecting small length), but len1 is large.
                // This suggests a potential issue: either find_longest_match should be constrained
                // by what can be encoded, or this token should not be a CopyToken.
                // For now, using the fixed 12/4 encoding as a fallback is a HACK.
                // A better solution would be to re-evaluate emitting a CopyToken here, possibly emitting literals.
                // eprintln!(\
                //     "Warning: Incompatible match (off:{}, len:{}) for dynamic bit_count:{}. LMax:{}, OMax:{}. Falling back to 12/4.",\
                //     offset1, len1, bit_count, max_val_for_length_field, max_val_for_offset_field\
                // );\n                token_data_word = (offset_to_encode << 4) | length_to_encode; // Fallback to fixed 12/4 format\n            } else {\n                // Encode with dynamic bit_count:\n                // Offset uses `bit_count` most significant bits.\n                // Length uses `16 - bit_count` least significant bits.\n                token_data_word = (offset_to_encode << (16 - bit_count)) | length_to_encode;\n            }\n            \n            token_data_buffer.push(token_data_word.to_le_bytes().to_vec());\n            token_flags |= 1 << token_count_in_flagbyte; // Set bit for CopyToken\n            current_pos += len1;\n        } else {\n            // Not a copy token, so emit a single LiteralToken\n            if current_pos < input_len { // Убедимся, что есть что читать\n                token_data_buffer.push(vec![input[current_pos]]); // Данные LiteralToken - это сам байт\n                // token_flags бит для этой позиции уже 0 (по умолчанию при инициализации token_flags = 0)\n                current_pos += 1; // Продвигаемся на 1 байт\n            } else {\n                // This case should ideally not be reached due to the outer while loop condition\n                break; \n            }\n            \n            token_count_in_flagbyte += 1;\n\n            if token_count_in_flagbyte == 8 || current_pos >= input_len {\n                output_buffer.push(token_flags);\n                for data_vec in &token_data_buffer {\n                    output_buffer.extend_from_slice(data_vec);\n                }\n                token_data_buffer.clear();\n                token_flags = 0;\n                token_count_in_flagbyte = 0;\n            }\n        }\n    }\n    Ok(output_buffer)
                token_data_word = (offset_to_encode << 4) | length_to_encode; // Fallback to fixed 12/4 format
            } else {
                // Encode with dynamic bit_count:
                // Offset uses `bit_count` most significant bits.
                // Length uses `16 - bit_count` least significant bits.
                token_data_word = (offset_to_encode << (16 - bit_count)) | length_to_encode;
            }
            
            token_data_buffer.push(token_data_word.to_le_bytes().to_vec());
            token_flags |= 1 << token_count_in_flagbyte; // Set bit for CopyToken
            current_pos += len1;
        } else {
            // Not a copy token, so emit a single LiteralToken
            if current_pos < input_len { // Убедимся, что есть что читать
                token_data_buffer.push(vec![input[current_pos]]); // Данные LiteralToken - это сам байт
                // token_flags бит для этой позиции уже 0 (по умолчанию при инициализации token_flags = 0)
                current_pos += 1; // Продвигаемся на 1 байт
            } else {
                // This case should ideally not be reached due to the outer while loop condition
                break; 
            }
        }
        
        token_count_in_flagbyte += 1;

        if token_count_in_flagbyte == 8 || current_pos >= input_len {
            output_buffer.push(token_flags);
            for data_vec in &token_data_buffer {
                output_buffer.extend_from_slice(data_vec);
            }
            token_data_buffer.clear();
            token_flags = 0;
            token_count_in_flagbyte = 0;
        }
    }
    Ok(output_buffer)
}

