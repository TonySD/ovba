//! An Office VBA project parser written in 100% safe Rust.
//!
//! This is a (partial) implementation of the [\[MS-OVBA\]: Office VBA File Format
//! Structure][MS-OVBA] protocol (Revision 9.1, published 2020-02-19).
//!
//! The main entry point into the API is the [`Project`] type, returned by the
//! [`open_project`] function.
//!
//! # Usage
//!
//! Opening a project:
//!
//! ```rust,no_run
//! use std::fs::read;
//! use ovba::open_project;
//!
//! let data = read("vbaProject.bin")?;
//! let project = open_project(data)?;
//! # Ok::<(), ovba::Error>(())
//! ```
//!
//! A more complete example that dumps an entire VBA project's source code:
//!
//! ```rust,no_run
//! use std::fs::{read, write};
//! use ovba::open_project;
//!
//! let data = read("vbaProject.bin")?;
//! let project = open_project(data)?;
//!
//! for module in &project.modules {
//!     let src_code = project.module_source_raw(&module.name)?;
//!     write("./out/".to_string() + &module.name, src_code)?;
//! }
//! # Ok::<(), ovba::Error>(())
//! ```
//!
//! The API also supports low-level access to the [\[MS-CFB\]: Compound File Binary File
//! Format][MS-CFB] data. The following example lists all CFB entries:
//!
//! ```rust,no_run
//! use std::fs::read;
//! use ovba::open_project;
//!
//! let data = read("vbaProject.bin")?;
//! let project = open_project(data)?;
//! for (name, path) in &project.list()? {
//!     println!(r#"Name: "{}"; Path: "{}""#, name, path);
//! }
//! # Ok::<(), ovba::Error>(())
//! ```
//!
//! [MS-OVBA]: https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc
//! [MS-CFB]: https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b

#![forbid(unsafe_code)]
#![warn(rust_2018_idioms, missing_docs)]

mod error;
pub use crate::error::{Error, Result};

mod parser;

use cfb::CompoundFile;
use parser::cp_to_string;
use std::collections::HashMap;
use std::convert::TryInto;
use std::fs;

use std::{
    cell::RefCell,
    io::{Cursor, Read, Write},
    path::{Path, PathBuf},
    str::FromStr,
};

/// Represents a VBA project.
///
/// This type serves as the entry point into this crate's functionality and exposes the
/// public API surface.
pub struct Project {
    /// Specifies version-independent information for the VBA project.
    pub information: Information,
    /// Specifies the external references of the VBA project.
    pub references: Vec<Reference>,
    /// Specifies the modules in the project.
    pub modules: Vec<Module>,
    container: RefCell<CompoundFile<Cursor<Vec<u8>>>>,
    /// The original raw data of the VBA project file (e.g., vbaProject.bin).
    /// This can be used for operations that require the untouched source or for full rebuilds.
    pub original_raw_data: Vec<u8>,
}

/// Specifies the platform for which the VBA project is created.
#[derive(Debug)]
pub enum SysKind {
    /// For 16-bit Windows Platforms.
    Win16,
    /// For 32-bit Windows Platforms.
    Win32,
    /// For Macintosh Platforms.
    MacOs,
    /// For 64-bit Windows Platforms.
    Win64,
}

// TODO: Remove exemption once the implementation is complete.
#[allow(dead_code)]
/// Specifies a reference to a twiddled type library and its extended type library.
#[derive(Debug)]
pub struct ReferenceControl {
    /// (Optional) Name entry
    name: Option<String>,
    libid_original: Option<String>,
    libid_twiddled: String,
    name_extended: Option<String>,
    libid_extended: String,
    guid: Vec<u8>, // Should be an `[u8; 16]`, though I'm not sure how to convert &[u8] returned by the parser into an array.
    /// MUST be Unique for each `ReferenceControl` in the VBA projectwith the same
    /// libid_original.
    cookie: u32,
}

// TODO: Remove exemption once the implementation is complete.
#[allow(dead_code)]
/// Specifies the identifier of the Automation type library the containing
/// [`ReferenceControl`]'s twiddled type library was generated from.
#[derive(Debug)]
pub struct ReferenceOriginal {
    /// (Optional) Name entry
    name: Option<String>,
    libid_original: String,
}

// TODO: Remove exemption once the implementation is complete.
#[allow(dead_code)]
/// Specifies a reference to an Automation type library.
#[derive(Debug)]
pub struct ReferenceRegistered {
    name: Option<String>,
    libid: String,
}

// TODO: Remove exemption once the implementation is complete.
#[allow(dead_code)]
/// Specifies a reference to an external VBA project.
#[derive(Debug)]
pub struct ReferenceProject {
    name: Option<String>,
    libid_absolute: String,
    libid_relative: String,
    major_version: u32,
    minor_version: u16,
}

/// Specifies a reference to an Automation type library or VBA project.
#[derive(Debug)]
pub enum Reference {
    /// The `Reference` is a [`ReferenceControl`].
    Control(ReferenceControl),
    /// The `Reference` is a [`ReferenceOriginal`].
    Original(ReferenceOriginal),
    /// The `Reference` is a [`ReferenceRegistered`].
    Registered(ReferenceRegistered),
    /// The `Reference` is a [`ReferenceProject`].
    Project(ReferenceProject),
}

// TODO: Remove exemption once the implementation is complete.
#[allow(dead_code)]
/// Specifies version-independent information for the VBA project.
#[derive(Debug)]
pub struct Information {
    /// Specifies the platform for which the VBA project is created.
    pub sys_kind: SysKind,
    compat: Option<u32>,
    lcid: u32,
    lcid_invoke: u32,
    /// Specifies the code page for the VBA project.
    ///
    pub code_page: u16,
    name: String,
    doc_string: String,
    help_file_1: String,
    help_context: u32,
    lib_flags: u32,
    version_major: u32,
    version_minor: u16,
    constants: Option<String>,
}

/// Specifies the containing module's type.
#[derive(Debug)]
pub enum ModuleType {
    /// Specifies a procedural module.
    ///
    /// A procedural module is a collection of subroutines and functions.
    Procedural,
    /// Specifies a document module, class module, or designer module.
    ///
    /// A document module is a type of VBA project item that specifies a module for
    /// embedded macros and programmatic access operations that are associated with a
    /// document.
    ///
    /// A class module is a module that contains the definition for a new object. Each
    /// instance of a class creates a new object, and procedures that are defined in the
    /// module become properties and methods of the object.
    ///
    /// A designer module is a VBA module that extends the methods and properties of an
    /// ActiveX control that has been registered with the project.
    ///
    /// The file format specification doesn't distinguish between these three module
    /// types and encodes them using a single umbrella type ID.
    DocClsDesigner,
}

/// Specifies data for a module.
#[derive(Debug)]
pub struct Module {
    /// Specifies a VBA identifier as the name of the containing `Module`.
    pub name: String,
    /// Specifies the stream name in the VBA storage corresponding to the containing
    /// `Module`.
    pub stream_name: String,
    /// Specifies the description for the containing `Module`.
    pub doc_string: String,
    /// Specifies the location of the source code within the stream that corresponds to
    /// the containing `Module`.
    pub text_offset: usize,
    /// Specifies the Help topic identifier for the containing `Module`.
    pub help_context: u32,
    /// Specifies whether the containing `Module` is a procedural module, document
    /// module, class module, or designer module.
    pub module_type: ModuleType,
    /// Specifies that the containing `Module` is read-only.
    pub read_only: bool,
    /// Specifies that the containing `Module` is only usable from within the current VBA
    /// project.
    pub private: bool,
}

impl Project {
    /// Returns a stream's decompressed data.
    ///
    /// This function reads a stream referenced by `stream_path` and passes the data
    /// starting at `offset` into the RLE decompressor.
    ///
    /// The primary use case for this function is to extract source code from VBA
    /// [`Module`]s. The respective `offset` is reported by [`Module::text_offset`].
    ///
    /// This is a low-level function that is useful for very specific use cases only.
    /// Client code that needs to read source code should use [`Project::module_source`]
    /// or [`Project::module_source_raw`] instead.
    // TODO: Code example
    pub fn decompress_stream_from<P>(&self, stream_path: P, offset: usize) -> Result<Vec<u8>>
    where
        P: AsRef<Path> + std::fmt::Debug,
    {
        println!("[DEBUG decompress_stream_from] Чтение потока: {:?}, offset: {}", stream_path, offset);
        let data = self.read_stream(stream_path.as_ref())?;
        println!("[DEBUG decompress_stream_from] Прочитано {} байт из потока. Данные (первые ~20 байт после offset, если есть): {:?}", data.len(), data.get(offset..std::cmp::min(data.len(), offset + 20)));

        if offset > data.len() {
            eprintln!("[ERROR decompress_stream_from] offset ({}) > длина данных ({})", offset, data.len());
            return Err(Error::Decompressor); 
        }
        
        if data.get(offset..).map_or(true, |s| s.is_empty()) {
            eprintln!("[ERROR decompress_stream_from] Срез данных для распаковки пуст (offset: {}, data.len(): {})", offset, data.len());
            return Err(Error::Decompressor);
        }

        if data.get(offset) != Some(&0x01) {
             println!("[DEBUG decompress_stream_from] Ожидался SigByte 0x01 по смещению {}, но найден {:X?}", offset, data.get(offset));
        }

        match parser::decompress(&data[offset..]) { 
            Ok((remainder, decompressed_data)) => {
                println!("[DEBUG decompress_stream_from] Распаковано {} байт. Остаток после распаковки (nom): {} байт.", decompressed_data.len(), remainder.len());
                if !remainder.is_empty() {
                    println!("[DEBUG decompress_stream_from] Внимание: после распаковки остался необработанный остаток {} байт: {:?}", remainder.len(), remainder.get(..std::cmp::min(remainder.len(), 20)));
                }
                Ok(decompressed_data)
            }
            Err(e) => {
                eprintln!("[ERROR decompress_stream_from] Ошибка nom при распаковке: {:?}", e);
                Err(Error::Decompressor)
            }
        }
    }

    // TODO: This should probably live someplace else. It exposes information internal to
    //       the CFB implementation, that's not *immediately* useful or related to this
    //       library's primary responsibility.

    /// Returns a list of entries (storages and streams) in the raw binary data. Each
    /// entry is represented as a tuple of two `String`s, where the first element
    /// contains the entry's name and the second element the entry's path inside the
    /// CFB.
    ///
    /// The raw binary data is encoded as a [Compound File Binary][MS-CFB]
    ///
    /// [MS-CFB]: https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b
    pub fn list(&self) -> Result<Vec<(String, String)>> {
        let mut result = Vec::new();
        for entry in self
            .container
            .borrow()
            .walk_storage("/")
            .map_err(Error::Cfb)?
        {
            result.push((
                entry.name().to_owned(),
                entry.path().to_str().unwrap_or_default().to_owned(),
            ));
        }
        Ok(result)
    }

    /// Returns a module's source code.
    ///
    /// Similar to [`Project::module_source_raw`] this function returns the source code
    /// of a project's module. After the raw source code has been decoded it is then
    /// converted to a `String` using the project's code page.
    pub fn module_source(&self, name: &str) -> Result<String> {
        let source_raw = self.module_source_raw(name)?;
        let source = cp_to_string(&source_raw, self.information.code_page);

        Ok(source)
    }

    /// Returns the raw source code from a module.
    ///
    /// The result contains a module's source code as is. No character encoding conversion
    /// is done. The data is encoded using the project's code page available through
    /// [`Information::code_page`].
    pub fn module_source_raw(&self, name: &str) -> Result<Vec<u8>> {
        let module = self
            .modules
            .iter()
            .find(|&module| module.name == name)
            .ok_or_else(|| Error::ModuleNotFound(name.to_owned()))?;

        // `PathBuf::from_str` cannot fail (`type Err = Infallible`). The subsequent
        // `unwrap` thus won't `panic!`. No path separator normalization is done in the
        // process; this is intentional.
        let path = PathBuf::from_str("/VBA").unwrap().join(&module.stream_name);
        let offset = module.text_offset;
        let src_code = self.decompress_stream_from(path, offset)?;

        Ok(src_code)
    }

    /// Returns a stream's contents.
    ///
    /// This is a low-level function operating on the CFB data. The CFB is the storage
    /// container of the raw binary VBA project.
    pub fn read_stream<P>(&self, stream_path: P) -> Result<Vec<u8>>
    where
        P: AsRef<Path>,
    {
        let mut buffer = Vec::new();
        self.container
            .borrow_mut()
            .open_stream(stream_path)
            .map_err(Error::Cfb)?
            .read_to_end(&mut buffer)
            .map_err(Error::Cfb)?;

        Ok(buffer)
    }

    /// Sets the source code for a given module.
    ///
    /// This function attempts an "in-place" modification of the CFB container.
    /// It modifies the `container` field directly.
    /// The `dir` stream is NOT updated by this function, which might lead to
    /// inconsistencies if module sizes/counts change significantly.
    /// Call `save()` to get the modified project bytes.
    pub fn set_module_source(&mut self, module_name: &str, new_source: &str) -> Result<()> {
        println!("[DEBUG set_module_source] Начало для модуля: {}", module_name);
        let module = self
            .modules
            .iter()
            .find(|m| m.name == module_name)
            .ok_or_else(|| Error::ModuleNotFound(module_name.to_owned()))?;
        println!("[DEBUG set_module_source] Модуль найден: {:?}, text_offset: {}", module.stream_name, module.text_offset);

        let stream_path = PathBuf::from("/VBA").join(&module.stream_name);

        // 1. Прочитать оригинальный поток модуля, чтобы получить PerformanceCache
        let original_module_stream_data = self.read_stream(&stream_path)?;
        println!("[DEBUG set_module_source] Размер оригинального потока модуля: {}", original_module_stream_data.len());

        if module.text_offset > original_module_stream_data.len() {
            eprintln!("[ERROR set_module_source] text_offset ({}) > длина потока ({})", module.text_offset, original_module_stream_data.len());
            return Err(Error::Generic(
                "Module text_offset exceeds stream data length".to_string(),
            ));
        }
        let performance_cache = &original_module_stream_data[..module.text_offset];
        println!("[DEBUG set_module_source] Размер PerformanceCache: {}", performance_cache.len());

        // 2. Преобразовать новый исходный код в байты
        println!("[DEBUG set_module_source] Новый исходный код (первые 50 символов): {:.50}", new_source.replace('\n', "\\n"));
        let new_source_bytes = parser::string_to_cp(new_source, self.information.code_page);
        println!("[DEBUG set_module_source] Размер нового исходного кода в байтах (до сжатия): {}", new_source_bytes.len());

        // 3. Сжать новые байты в CompressedData (это результат parser::compress, он включает FlagByte и TokenData)
        let compressed_module_code_data = parser::compress(&new_source_bytes)?;
        println!("[DEBUG set_module_source] Размер compressed_module_code_data (FlagByte + TokenData): {}", compressed_module_code_data.len());

        // 4. Сформировать CompressedChunkHeader (u16)
        // CompressedChunkHeader: [F S S S] [S S S S] [S S S S S S S S]
        // F: бит 15 (CompressedFlag: 1 if compressed, 0 if raw)
        // SSS: биты 12-14 (Signature: 0b011)
        // SSSSSSSSSSSSS: биты 0-11 (Size: actual_compressed_data_length_in_chunk - 1)
        
        let actual_compressed_data_length_in_chunk = compressed_module_code_data.len();
        if actual_compressed_data_length_in_chunk == 0 {
            // Не должно быть 0, т.к. parser::compress должен вернуть хотя бы FlagByte
            // или если вход пустой, то это особый случай (но VBA модули редко пустые)
             println!("[WARN set_module_source] compressed_module_code_data имеет нулевую длину. Это неожиданно.");
             // Если это возможно, нужно решить, какой заголовок ставить. Пока что оставим ошибку.
             return Err(Error::Generic("Compressed data part is empty, cannot form header".to_string()));
        }
        if actual_compressed_data_length_in_chunk > 0xFFF + 1 { // 0xFFF (4095) + 1 = 4096
            return Err(Error::Generic(
                format!("Compressed data part too large ({}) for chunk size field (max 4096)", actual_compressed_data_length_in_chunk)
            ));
        }
        let size_for_header_field = (actual_compressed_data_length_in_chunk - 1) as u16; // Max 0xFFF (4095)

        const COMPRESSED_FLAG_BIT: u16 = 1 << 15; // 1 для сжатого
        const SIGNATURE_BITS: u16 = 0b011 << 12;

        let chunk_header: u16 = COMPRESSED_FLAG_BIT | SIGNATURE_BITS | size_for_header_field;
        
        println!("[DEBUG set_module_source] CompressedChunkHeader: Value=0x{:04X} (Flag:1, Sig:0b011, SizeField:{})", 
            chunk_header, size_for_header_field
        );

        // 5. Собрать новое содержимое потока модуля: PerformanceCache + SigByte (для всего контейнера) + CompressedChunkHeader + CompressedData
        const SIG_BYTE_CONTAINER: u8 = 0x01;
        let mut new_module_stream_content = Vec::with_capacity(
            performance_cache.len() + 1 + 2 + compressed_module_code_data.len() // PerfCache + SigByteContainer + ChunkHeader(u16) + CompressedData
        );
        new_module_stream_content.extend_from_slice(performance_cache);
        new_module_stream_content.push(SIG_BYTE_CONTAINER); 
        new_module_stream_content.extend_from_slice(&chunk_header.to_le_bytes()); // CompressedChunkHeader (2 байта)
        new_module_stream_content.extend_from_slice(&compressed_module_code_data);    // CompressedData (FlagByte + TokenData)
        
        println!("[DEBUG set_module_source] Общий размер new_module_stream_content для записи в поток: {}", new_module_stream_content.len());

        // 6. Заменить поток в CFB контейнере
        println!("[DEBUG set_module_source] Попытка записи в CFB поток: {:?}", stream_path);
        let mut cfb = self.container.borrow_mut();

        if cfb.exists(&stream_path) {
            println!("[DEBUG set_module_source] Удаление существующего потока: {:?}", stream_path);
            cfb.remove_stream(&stream_path).map_err(Error::Cfb)?;
        } else {
            println!("[DEBUG set_module_source] Поток {:?} не существует, будет создан новый.", stream_path);
        }

        let mut stream_writer = cfb.create_stream(&stream_path).map_err(Error::Cfb)?;
        println!("[DEBUG set_module_source] Поток создан, запись {} байт...", new_module_stream_content.len());
        stream_writer
            .write_all(&new_module_stream_content)
            .map_err(Error::Io)?;
        println!("[DEBUG set_module_source] Запись в поток завершена.");
        
        println!("[DEBUG set_module_source] Завершение.");
        Ok(())
    }

    /// Saves the project to the specified file path.
    /// After this operation, the internal CFB container of this `Project` instance
    /// will be consumed and no longer usable for further CFB operations.
    pub fn save(&mut self, output_path: &Path) -> Result<()> { 
        // It is crucial to manage the borrow of self.container carefully.
        // First, we borrow mutably to flush.
        let mut cfb_guard = self.container.borrow_mut();
        cfb_guard.flush().map_err(Error::Cfb)?;
        // Then, we must drop the guard *before* calling `self.container.replace()`,
        // as `replace()` also needs to borrow mutably.
        drop(cfb_guard);

        // To move the `CompoundFile` out of the `RefCell`, we replace it with a dummy.
        // `CompoundFile::create` will give us an empty but valid CFB in memory.
        let dummy_cfb = match CompoundFile::create(Cursor::new(Vec::new())) {
            Ok(cfb) => cfb,
            Err(e) => {
                // This should ideally not happen with a Vec-backed cursor for create.
                // If it does, it's a more fundamental issue with the cfb crate or environment.
                eprintln!("[ERROR save] Failed to create dummy CompoundFile: {:?}", e);
                return Err(Error::Cfb(e)); // Convert cfb::Error to our Error::Cfb
            }
        };
        let original_compound_file = self.container.replace(dummy_cfb);

        // `original_compound_file.into_inner()` should directly return `Cursor<Vec<u8>>` (the W type).
        // If there were an error here (e.g., if W was a file and couldn't be finalized),
        // `cfb` crate design implies it would panic or handle error during flush/write, not in `into_inner()` for `Cursor<Vec<u8>>`.
        let cursor: Cursor<Vec<u8>> = original_compound_file.into_inner(); 
        
        let updated_project_bytes = cursor.into_inner();

        println!("[DEBUG save] Обновленный размер проекта для записи: {} байт", updated_project_bytes.len());

        fs::write(output_path, updated_project_bytes).map_err(Error::Io)?;
        
        println!("[DEBUG save] Проект успешно сохранен в: {:?}", output_path);
        Ok(())
    }
}

/// Opens a VBA project.
///
/// This function consumes `raw` and returns a [`Project`] struct on success, populated
/// with data from the parsed binary input.
pub fn open_project(raw: Vec<u8>) -> Result<Project> {
    // `raw` будет использован для создания Cursor для CompoundFile.
    // Также сохраним копию `raw` в структуре Project для возможной полной пересборки
    // или если понадобится оригинальный, нетронутый файл.
    let original_data_clone = raw.clone(); // Клонируем для хранения в Project
    let data_for_reading_cursor = Cursor::new(raw); // `raw` перемещается сюда

    let mut cfb_container = CompoundFile::open(data_for_reading_cursor).map_err(Error::Cfb)?;

    // Read *dir* stream
    #[cfg(target_family = "windows")]
    const DIR_STREAM_PATH: &str = "/VBA\\dir";
    #[cfg(not(target_family = "windows"))]
    const DIR_STREAM_PATH: &str = "/VBA/dir";

    let mut buffer = Vec::new();
    cfb_container
        .open_stream(DIR_STREAM_PATH)
        .map_err(Error::Cfb)?
        .read_to_end(&mut buffer)
        .map_err(Error::Cfb)?;

    // Decompress stream
    let (_remainder_after_decompress, decompressed_buffer) =
        match parser::decompress(&buffer) {
            Ok((remainder, buffer_content)) => (remainder, buffer_content),
            Err(_nom_err) => {
                // Можно добавить логирование _nom_err здесь, если необходимо
                // eprintln!("Decompression error: {:?}", _nom_err);
                return Err(Error::Decompressor);
            }
        };
    // Убедимся, что весь буфер был использован, если это важно для логики парсера.
    // Оригинальный код использовал debug_assert!(_remainder_after_decompress.is_empty());
    // Если parser::decompress должен всегда потреблять весь ввод или это ошибка,
    // то здесь можно добавить проверку if !_remainder_after_decompress.is_empty() { return Err(Error::Decompressor); }

    // Parse binary data
    let (_remainder_after_parse, project_info_data) =
        match parser::parse_project_information(&decompressed_buffer) {
            Ok((remainder, info_content)) => (remainder, info_content),
            Err(_nom_err) => {
                // Можно добавить логирование _nom_err здесь
                // eprintln!("Parsing error: {:?}", _nom_err);
                return Err(Error::Parser);
            }
        };
    // Аналогично, проверка остатка, если необходимо.
    // if !_remainder_after_parse.is_empty() { return Err(Error::Parser); }

    Ok(Project {
        information: project_info_data.information,
        references: project_info_data.references,
        modules: project_info_data.modules,
        container: RefCell::new(cfb_container),
        original_raw_data: original_data_clone,
    })
}

#[cfg(test)]
mod main_test {
    use super::*; // Импорт всего из lib.rs
    use std::fs;
    use std::path::Path;

    fn run_vba_modification_simulation(project_path: &Path, module_to_modify: &str) -> Result<()> {
        println!("Загрузка VBA проекта из: {:?}", project_path);
        let initial_data = fs::read(project_path).map_err(Error::Io)?;
        
        let mut project = open_project(initial_data.clone())?; 
        println!("Проект успешно открыт.");

        println!("\nЧтение исходного кода модуля '{}' (до изменения):", module_to_modify);
        match project.module_source(module_to_modify) {
            Ok(source) => {
                println!("------------------------------------");
                println!("{}", source);
                println!("------------------------------------");
            }
            Err(e) => {
                eprintln!("Ошибка чтения исходного кода модуля '{}': {:?}", module_to_modify, e);
                if project.modules.is_empty() {
                    println!("Распарсенных модулей в проекте не найдено.");
                } else {
                    println!("\nРаспарсенные модули в проекте:");
                    for m in &project.modules {
                        println!(" - Имя: {}, Поток: {}", m.name, m.stream_name);
                    }
                }
                return Err(e);
            }
        }

        let original_source = project.module_source(module_to_modify)?;
        let modified_source = format!("{}\n' This is a test\n", original_source);
        println!("\nИзменение исходного кода модуля '{}'...", module_to_modify);
        project.set_module_source(module_to_modify, &modified_source)?;
        println!("Исходный код модуля '{}' обновлен в памяти.", module_to_modify);

        println!("\nСохранение изменений в файл...");
        let output_file_path = project_path.with_file_name(format!("{}_patched.bin", project_path.file_stem().unwrap_or_default().to_string_lossy()));
        project.save(&output_file_path)?;
        println!("Изменения сохранены в файл: {:?}", output_file_path);
        println!("Оригинальный экземпляр 'project' больше не должен использоваться для операций с CFB.");

        // ТЕПЕРЬ ПРОВЕРЯЕМ ЗАГРУЗКОЙ ИЗ СОХРАНЕННОГО ФАЙЛА
        println!("\nЧтение и проверка сохраненного файла: {:?}", output_file_path);
        let patched_data = fs::read(&output_file_path).map_err(Error::Io)?;
        let patched_project = open_project(patched_data)?; // Загружаем в новый экземпляр

        println!("\nЧтение исходного кода модуля '{}' из перезагруженного проекта:", module_to_modify);
        match patched_project.module_source(module_to_modify) {
            Ok(source) => {
                println!("------------------------------------");
                println!("{}", source);
                println!("------------------------------------");

                if source.contains("' This is a test") {
                    println!("\nУСПЕХ: Изменения найдены в сохраненном и перезагруженном файле!");
                } else {
                    eprintln!("\nОШИБКА: Изменения НЕ найдены в сохраненном и перезагруженном файле!");
                    return Err(Error::Generic("Verification failed: Changes not found in reloaded file".to_string())); // Возвращаем ошибку, чтобы тест упал
                }
            }
            Err(e) => {
                eprintln!("Ошибка чтения исходного кода модуля из сохраненного файла: {:?}", e);
                return Err(e);
            }
        }
        
        Ok(())
    }

    #[test]
    fn test_vba_modification() {
        // !!! ЗАМЕНИТЕ ЭТИ ЗНАЧЕНИЯ !!!
        let project_file_path_str = "/tmp/tmp.wnwowFBJ8k/ovba/test/vbaProject.bin"; // Например, "tests/test-files/vbaProject.bin"
        let module_name = "NewMacros"; // Например, "Module1" или имя существующего модуля
        // !!! КОНЕЦ ЗАМЕНЫ !!!

        let project_file_path = Path::new(project_file_path_str);

        if !project_file_path.exists() {
            eprintln!("Тестовый файл VBA проекта {:?} не найден. Тест будет проигнорирован.", project_file_path);
            // Чтобы тест не падал, а просто игнорировался, если файла нет, можно сделать так:
            //eprintln!("Пожалуйста, создайте файл или укажите правильный путь.");
            //return; // Игнорировать тест, если файла нет.
            // Или паниковать, чтобы CI видел проблему:
            panic!("Тестовый файл vbaProject.bin не найден по пути: {:?}. Пожалуйста, создайте его или исправьте путь.", project_file_path);
        }

        match run_vba_modification_simulation(project_file_path, module_name) {
            Ok(_) => println!("\nСимуляция модификации VBA успешно завершена для модуля '{}'.", module_name),
            Err(e) => {
                eprintln!("\nОшибка в симуляции модификации VBA для модуля '{}': {:?}", module_name, e);
                panic!("Тест модификации VBA провален: {:?}", e); // Падение теста при ошибке
            }
        }
    }
}

