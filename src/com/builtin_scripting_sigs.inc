// ============================================================
// 本文件是**生成物** —— 不要手改。
//
// 用途: TypeLibParser::registerBuiltinScriptingTypelib() 的内建签名表。
//       内建兜底原本只登类名（默认接口是空的），于是没有 scrrun 类型库的机器上成员
//       只能运行期 by-name；这张表把本机真库（scrrun.dll）的成员签名原样搬过来。
//
// 来源: `C3.exe <任一.bas> --dump-com-sigs Scripting` 的输出，
//       再由 scripts\gen_builtin_com_sigs.ps1 机械转换（重新生成的办法写在那个脚本头部）。
//       生成时间: 
//       源库: 
//
// 已剔除: IUnknown/IDispatch 那 7 个继承成员（按名字调用不到）、以及本运行期不实现的
//         Encoder coclass。clsidStr 也**故意**留空 —— 内建表不该进 ComLib=/Object=
//         那几条免注册表路（见函数内注释）。
//
// 验收: scripts\check_builtin_com_sigs.ps1 —— 同一份探针在「默认」与
//       C3_TEST_NO_SCRRUN_TLB=1 两种编译下，发出的 C 必须逐字节相同。
// ============================================================

struct BuiltinParam { const char* name; Vb6Type type; ComParamDir dir; bool opt; };
struct BuiltinMember { const char* real; ComMemberKind kind; Vb6Type ret; int32_t memid; int32_t vtbl; int nparams; const BuiltinParam* params; };
struct BuiltinClass { const char* cls; const char* progId; const char* iface; const char* iid; bool dual; bool dispatch; const char* dfltName; const char* dfltReal; int nmembers; const BuiltinMember* members; };

static const BuiltinParam kBP_Dictionary_0[] = {
    { "Key", Vb6Type::Object, ComParamDir::In, false },
    { "p1", Vb6Type::Object, ComParamDir::In, false },
    { "Key", Vb6Type::Object, ComParamDir::In, false },
    { "p1", Vb6Type::Object, ComParamDir::In, false },
    { "Key", Vb6Type::Object, ComParamDir::In, false },
};
static const BuiltinParam kBP_Dictionary_1[] = {
    { "Key", Vb6Type::Object, ComParamDir::In, false },
    { "Item", Vb6Type::Object, ComParamDir::In, false },
};
static const BuiltinParam kBP_Dictionary_3[] = {
    { "Key", Vb6Type::Object, ComParamDir::In, false },
};
static const BuiltinParam kBP_Dictionary_5[] = {
    { "Key", Vb6Type::Object, ComParamDir::In, false },
    { "p1", Vb6Type::Object, ComParamDir::In, false },
};
static const BuiltinParam kBP_Dictionary_7[] = {
    { "Key", Vb6Type::Object, ComParamDir::In, false },
};
static const BuiltinParam kBP_Dictionary_9[] = {
    { "p0", Vb6Type::Long, ComParamDir::In, false },
};
static const BuiltinParam kBP_Dictionary_11[] = {
    { "Key", Vb6Type::Object, ComParamDir::In, false },
};
static const BuiltinMember kBM_Dictionary[] = {
    { "Item", ComMemberKind::PropertyPutRef, Vb6Type::Empty, 0, 7, 5, kBP_Dictionary_0 },
    { "Add", ComMemberKind::Method, Vb6Type::Empty, 1, 10, 2, kBP_Dictionary_1 },
    { "Count", ComMemberKind::PropertyGet, Vb6Type::Long, 2, 11, 0, nullptr },
    { "Exists", ComMemberKind::Method, Vb6Type::Boolean, 3, 12, 1, kBP_Dictionary_3 },
    { "Items", ComMemberKind::Method, Vb6Type::Variant, 4, 13, 0, nullptr },
    { "Key", ComMemberKind::PropertyPut, Vb6Type::Empty, 5, 14, 2, kBP_Dictionary_5 },
    { "Keys", ComMemberKind::Method, Vb6Type::Variant, 6, 15, 0, nullptr },
    { "Remove", ComMemberKind::Method, Vb6Type::Empty, 7, 16, 1, kBP_Dictionary_7 },
    { "RemoveAll", ComMemberKind::Method, Vb6Type::Empty, 8, 17, 0, nullptr },
    { "CompareMode", ComMemberKind::PropertyPut, Vb6Type::Empty, 9, 18, 1, kBP_Dictionary_9 },
    { "_NewEnum", ComMemberKind::Method, Vb6Type::Object, -4, 20, 0, nullptr },
    { "HashVal", ComMemberKind::PropertyGet, Vb6Type::Variant, 10, 21, 1, kBP_Dictionary_11 },
};

static const BuiltinParam kBP_FileSystemObject_1[] = {
    { "Path", Vb6Type::String, ComParamDir::In, false },
    { "Name", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_2[] = {
    { "Path", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_3[] = {
    { "Path", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_4[] = {
    { "Path", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_5[] = {
    { "Path", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_6[] = {
    { "Path", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_7[] = {
    { "Path", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_9[] = {
    { "DriveSpec", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_10[] = {
    { "FileSpec", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_11[] = {
    { "FolderSpec", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_12[] = {
    { "DriveSpec", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_13[] = {
    { "FilePath", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_14[] = {
    { "FolderPath", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_15[] = {
    { "SpecialFolder", Vb6Type::Long, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_16[] = {
    { "FileSpec", Vb6Type::String, ComParamDir::In, false },
    { "Force", Vb6Type::Boolean, ComParamDir::In, true },
};
static const BuiltinParam kBP_FileSystemObject_17[] = {
    { "FolderSpec", Vb6Type::String, ComParamDir::In, false },
    { "Force", Vb6Type::Boolean, ComParamDir::In, true },
};
static const BuiltinParam kBP_FileSystemObject_18[] = {
    { "Source", Vb6Type::String, ComParamDir::In, false },
    { "Destination", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_19[] = {
    { "Source", Vb6Type::String, ComParamDir::In, false },
    { "Destination", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_20[] = {
    { "Source", Vb6Type::String, ComParamDir::In, false },
    { "Destination", Vb6Type::String, ComParamDir::In, false },
    { "OverWriteFiles", Vb6Type::Boolean, ComParamDir::In, true },
};
static const BuiltinParam kBP_FileSystemObject_21[] = {
    { "Source", Vb6Type::String, ComParamDir::In, false },
    { "Destination", Vb6Type::String, ComParamDir::In, false },
    { "OverWriteFiles", Vb6Type::Boolean, ComParamDir::In, true },
};
static const BuiltinParam kBP_FileSystemObject_22[] = {
    { "Path", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_FileSystemObject_23[] = {
    { "FileName", Vb6Type::String, ComParamDir::In, false },
    { "Overwrite", Vb6Type::Boolean, ComParamDir::In, true },
    { "Unicode", Vb6Type::Boolean, ComParamDir::In, true },
};
static const BuiltinParam kBP_FileSystemObject_24[] = {
    { "FileName", Vb6Type::String, ComParamDir::In, false },
    { "IOMode", Vb6Type::Long, ComParamDir::In, true },
    { "Create", Vb6Type::Boolean, ComParamDir::In, true },
    { "Format", Vb6Type::Long, ComParamDir::In, true },
};
static const BuiltinParam kBP_FileSystemObject_25[] = {
    { "StandardStreamType", Vb6Type::Long, ComParamDir::In, false },
    { "Unicode", Vb6Type::Boolean, ComParamDir::In, true },
};
static const BuiltinParam kBP_FileSystemObject_26[] = {
    { "FileName", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinMember kBM_FileSystemObject[] = {
    { "Drives", ComMemberKind::PropertyGet, Vb6Type::Object, 10010, 7, 0, nullptr },
    { "BuildPath", ComMemberKind::Method, Vb6Type::String, 10000, 8, 2, kBP_FileSystemObject_1 },
    { "GetDriveName", ComMemberKind::Method, Vb6Type::String, 10004, 9, 1, kBP_FileSystemObject_2 },
    { "GetParentFolderName", ComMemberKind::Method, Vb6Type::String, 10005, 10, 1, kBP_FileSystemObject_3 },
    { "GetFileName", ComMemberKind::Method, Vb6Type::String, 10006, 11, 1, kBP_FileSystemObject_4 },
    { "GetBaseName", ComMemberKind::Method, Vb6Type::String, 10007, 12, 1, kBP_FileSystemObject_5 },
    { "GetExtensionName", ComMemberKind::Method, Vb6Type::String, 10008, 13, 1, kBP_FileSystemObject_6 },
    { "GetAbsolutePathName", ComMemberKind::Method, Vb6Type::String, 10002, 14, 1, kBP_FileSystemObject_7 },
    { "GetTempName", ComMemberKind::Method, Vb6Type::String, 10003, 15, 0, nullptr },
    { "DriveExists", ComMemberKind::Method, Vb6Type::Boolean, 10015, 16, 1, kBP_FileSystemObject_9 },
    { "FileExists", ComMemberKind::Method, Vb6Type::Boolean, 10016, 17, 1, kBP_FileSystemObject_10 },
    { "FolderExists", ComMemberKind::Method, Vb6Type::Boolean, 10017, 18, 1, kBP_FileSystemObject_11 },
    { "GetDrive", ComMemberKind::Method, Vb6Type::Object, 10011, 19, 1, kBP_FileSystemObject_12 },
    { "GetFile", ComMemberKind::Method, Vb6Type::Object, 10012, 20, 1, kBP_FileSystemObject_13 },
    { "GetFolder", ComMemberKind::Method, Vb6Type::Object, 10013, 21, 1, kBP_FileSystemObject_14 },
    { "GetSpecialFolder", ComMemberKind::Method, Vb6Type::Object, 10014, 22, 1, kBP_FileSystemObject_15 },
    { "DeleteFile", ComMemberKind::Method, Vb6Type::Empty, 1200, 23, 2, kBP_FileSystemObject_16 },
    { "DeleteFolder", ComMemberKind::Method, Vb6Type::Empty, 1201, 24, 2, kBP_FileSystemObject_17 },
    { "MoveFile", ComMemberKind::Method, Vb6Type::Empty, 1204, 25, 2, kBP_FileSystemObject_18 },
    { "MoveFolder", ComMemberKind::Method, Vb6Type::Empty, 1205, 26, 2, kBP_FileSystemObject_19 },
    { "CopyFile", ComMemberKind::Method, Vb6Type::Empty, 1202, 27, 3, kBP_FileSystemObject_20 },
    { "CopyFolder", ComMemberKind::Method, Vb6Type::Empty, 1203, 28, 3, kBP_FileSystemObject_21 },
    { "CreateFolder", ComMemberKind::Method, Vb6Type::Object, 1120, 29, 1, kBP_FileSystemObject_22 },
    { "CreateTextFile", ComMemberKind::Method, Vb6Type::Object, 1101, 30, 3, kBP_FileSystemObject_23 },
    { "OpenTextFile", ComMemberKind::Method, Vb6Type::Object, 1100, 31, 4, kBP_FileSystemObject_24 },
    { "GetStandardStream", ComMemberKind::Method, Vb6Type::Object, 20000, 32, 2, kBP_FileSystemObject_25 },
    { "GetFileVersion", ComMemberKind::Method, Vb6Type::String, 20010, 33, 1, kBP_FileSystemObject_26 },
};

static const BuiltinParam kBP_Drive_8[] = {
    { "p0", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinMember kBM_Drive[] = {
    { "Path", ComMemberKind::PropertyGet, Vb6Type::String, 0, 7, 0, nullptr },
    { "DriveLetter", ComMemberKind::PropertyGet, Vb6Type::String, 10000, 8, 0, nullptr },
    { "ShareName", ComMemberKind::PropertyGet, Vb6Type::String, 10001, 9, 0, nullptr },
    { "DriveType", ComMemberKind::PropertyGet, Vb6Type::Long, 10002, 10, 0, nullptr },
    { "RootFolder", ComMemberKind::PropertyGet, Vb6Type::Object, 10003, 11, 0, nullptr },
    { "AvailableSpace", ComMemberKind::PropertyGet, Vb6Type::Variant, 10005, 12, 0, nullptr },
    { "FreeSpace", ComMemberKind::PropertyGet, Vb6Type::Variant, 10004, 13, 0, nullptr },
    { "TotalSize", ComMemberKind::PropertyGet, Vb6Type::Variant, 10006, 14, 0, nullptr },
    { "VolumeName", ComMemberKind::PropertyGet, Vb6Type::String, 10007, 15, 1, kBP_Drive_8 },
    { "FileSystem", ComMemberKind::PropertyGet, Vb6Type::String, 10008, 17, 0, nullptr },
    { "SerialNumber", ComMemberKind::PropertyGet, Vb6Type::Long, 10009, 18, 0, nullptr },
    { "IsReady", ComMemberKind::PropertyGet, Vb6Type::Boolean, 10010, 19, 0, nullptr },
};

static const BuiltinParam kBP_Drives_0[] = {
    { "Key", Vb6Type::Variant, ComParamDir::In, false },
};
static const BuiltinMember kBM_Drives[] = {
    { "Item", ComMemberKind::PropertyGet, Vb6Type::Object, 0, 7, 1, kBP_Drives_0 },
    { "_NewEnum", ComMemberKind::PropertyGet, Vb6Type::Object, -4, 8, 0, nullptr },
    { "Count", ComMemberKind::PropertyGet, Vb6Type::Long, 1, 9, 0, nullptr },
};

static const BuiltinParam kBP_Folder_1[] = {
    { "p0", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_Folder_6[] = {
    { "p0", Vb6Type::Long, ComParamDir::In, false },
};
static const BuiltinParam kBP_Folder_11[] = {
    { "Force", Vb6Type::Boolean, ComParamDir::In, true },
};
static const BuiltinParam kBP_Folder_12[] = {
    { "Destination", Vb6Type::String, ComParamDir::In, false },
    { "OverWriteFiles", Vb6Type::Boolean, ComParamDir::In, true },
};
static const BuiltinParam kBP_Folder_13[] = {
    { "Destination", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_Folder_18[] = {
    { "FileName", Vb6Type::String, ComParamDir::In, false },
    { "Overwrite", Vb6Type::Boolean, ComParamDir::In, true },
    { "Unicode", Vb6Type::Boolean, ComParamDir::In, true },
};
static const BuiltinMember kBM_Folder[] = {
    { "Path", ComMemberKind::PropertyGet, Vb6Type::String, 0, 7, 0, nullptr },
    { "Name", ComMemberKind::PropertyGet, Vb6Type::String, 1000, 8, 1, kBP_Folder_1 },
    { "ShortPath", ComMemberKind::PropertyGet, Vb6Type::String, 1002, 10, 0, nullptr },
    { "ShortName", ComMemberKind::PropertyGet, Vb6Type::String, 1001, 11, 0, nullptr },
    { "Drive", ComMemberKind::PropertyGet, Vb6Type::Object, 1004, 12, 0, nullptr },
    { "ParentFolder", ComMemberKind::PropertyGet, Vb6Type::Object, 1005, 13, 0, nullptr },
    { "Attributes", ComMemberKind::PropertyGet, Vb6Type::Long, 1003, 14, 1, kBP_Folder_6 },
    { "DateCreated", ComMemberKind::PropertyGet, Vb6Type::Date, 1006, 16, 0, nullptr },
    { "DateLastModified", ComMemberKind::PropertyGet, Vb6Type::Date, 1007, 17, 0, nullptr },
    { "DateLastAccessed", ComMemberKind::PropertyGet, Vb6Type::Date, 1008, 18, 0, nullptr },
    { "Type", ComMemberKind::PropertyGet, Vb6Type::String, 1010, 19, 0, nullptr },
    { "Delete", ComMemberKind::Method, Vb6Type::Empty, 1201, 20, 1, kBP_Folder_11 },
    { "Copy", ComMemberKind::Method, Vb6Type::Empty, 1203, 21, 2, kBP_Folder_12 },
    { "Move", ComMemberKind::Method, Vb6Type::Empty, 1205, 22, 1, kBP_Folder_13 },
    { "IsRootFolder", ComMemberKind::PropertyGet, Vb6Type::Boolean, 10000, 23, 0, nullptr },
    { "Size", ComMemberKind::PropertyGet, Vb6Type::Variant, 1009, 24, 0, nullptr },
    { "SubFolders", ComMemberKind::PropertyGet, Vb6Type::Object, 10001, 25, 0, nullptr },
    { "Files", ComMemberKind::PropertyGet, Vb6Type::Object, 10002, 26, 0, nullptr },
    { "CreateTextFile", ComMemberKind::Method, Vb6Type::Object, 1101, 27, 3, kBP_Folder_18 },
};

static const BuiltinParam kBP_Folders_0[] = {
    { "Name", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_Folders_1[] = {
    { "Key", Vb6Type::Variant, ComParamDir::In, false },
};
static const BuiltinMember kBM_Folders[] = {
    { "Add", ComMemberKind::Method, Vb6Type::Object, 2, 7, 1, kBP_Folders_0 },
    { "Item", ComMemberKind::PropertyGet, Vb6Type::Object, 0, 8, 1, kBP_Folders_1 },
    { "_NewEnum", ComMemberKind::PropertyGet, Vb6Type::Object, -4, 9, 0, nullptr },
    { "Count", ComMemberKind::PropertyGet, Vb6Type::Long, 1, 10, 0, nullptr },
};

static const BuiltinParam kBP_File_1[] = {
    { "p0", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_File_6[] = {
    { "p0", Vb6Type::Long, ComParamDir::In, false },
};
static const BuiltinParam kBP_File_12[] = {
    { "Force", Vb6Type::Boolean, ComParamDir::In, true },
};
static const BuiltinParam kBP_File_13[] = {
    { "Destination", Vb6Type::String, ComParamDir::In, false },
    { "OverWriteFiles", Vb6Type::Boolean, ComParamDir::In, true },
};
static const BuiltinParam kBP_File_14[] = {
    { "Destination", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_File_15[] = {
    { "IOMode", Vb6Type::Long, ComParamDir::In, true },
    { "Format", Vb6Type::Long, ComParamDir::In, true },
};
static const BuiltinMember kBM_File[] = {
    { "Path", ComMemberKind::PropertyGet, Vb6Type::String, 0, 7, 0, nullptr },
    { "Name", ComMemberKind::PropertyGet, Vb6Type::String, 1000, 8, 1, kBP_File_1 },
    { "ShortPath", ComMemberKind::PropertyGet, Vb6Type::String, 1002, 10, 0, nullptr },
    { "ShortName", ComMemberKind::PropertyGet, Vb6Type::String, 1001, 11, 0, nullptr },
    { "Drive", ComMemberKind::PropertyGet, Vb6Type::Object, 1004, 12, 0, nullptr },
    { "ParentFolder", ComMemberKind::PropertyGet, Vb6Type::Object, 1005, 13, 0, nullptr },
    { "Attributes", ComMemberKind::PropertyGet, Vb6Type::Long, 1003, 14, 1, kBP_File_6 },
    { "DateCreated", ComMemberKind::PropertyGet, Vb6Type::Date, 1006, 16, 0, nullptr },
    { "DateLastModified", ComMemberKind::PropertyGet, Vb6Type::Date, 1007, 17, 0, nullptr },
    { "DateLastAccessed", ComMemberKind::PropertyGet, Vb6Type::Date, 1008, 18, 0, nullptr },
    { "Size", ComMemberKind::PropertyGet, Vb6Type::Variant, 1009, 19, 0, nullptr },
    { "Type", ComMemberKind::PropertyGet, Vb6Type::String, 1010, 20, 0, nullptr },
    { "Delete", ComMemberKind::Method, Vb6Type::Empty, 1200, 21, 1, kBP_File_12 },
    { "Copy", ComMemberKind::Method, Vb6Type::Empty, 1202, 22, 2, kBP_File_13 },
    { "Move", ComMemberKind::Method, Vb6Type::Empty, 1204, 23, 1, kBP_File_14 },
    { "OpenAsTextStream", ComMemberKind::Method, Vb6Type::Object, 1100, 24, 2, kBP_File_15 },
};

static const BuiltinParam kBP_Files_0[] = {
    { "Key", Vb6Type::Variant, ComParamDir::In, false },
};
static const BuiltinMember kBM_Files[] = {
    { "Item", ComMemberKind::PropertyGet, Vb6Type::Object, 0, 7, 1, kBP_Files_0 },
    { "_NewEnum", ComMemberKind::PropertyGet, Vb6Type::Object, -4, 8, 0, nullptr },
    { "Count", ComMemberKind::PropertyGet, Vb6Type::Long, 1, 9, 0, nullptr },
};

static const BuiltinParam kBP_TextStream_4[] = {
    { "Characters", Vb6Type::Long, ComParamDir::In, false },
};
static const BuiltinParam kBP_TextStream_7[] = {
    { "Text", Vb6Type::String, ComParamDir::In, false },
};
static const BuiltinParam kBP_TextStream_8[] = {
    { "Text", Vb6Type::String, ComParamDir::In, true },
};
static const BuiltinParam kBP_TextStream_9[] = {
    { "Lines", Vb6Type::Long, ComParamDir::In, false },
};
static const BuiltinParam kBP_TextStream_10[] = {
    { "Characters", Vb6Type::Long, ComParamDir::In, false },
};
static const BuiltinMember kBM_TextStream[] = {
    { "Line", ComMemberKind::PropertyGet, Vb6Type::Long, 10000, 7, 0, nullptr },
    { "Column", ComMemberKind::PropertyGet, Vb6Type::Long, -529, 8, 0, nullptr },
    { "AtEndOfStream", ComMemberKind::PropertyGet, Vb6Type::Boolean, 10002, 9, 0, nullptr },
    { "AtEndOfLine", ComMemberKind::PropertyGet, Vb6Type::Boolean, 10003, 10, 0, nullptr },
    { "Read", ComMemberKind::Method, Vb6Type::String, 10004, 11, 1, kBP_TextStream_4 },
    { "ReadLine", ComMemberKind::Method, Vb6Type::String, 10005, 12, 0, nullptr },
    { "ReadAll", ComMemberKind::Method, Vb6Type::String, 10006, 13, 0, nullptr },
    { "Write", ComMemberKind::Method, Vb6Type::Empty, 10007, 14, 1, kBP_TextStream_7 },
    { "WriteLine", ComMemberKind::Method, Vb6Type::Empty, 10008, 15, 1, kBP_TextStream_8 },
    { "WriteBlankLines", ComMemberKind::Method, Vb6Type::Empty, 10009, 16, 1, kBP_TextStream_9 },
    { "Skip", ComMemberKind::Method, Vb6Type::Empty, 10010, 17, 1, kBP_TextStream_10 },
    { "SkipLine", ComMemberKind::Method, Vb6Type::Empty, 10011, 18, 0, nullptr },
    { "Close", ComMemberKind::Method, Vb6Type::Empty, 10012, 19, 0, nullptr },
};

static const BuiltinClass kBuiltinClasses[] = {
    { "Dictionary", "Scripting.Dictionary", "IDictionary", "{42C642C1-97E1-11CF-978F-00A02463E06F}", true, true, "item", "Item", 12, kBM_Dictionary },
    { "FileSystemObject", "Scripting.FileSystemObject", "IFileSystem3", "{2A0B9D10-4B87-11D3-A97A-00104B365C9F}", true, true, "", "", 27, kBM_FileSystemObject },
    { "Drive", "Scripting.Drive", "IDrive", "{C7C3F5A0-88A3-11D0-ABCB-00A0C90FFFC0}", true, true, "path", "Path", 12, kBM_Drive },
    { "Drives", "Scripting.Drives", "IDriveCollection", "{C7C3F5A1-88A3-11D0-ABCB-00A0C90FFFC0}", true, true, "item", "Item", 3, kBM_Drives },
    { "Folder", "Scripting.Folder", "IFolder", "{C7C3F5A2-88A3-11D0-ABCB-00A0C90FFFC0}", true, true, "path", "Path", 19, kBM_Folder },
    { "Folders", "Scripting.Folders", "IFolderCollection", "{C7C3F5A3-88A3-11D0-ABCB-00A0C90FFFC0}", true, true, "item", "Item", 4, kBM_Folders },
    { "File", "Scripting.File", "IFile", "{C7C3F5A4-88A3-11D0-ABCB-00A0C90FFFC0}", true, true, "path", "Path", 16, kBM_File },
    { "Files", "Scripting.Files", "IFileCollection", "{C7C3F5A5-88A3-11D0-ABCB-00A0C90FFFC0}", true, true, "item", "Item", 3, kBM_Files },
    { "TextStream", "Scripting.TextStream", "ITextStream", "{53BAD8C1-E718-11CF-893D-00A0C9054228}", true, true, "", "", 13, kBM_TextStream },
};

