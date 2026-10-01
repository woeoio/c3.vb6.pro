#pragma once
// VB6 工程文件 (.vbp) 解析器
// .vbp 是纯文本行导向的INI风格文件, 每行 Key=Value

#include <string>
#include <vector>
#include <filesystem>
#include "common/encoding.hpp"

namespace vb6c3 {

// 源文件条目类型
enum class VbpSourceType {
    Unknown,        // 未识别的源文件类型
    Form,           // Form=xxx.frm
    Module,         // Module=Name; xxx.bas
    Class,          // Class=Name; xxx.cls
    UserControl,    // UserControl=Name; xxx.ctl
    PropertyPage,   // PropertyPage=Name; xxx.pag
    Designer,       // Designer=Name; xxx.dsr
};

// 源文件条目
struct VbpSourceEntry {
    VbpSourceType type;
    std::string moduleName;     // 模块名 (Form无此字段, Module/Class有)
    std::string filePath;       // 文件路径 (相对或绝对)
    std::string clsidStr;       // P6.8: CLSID (仅Class, 格式: Class=Name; File.cls; {CLSID})
};

// 项目类型
enum class VbpProjectType {
    StandardExe,       // Type=Exe
    ActiveXExe,        // Type=OLE Exe
    ActiveXDLL,        // Type=DLL
    ActiveXControl,    // Type=Control
    Unknown,
};

// VBP 工程信息
struct VbpProject {
    // 项目基本信息
    VbpProjectType projectType = VbpProjectType::Unknown;
    std::string projectName;       // Name="xxx"
    std::string title;             // Title="xxx"
    std::string exeName;           // ExeName32="xxx.exe"
    std::string outputPath;        // Path32="xxx"
    std::string startupObject;     // Startup="Sub Main" 或 "Form1"
    std::string iconForm;          // IconForm="Form1"
    std::string helpFile;          // HelpFile=""
    std::string commandLine;       // Command32=""

    // 编译选项
    int compilationType = 0;       // 0=P-Code, 1=Native Code
    int optimizationType = 0;

    // 源文件列表
    std::vector<VbpSourceEntry> sources;

    // 引用 (类型库)
    struct Reference {
        std::string guid;
        int major = 0;
        int minor = 0;
        std::string path;
        std::string description;
    };
    std::vector<Reference> references;

    // Object引用 (OCX控件)
    struct ObjectRef {
        std::string guid;
        int major = 0;
        int minor = 0;
        std::string fileName;
    };
    std::vector<ObjectRef> objects;

    // Fix <vbeclipse>: Object=*\A<路径>.vbp 私有控件工程引用 (相对本 vbp 目录)
    std::vector<std::string> projectRefs;

    // C3 扩展 (Fix 160): 免注册 COM 组件 DLL。
    // 格式: ComLib=<相对 exe 的 DLL 路径>  (可多行, 一行一个组件 DLL)
    // VB6 标准 VBP 不含此字段。用途: 运行期 CreateObject("ProgID") 在目标机未注册
    // 时, 改走 LoadLibrary(<该 DLL>) + DllGetClassObject 免注册激活, 与 Object=
    // 控件的免注册路径 (Fix 143/148) 同款机制。路径相对 exe, 便于便携分发。
    // C3 加载该 DLL 的 typelib, 枚举其中全部 coclass, 烘焙
    // {ProgID, CLSID, coclass名, 路径} 表进产物; 用户无需手写 ProgID/CLSID。
    std::vector<std::string> comLibs;

    // 静态库搜索根 (ai/024 E2, 批次 T01)。
    // 格式: LibDir=<目录>   (可多行, 一行一个根; 顺序即搜索顺序)
    //       一行内也可用 ';' 分隔多个目录 (便利写法, Windows 路径不含 ';')
    // 语义与 ComLib= 同族: 相对路径以【vbp 所在目录】为基准 (resolvePath), 绝不引入 cwd。
    // 用途: 静态模式 Declare (Lib "xxx.lib") 的裸文件名在哪些目录下找。
    // 缺省根 <vbp目录>/Lib 由 driver 自动追加, 无需在此列出。
    // VB6 标准 VBP 不含此字段, 是 C3 扩展。
    std::vector<std::string> libDirs;

    // 附加静态库 (ai/024 M4, 批次 T01)。
    // 格式: ExtraLib=<.lib/.obj 路径或裸名>   (可多行; 一行内可用 ';' 分隔多个)
    // 用途: 链接**没有对应 Declare** 的静态库/目标文件 —— 典型场景是静态库之间互相
    //       依赖 (A.lib 内部引用了 B.lib 的符号), 用户不会为 B 写 Declare。
    // 寻址与静态 Declare 的 Lib 串**完全同一条路** (绝对 → 工程相对 → 搜索根),
    // 所以裸名也吃 LibDir= / --libdir / <vbp目录>/Lib。
    // 与 Lib "x.lib" 的分工: 写 Declare 的地方就写 Lib, 纯链接依赖才写 ExtraLib。
    std::vector<std::string> extraLibs;

    // 工程引用/包 (ai/023 S01)。
    // 格式: Package=Name; Version   (可多行, 一行一个包)
    // vbp 里只写名字+版本, 不写路径 (023 D3); 寻址唯一:
    //   <搜索根>/<Name>-<Version>/package.c3d
    // 搜索根 = <vbp目录>/packages (缺省) + CLI --package-root (追加)。
    // VB6 标准 VBP 不含此字段, 是 C3 扩展。
    struct PackageRef {
        std::string name;     // 包名 (A-Za-z0-9_); 也是限定名前缀 Pkg.Mod.Func
        std::string version;  // "1.2" → 目录 VBFlexGrid-1.2 (023 六节, '-' 分隔)
    };
    std::vector<PackageRef> packageRefs;

    // 资源文件
    std::string resFile;

    // 版本信息 (P20-22)
    int majorVer = 1;
    int minorVer = 0;
    int revisionVer = 0;
    int autoIncrementVer = 0;
    std::string companyName;
    std::string fileDescription;
    std::string legalCopyright;
    std::string productName;
    std::string comments;
    std::string legalTrademarks;
    std::string originalFileName;

    // P22: ActiveX 二进制兼容和启动模式
    int compatibleMode = 0;  // 0=None, 1=Project, 2=Binary
    int startMode = 0;       // 0=Standalone, 1=ActiveX

    // C3 扩展: ActiveX DLL TypeLib 的 LibID (UUID 格式 "{...}")
    // VB6 标准 VBP 不含此字段 (VB6 通过二进制兼容模式管理 LibID)
    // C3 用此字段让 LibID 跨构建稳定, 避免每次编译生成新 LibID 导致
    // VBA References 失效和注册表残留堆积
    std::string libidStr;

    // .vbp 文件路径 (用于解析相对路径)
    std::filesystem::path vbpFilePath;

    // 获取源文件的绝对路径
    std::filesystem::path resolvePath(const std::string& relativePath) const {
        // relativePath is UTF-8 (from VBP parser GBK→UTF-8 conversion)
        // Must use utf8ToPath() to avoid ACP reinterpretation on Windows
        if (utf8ToPath(relativePath).is_absolute()) {
            return utf8ToPath(relativePath);
        }
        return vbpFilePath.parent_path() / utf8ToPath(relativePath);
    }
};

// VBP 文件解析器
class VbpParser {
public:
    // 解析 .vbp 文件
    static VbpProject parse(const std::string& vbpFilePath);

    // 解析 .vbp 内容字符串
    static VbpProject parseString(const std::string& content, const std::string& vbpFilePath = "");

private:
    static VbpProjectType parseProjectType(const std::string& value);
    static VbpSourceType parseSourceType(const std::string& key);
    static std::string unquote(const std::string& s);
};

} // namespace vb6c3
