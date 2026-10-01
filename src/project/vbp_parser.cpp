// vb6c3 - VB6工程文件(.vbp)解析器
// .vbp 是纯文本行导向的INI风格文件, 每行 Key=Value

#include "project/vbp_parser.hpp"
#include "common/source_manager.hpp"
#include <sstream>
#include <algorithm>

namespace vb6c3 {

VbpProject VbpParser::parse(const std::string& vbpFilePath) {
    VbpProject project;
    // Fix 196: 必须经 utf8ToPath —— std::filesystem::absolute(窄串) 在 Windows 上按
    // **ACP** 解释 char*, 而这里收的是 UTF-8。中文目录下 vbpFilePath 会变成乱码,
    // 于是 resolvePath() 拼出的每个源文件路径都不存在 (后果见 frm_parser 的静默空模块)。
    project.vbpFilePath = std::filesystem::absolute(utf8ToPath(vbpFilePath));

    // M22: 使用编码检测+转换读取, 确保GBK等非UTF-8文件正确解码
    auto readResult = SourceBuffer::readAndConvertToUtf8(vbpFilePath);
    if (readResult.content.empty()) {
        return project;
    }

    return parseString(readResult.content, vbpFilePath);
}

VbpProject VbpParser::parseString(const std::string& content, const std::string& vbpFilePath) {
    VbpProject project;
    if (!vbpFilePath.empty()) {
        // Fix 196: 同 parse(), 窄串是 UTF-8, 不走 utf8ToPath 会被按 ACP 解释
        project.vbpFilePath = std::filesystem::absolute(utf8ToPath(vbpFilePath));
    }

    std::istringstream stream(content);
    std::string line;

    while (std::getline(stream, line)) {
        // 去除行尾 \r
        if (!line.empty() && line.back() == '\r') {
            line.pop_back();
        }

        // 跳过空行和注释
        if (line.empty() || line[0] == '\'') continue;

        // 跳过扩展段 [Section]
        if (line[0] == '[') continue;

        // 解析 Key=Value
        auto eqPos = line.find('=');
        if (eqPos == std::string::npos) continue;

        std::string key = line.substr(0, eqPos);
        std::string value = line.substr(eqPos + 1);

        // 根据 key 分类处理
        VbpSourceType srcType = parseSourceType(key);
        if (srcType != VbpSourceType::Unknown ||
            key == "Form" || key == "Module" || key == "Class" ||
            key == "UserControl" || key == "PropertyPage" || key == "Designer") {

            // 确定源文件类型
            if (key == "Form") srcType = VbpSourceType::Form;
            else if (key == "Module") srcType = VbpSourceType::Module;
            else if (key == "Class") srcType = VbpSourceType::Class;
            else if (key == "UserControl") srcType = VbpSourceType::UserControl;
            else if (key == "PropertyPage") srcType = VbpSourceType::PropertyPage;
            else if (key == "Designer") srcType = VbpSourceType::Designer;

            VbpSourceEntry entry;
            entry.type = srcType;

            if (srcType == VbpSourceType::Form) {
                // Form=xxx.frm (无模块名)
                entry.filePath = value;
            } else {
                // Module=Name; xxx.bas 或 Class=Name; xxx.cls
                // Class还支持3段格式: Class=Name; xxx.cls; {CLSID}
                // 分号分隔: 可能是 "Name; File" 或 "Name; File; {CLSID}"
                auto semiPos = value.find(';');
                if (semiPos != std::string::npos) {
                    entry.moduleName = value.substr(0, semiPos);
                    auto rest = value.substr(semiPos + 1);
                    // 检查是否有第三个段 (CLSID)
                    auto semiPos2 = rest.find(';');
                    if (semiPos2 != std::string::npos) {
                        // 3段格式: File; {CLSID}
                        auto filePart = rest.substr(0, semiPos2);
                        auto clsidPart = rest.substr(semiPos2 + 1);
                        // ltrim/rtrim filePart
                        size_t start = filePart.find_first_not_of(" \t");
                        if (start != std::string::npos) {
                            entry.filePath = filePart.substr(start);
                        } else {
                            entry.filePath = filePart;
                        }
                        while (!entry.filePath.empty() &&
                               (entry.filePath.back() == ' ' || entry.filePath.back() == '\t')) {
                            entry.filePath.pop_back();
                        }
                        // 解析CLSID: 去除前后空白, 验证{...}格式
                        size_t cs = clsidPart.find_first_not_of(" \t");
                        if (cs != std::string::npos) clsidPart = clsidPart.substr(cs);
                        while (!clsidPart.empty() &&
                               (clsidPart.back() == ' ' || clsidPart.back() == '\t')) {
                            clsidPart.pop_back();
                        }
                        if (clsidPart.size() >= 2 && clsidPart.front() == '{' && clsidPart.back() == '}') {
                            entry.clsidStr = clsidPart;  // P6.8: 存储CLSID
                        }
                    } else {
                        // 2段格式: Name; File
                        size_t start = rest.find_first_not_of(" \t");
                        if (start != std::string::npos) {
                            entry.filePath = rest.substr(start);
                        } else {
                            entry.filePath = rest;
                        }
                        while (!entry.filePath.empty() &&
                               (entry.filePath.back() == ' ' || entry.filePath.back() == '\t')) {
                            entry.filePath.pop_back();
                        }
                    }
                } else {
                    entry.filePath = value;
                }
            }

            if (!entry.filePath.empty()) {
                project.sources.push_back(std::move(entry));
            }
            continue;
        }

        // Reference=*\G{GUID}#major#minor#lcid#path#desc
        if (key == "Reference") {
            VbpProject::Reference ref;
            // 格式: *\G{GUID}#major#minor#lcid#path#desc
            // 简化解析: 按 # 分割
            std::vector<std::string> parts;
            std::istringstream refStream(value);
            std::string part;
            while (std::getline(refStream, part, '#')) {
                parts.push_back(part);
            }
            if (parts.size() >= 2) {
                // parts[0] = *\G{GUID} 或 {GUID}
                std::string guidPart = parts[0];
                auto gPos = guidPart.find("{");
                if (gPos != std::string::npos) {
                    auto ePos = guidPart.find("}", gPos);
                    if (ePos != std::string::npos) {
                        ref.guid = guidPart.substr(gPos, ePos - gPos + 1);
                    }
                }
            }
            if (parts.size() >= 3) {
                try { ref.major = std::stoi(parts[1]); } catch (...) {}
                try { ref.minor = std::stoi(parts[2]); } catch (...) {}
            }
            // VBP Reference: GUID#version#lcid#path#desc (5 fields standard)
            // Also: GUID#major#minor#lcid#path#desc (6 fields)
            if (parts.size() == 5) {
                ref.path = parts[3];
                ref.description = parts[4];
            } else if (parts.size() >= 6) {
                ref.path = parts[4];
                ref.description = parts[5];
            }
            project.references.push_back(std::move(ref));
            continue;
        }

        // Object={GUID}#major#minor; filename.ocx
        if (key == "Object") {
            // Fix <vbeclipse>: Object=*\A<相对路径>.vbp —— 私有控件工程引用。
            // VB6 里被引工程的 UserControl/类/模块全部公开符号对本工程可见,
            // 控件实例也随 EXE 一起编译。这里只登记路径, 源码展开在 driver_compile。
            if (value.size() > 3 && value[0] == '*' && value[1] == '\\' && value[2] == 'A') {
                std::string refPath = value.substr(3);
                size_t s2 = refPath.find_first_not_of(" \t");
                if (s2 != std::string::npos) refPath = refPath.substr(s2);
                while (!refPath.empty() &&
                       (refPath.back() == ' ' || refPath.back() == '\t')) {
                    refPath.pop_back();
                }
                project.projectRefs.push_back(refPath);
                continue;
            }
            VbpProject::ObjectRef obj;
            // 格式: {GUID}#major#minor; filename.ocx
            // 用 ; 分割最后部分
            auto semiPos = value.find(';');
            std::string guidPart = (semiPos != std::string::npos)
                ? value.substr(0, semiPos) : value;
            std::string filePart = (semiPos != std::string::npos)
                ? value.substr(semiPos + 1) : "";

            // ltrim filePart
            size_t start = filePart.find_first_not_of(" \t");
            if (start != std::string::npos) {
                obj.fileName = filePart.substr(start);
            }

            // 从 guidPart 中提取 GUID 和版本
            auto gPos = guidPart.find("{");
            if (gPos != std::string::npos) {
                auto ePos = guidPart.find("}", gPos);
                if (ePos != std::string::npos) {
                    obj.guid = guidPart.substr(gPos, ePos - gPos + 1);
                }
            }
            // 版本: #major#minor
            auto hashPos = guidPart.find('#');
            if (hashPos != std::string::npos) {
                auto rest = guidPart.substr(hashPos + 1);
                auto hashPos2 = rest.find('#');
                if (hashPos2 != std::string::npos) {
                    try { obj.major = std::stoi(rest.substr(0, hashPos2)); } catch (...) {}
                    try { obj.minor = std::stoi(rest.substr(hashPos2 + 1)); } catch (...) {}
                }
            }

            project.objects.push_back(std::move(obj));
            continue;
        }

        // C3 扩展 (ai/023 S01): Package=Name; Version —— 工程引用/包。
        // 只记名字+版本, 不写路径 (023 D3); 解析与三条硬校验在 driver 的
        // checkPackages (driver_compile.cpp), 这里不做诊断 (parseString 无 diag 通道)。
        if (key == "Package") {
            VbpProject::PackageRef ref;
            auto semiPos = value.find(';');
            if (semiPos == std::string::npos) {
                // 宽容写法: Package=Foo-1.2 (目录名即引用)。
                // 包名只允许 A-Za-z0-9_ (不含 '-'), 所以首个 '-' 必是分隔位。
                std::string s = value;
                size_t a = s.find_first_not_of(" \t");
                size_t b = s.find_last_not_of(" \t\r\n");
                s = (a != std::string::npos) ? s.substr(a, b - a + 1) : "";
                auto dash = s.find('-');
                if (dash != std::string::npos) {
                    ref.name = s.substr(0, dash);
                    ref.version = s.substr(dash + 1);
                } else {
                    ref.name = s; // 无版本 → 下游 checkPackages 报错
                }
            } else {
                auto trimStr = [](const std::string& x) {
                    size_t a = x.find_first_not_of(" \t");
                    size_t b = x.find_last_not_of(" \t\r\n");
                    return (a != std::string::npos) ? x.substr(a, b - a + 1) : std::string();
                };
                ref.name = trimStr(value.substr(0, semiPos));
                ref.version = trimStr(value.substr(semiPos + 1));
            }
            if (!ref.name.empty()) project.packageRefs.push_back(std::move(ref));
            continue;
        }

        // C3 扩展 (Fix 160): ComLib=<相对 exe 的组件 DLL 路径>
        // 免注册 COM: 运行期 CreateObject 在本机未注册时改走 LoadLibrary+DllGetClassObject。
        // 存原始值 (不 trim 内部空白), 路径解析与去重交给 driver_compile。
        if (key == "ComLib") {
            std::string s = value;
            size_t a = s.find_first_not_of(" \t");
            size_t b = s.find_last_not_of(" \t\r\n");
            if (a != std::string::npos && b != std::string::npos) {
                std::string trimmed = s.substr(a, b - a + 1);
                if (!trimmed.empty()) project.comLibs.push_back(std::move(trimmed));
            }
            continue;
        }

        // C3 扩展 (ai/024, 批次 T01): 静态链接相关的两个多行键。
        //   LibDir=<目录>          静态库搜索根 (顺序 = 搜索顺序)
        //   ExtraLib=<.lib/.obj…>  附加静态库 (没有对应 Declare 的链接依赖, 如库间互引)
        // 两键都可多行; 一行内也可用 ';' 分隔多项 (Windows 路径不含 ';')。
        // 顺序保留 = 搜索顺序。相对路径的基准由 driver 统一用 resolvePath 给
        // (vbp 所在目录), 与 ComLib= 同族。
        if (key == "LibDir" || key == "ExtraLib") {
            auto& sink = (key == "LibDir") ? project.libDirs : project.extraLibs;
            const std::string s = value;
            size_t pos = 0;
            while (pos <= s.size()) {
                const size_t sep = s.find(';', pos);
                const std::string item =
                    (sep == std::string::npos) ? s.substr(pos) : s.substr(pos, sep - pos);
                const size_t a = item.find_first_not_of(" \t");
                const size_t b = item.find_last_not_of(" \t\r\n");
                if (a != std::string::npos && b != std::string::npos) {
                    std::string trimmed = item.substr(a, b - a + 1);
                    // 去掉可能带的引号 ("Lib" 与 Lib 等价)
                    if (trimmed.size() >= 2 && trimmed.front() == '"' && trimmed.back() == '"') {
                        trimmed = trimmed.substr(1, trimmed.size() - 2);
                    }
                    if (!trimmed.empty()) sink.push_back(std::move(trimmed));
                }
                if (sep == std::string::npos) break;
                pos = sep + 1;
            }
            continue;
        }

        // 项目属性
        if (key == "Type") {
            project.projectType = parseProjectType(value);
        } else if (key == "Name") {
            project.projectName = unquote(value);
        } else if (key == "Title") {
            project.title = unquote(value);
        } else if (key == "ExeName32") {
            project.exeName = unquote(value);
        } else if (key == "Path32") {
            project.outputPath = unquote(value);
        } else if (key == "Startup") {
            project.startupObject = unquote(value);
        } else if (key == "IconForm") {
            project.iconForm = unquote(value);
        } else if (key == "HelpFile") {
            project.helpFile = unquote(value);
        } else if (key == "Command32") {
            project.commandLine = unquote(value);
        } else if (key == "ResFile32") {
            project.resFile = unquote(value);
        } else if (key == "CompilationType") {
            try { project.compilationType = std::stoi(value); } catch (...) {}
        } else if (key == "OptimizationType") {
            try { project.optimizationType = std::stoi(value); } catch (...) {}
        } else if (key == "CompatibleMode") {
            try { project.compatibleMode = std::stoi(value); } catch (...) {}
        } else if (key == "StartMode") {
            try { project.startMode = std::stoi(value); } catch (...) {}
        }
        // C3 扩展: LibID={...} 固定 TypeLib 的 LibID, 跨构建稳定
        // 值格式: "{xxxxxxxx-xxxx-...}" 或裸 GUID (不带引号)
        else if (key == "LibID") {
            std::string v = value;
            // 兼容用户写成 LibID="{...}" 带引号的形式
            if (v.size() >= 2 && v.front() == '"' && v.back() == '"') {
                v = v.substr(1, v.size() - 2);
            }
            project.libidStr = v;
        }
        // 版本信息 (P20-22)
        else if (key == "MajorVer") { try { project.majorVer = std::stoi(value); } catch (...) {} }
        else if (key == "MinorVer") { try { project.minorVer = std::stoi(value); } catch (...) {} }
        else if (key == "RevisionVer") { try { project.revisionVer = std::stoi(value); } catch (...) {} }
        else if (key == "AutoIncrementVer") { try { project.autoIncrementVer = std::stoi(value); } catch (...) {} }
        else if (key == "VersionCompanyName") { project.companyName = unquote(value); }
        else if (key == "VersionFileDescription") { project.fileDescription = unquote(value); }
        else if (key == "VersionLegalCopyright") { project.legalCopyright = unquote(value); }
        else if (key == "VersionProductName") { project.productName = unquote(value); }
        else if (key == "VersionComments") { project.comments = unquote(value); }
        else if (key == "VersionLegalTrademarks") { project.legalTrademarks = unquote(value); }
        else if (key == "VersionOriginalFileName") { project.originalFileName = unquote(value); }
    }

    return project;
}

VbpProjectType VbpParser::parseProjectType(const std::string& value) {
    if (value == "Exe") return VbpProjectType::StandardExe;
    if (value == "OLE Exe") return VbpProjectType::ActiveXExe;
    // 真 VB6 保存 ActiveX DLL 工程时写的是 "OleDll" (无空格连写), 兼容旧写法 "DLL"
    if (value == "OleDll") return VbpProjectType::ActiveXDLL;
    if (value == "DLL") return VbpProjectType::ActiveXDLL;
    if (value == "Control") return VbpProjectType::ActiveXControl;
    return VbpProjectType::Unknown;
}

VbpSourceType VbpParser::parseSourceType(const std::string& key) {
    // 已在内联处理
    return VbpSourceType::Unknown;
}

std::string VbpParser::unquote(const std::string& s) {
    if (s.size() >= 2 && s.front() == '"' && s.back() == '"') {
        return s.substr(1, s.size() - 2);
    }
    return s;
}

} // namespace vb6c3
