// driver_frontend.cpp - C3 编译器驱动: 前端流水线（词法/预处理/解析/TypeLib 导入）
// 2026-09-17 从 src/driver/driver.cpp 纯搬移（逐行未改）：
//   原第 547~865 行

#include "driver/driver.hpp"
#include "common/diagnostics.hpp"
#include "common/encoding.hpp"
#include "common/source_manager.hpp"
#include "lexer/lexer.hpp"
#include "lexer/token.hpp"
#include "parser/parser.hpp"
#include "ast/ast_printer.hpp"
#include <iostream>
#include <filesystem>
#include <algorithm>
#include <cctype>
#include <unordered_set>

namespace vb6c3 {

// === 词法分析阶段 ===

bool Driver::runLexer(const CompileOptions& options) {
    if (options.dumpTokens) {
        for (const auto& filePath : options.sourceFiles) {
            auto buffer = SourceBuffer::fromFile(filePath);
            if (!buffer) {
                SourceLocation loc{filePath, 0, 0};
                diag_->error(DiagnosticID::LexFileEncodingError, loc,
                    "无法打开文件: " + filePath);
                return false;
            }

            Lexer lexer(std::move(buffer), *diag_);
            Token tok;
            do {
                tok = lexer.nextToken();
                std::cout << tok.toString() << std::endl;
            } while (tok.kind != TokenKind::EndOfFile);
        }

        // dump-tokens模式下, 输出完token就结束
        return true;
    }

    // 正常模式: 词法分析结果传给语法分析
    // TODO: 将token流缓存供后续阶段使用
    return true;
}

// === 预处理阶段 (条件编译) ===

bool Driver::runPreprocess(const CompileOptions& options) {
    if (!options.dumpPreprocess && options.defines.empty()) {
        // 无需单独预处理, 由Parser内部的Preprocessor处理
        return true;
    }

    if (options.dumpPreprocess) {
        // 解析命令行定义
        PreprocessOptions ppOpts;
        ppOpts.is64Bit = (options.arch == "x64");  // Fix 081h
        for (const auto& def : options.defines) {
            // NAME=VALUE 格式
            auto eq = def.find('=');
            if (eq != std::string::npos) {
                std::string name = def.substr(0, eq);
                std::string valStr = def.substr(eq + 1);
                std::string nameLower = name;
                std::transform(nameLower.begin(), nameLower.end(), nameLower.begin(), ::tolower);
                // 尝试解析为布尔或整数
                if (valStr == "True" || valStr == "true" || valStr == "-1") {
                    ppOpts.defines[nameLower] = CondCompileValue::fromBool(true);
                } else if (valStr == "False" || valStr == "false" || valStr == "0") {
                    ppOpts.defines[nameLower] = CondCompileValue::fromBool(false);
                } else {
                    try {
                        ppOpts.defines[nameLower] = CondCompileValue::fromLong(std::stoll(valStr));
                    } catch (...) {
                        ppOpts.defines[nameLower] = CondCompileValue::fromBool(false);
                    }
                }
            }
        }

        for (const auto& filePath : options.sourceFiles) {
            auto buffer = SourceBuffer::fromFile(filePath);
            if (!buffer) {
                SourceLocation loc{filePath, 0, 0};
                diag_->error(DiagnosticID::LexFileEncodingError, loc,
                    "无法打开文件: " + filePath);
                return false;
            }

            Preprocessor preproc(std::move(buffer), *diag_, ppOpts);
            Token tok;
            do {
                tok = preproc.nextToken();
                std::cout << tok.toString() << std::endl;
            } while (tok.kind != TokenKind::EndOfFile);

            // 输出条件编译常量表
            if (options.verbose) {
                std::cout << "\n--- 条件编译常量 ---" << std::endl;
                for (const auto& [name, val] : preproc.constants()) {
                    std::cout << "  " << name << " = ";
                    if (val.kind == CondCompileValue::Boolean) {
                        std::cout << (val.boolValue ? "True" : "False");
                    } else if (val.kind == CondCompileValue::Long) {
                        std::cout << val.longValue;
                    } else {
                        std::cout << "<undefined>";
                    }
                    std::cout << std::endl;
                }
            }
        }
    }

    return true;
}

// === 解析阶段 ===

bool Driver::runParser(const CompileOptions& options) {
    // 解析命令行定义
    PreprocessOptions ppOpts;
    ppOpts.is64Bit = (options.arch == "x64");  // Fix 081h
    for (const auto& def : options.defines) {
        auto eq = def.find('=');
        if (eq != std::string::npos) {
            std::string name = def.substr(0, eq);
            std::string valStr = def.substr(eq + 1);
            std::string nameLower = name;
            std::transform(nameLower.begin(), nameLower.end(), nameLower.begin(), ::tolower);
            if (valStr == "True" || valStr == "true" || valStr == "-1") {
                ppOpts.defines[nameLower] = CondCompileValue::fromBool(true);
            } else if (valStr == "False" || valStr == "false" || valStr == "0") {
                ppOpts.defines[nameLower] = CondCompileValue::fromBool(false);
            } else {
                try {
                    ppOpts.defines[nameLower] = CondCompileValue::fromLong(std::stoll(valStr));
                } catch (...) {
                    ppOpts.defines[nameLower] = CondCompileValue::fromBool(false);
                }
            }
        }
    }

    for (const auto& filePath : options.sourceFiles) {
        // 根据文件扩展名判断模块类型
        bool isClassModule = false;
        bool isFormModule = false;
        bool isControlModule = false;
        bool isPropertyPageModule = false;
        FrmFile frmDesc;  // Designer metadata for .frm/.ctl/.pag
        if (filePath.size() >= 4) {
            std::string ext = filePath.substr(filePath.size() - 4);
            for (auto& c : ext) c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
            isControlModule = (ext == ".ctl");
            isPropertyPageModule = (ext == ".pag");
            isClassModule = (ext == ".cls" || isControlModule || isPropertyPageModule);
            isFormModule = (ext == ".frm");  // P7: 窗体模块
        }

        // M22: 统一编码处理 — .frm经由FrmParser读取(已含编码转换+代码段提取),
        // .bas/.cls经由SourceBuffer::fromFile读取(含编码转换)
        std::unique_ptr<SourceBuffer> buffer;
        if (isFormModule || isControlModule || isPropertyPageModule) {
            // Designer headers are not VB statements; extract the code section.
            // M22: FrmParser.parse已使用readAndConvertToUtf8, 返回的codeSection是UTF-8
            frmDesc = FrmParser::parse(filePath);
            // Fix 196: 读不进来必须报错。原先静默返回空 FrmFile, 于是整张窗体
            // (设计器 + 代码段) 变成空模块, 编出的 exe 什么都不做却报"编译成功" ——
            // 非 ASCII 路径 + 由 .vbp 解析出的路径曾经就是这条静默通路。
            if (frmDesc.readFailed) {
                diag_->error(DiagnosticID::LexFileEncodingError,
                    SourceLocation{filePath, 0, 0},
                    "无法打开文件: " + filePath);
                return false;
            }
            // P24: 设置 .frx 文件路径 (与 .frm 同目录同名)
            // Fix 195: 定位规则收进 resolveFrmResourceFile; 找不到时告警而非静默 ——
            //          原先静默丢弃会让编出的 exe 与 VB6 不一致却毫无提示。
            {
                std::vector<FrmResourceRef> frxRefs;
                collectFrmResourceRefs(frmDesc.form.formControl, frxRefs);

                std::string fallbackExt = isControlModule ? ".ctx" : (isPropertyPageModule ? ".pgx" : ".frx");
                auto frxPath = resolveFrmResourceFile(frmDesc, fallbackExt);

                if (!frxPath.empty()) {
                    frmDesc.form.frxFilePath = frxPath;
                } else if (!frxRefs.empty()) {
                    // 只列前 6 个属性, 多则省略 (诊断行不该被刷屏)
                    std::string props;
                    size_t shown = 0;
                    for (const auto& ref : frxRefs) {
                        if (shown == 6) { props += ", ..."; break; }
                        if (shown) props += ", ";
                        props += ref.where;
                        shown++;
                    }
                    // Fix 196: 两处取文件名都必须走 utf8ToPath/pathToUtf8。
                    //   `std::filesystem::path(std::string)` 的窄串重载按**系统 ACP(936)** 解释
                    //   char*, 而 filePath 是 UTF-8 —— 字节序列一旦不是合法 GBK (路径含奇数个
                    //   非 ASCII 字符时极常见, 如 "中文叉"), MultiByteToWideChar 失败 ->
                    //   _Convert_narrow_to_wide 抛 filesystem_error -> 无人接 -> terminate -> abort()。
                    //   实测: 中文目录 + 缺 .frx 必崩 (退出码 3 + 模态对话框)。
                    std::string refFileName = frxRefs.front().fileName;
                    if (refFileName.empty()) {
                        std::filesystem::path p = std::filesystem::path(frmDesc.frmFilePath);
                        p.replace_extension(fallbackExt);
                        refFileName = pathToUtf8(p.filename());
                    }
                    diag_->warn(DiagnosticID::CodeGenFormResourceMissing,
                        SourceLocation{filePath, 1, 0},
                        "找不到二进制资源文件 " + refFileName + " (被引用的属性: " + props + ")。"
                        "这些属性的设计期取值将被忽略, 编出的程序会与 VB6 不一致。"
                        "该文件由 VB6 保存窗体时自动生成, 请把它与 "
                        + pathToUtf8(utf8ToPath(filePath).filename()) + " 放在同一目录;"
                        "或用 C3 --extract-frx 把它的取值导出成 VB 代码, 从此不再需要它。");
                }
            }
            // Empty designer code is valid; do not feed the header to the VB parser.
            buffer = SourceBuffer::fromString(filePath, frmDesc.codeSection);
        } else {
            buffer = SourceBuffer::fromFile(filePath);
        }

        if (!buffer) {
            SourceLocation loc{filePath, 0, 0};
            diag_->error(DiagnosticID::LexFileEncodingError, loc,
                "无法打开文件: " + filePath);
            return false;
        }

        Parser parser(std::move(buffer), *diag_, ppOpts);
        auto module = parser.parseModule(isClassModule);

        // 泛型 (tB, G2): 合并本模块 parse 期收集的使用点扁名表
        for (const auto& [fk, fv] : parser.flatGenerics()) {
            GenericUseRec rec;
            rec.base = fv.base;
            rec.args = fv.args;
            genericUses_[fk] = std::move(rec);  // 同源内容一致, 覆盖幂等
        }

        // P7: 设置窗体模块标志。文档类别 (.frm/.ctl/.pag) 的唯一写入点就在这里 ——
        // 语义层要按它放行文档隐式对象 (账 #217 第二刀)，发码层也改读它，
        // 所以别再往下的任何地方补第二处判定。
        if (module) {
            if (isFormModule) module->isFormModule = true;
            module->docKind = isControlModule ? DocumentKind::UserControl
                          : isPropertyPageModule ? DocumentKind::PropertyPage
                          : isFormModule ? DocumentKind::Form
                                         : DocumentKind::Standard;
        }

        if (options.dumpAST && module) {
            ASTPrinter printer(std::cout);
            printer.print(*module);
        }

        if (module) {
            // 泛型类 (G4): parse 期头行已写 moduleName; VB_Name 若存在须同名
            const std::string headerName = module->moduleName;
            // 从 Attribute VB_Name 提取模块名
            // VB6 模块名来自 Attribute VB_Name = "ModuleName"
            for (const auto& attr : module->attributes) {
                if (attr->attrName == "VB_Name" && attr->value) {
                    // 值应为字符串字面量
                    if (attr->value->kind == ASTNodeKind::LiteralExpr) {
                        auto& lit = static_cast<LiteralExpr&>(*attr->value);
                        if (lit.literalKind == LiteralKind::String && !lit.rawText.empty()) {
                            // rawText包含引号, 去掉首尾引号
                            std::string name = lit.rawText;
                            if (name.size() >= 2 && name.front() == '"' && name.back() == '"') {
                                name = name.substr(1, name.size() - 2);
                            }
                            module->moduleName = name;
                        }
                    }
                }
            }
            // 泛型类 (G4): 头行名与 VB_Name 属性不一致 → 模板登记名/成员引用
            // 会分裂, 直接拒绝 (特化名以最终 moduleName 为准).
            if (!module->classTypeParams.empty() && !headerName.empty() &&
                !module->moduleName.empty() &&
                Symbol::toLower(headerName) != Symbol::toLower(module->moduleName)) {
                diag_->error(DiagnosticID::ParseExpectedToken, module->loc,
                    "泛型类 Class 头行名 '" + headerName + "' 与 Attribute VB_Name '" +
                    module->moduleName + "' 不一致");
                return false;
            }
            // 如果没有 VB_Name 属性，使用文件名（去掉扩展名）作为模块名
            if (module->moduleName.empty()) {
                std::filesystem::path p(utf8ToPath(filePath));
                module->moduleName = pathToUtf8(p.stem());
            }
            // ai/023 S03: 包归属回填 (S02 在 driver_compile 登记的 文件→包 映射;
            // 查不到 = 宿主源, packageName 留空)
            {
                auto pit = packageOfFile_.find(normSourceKey(filePath));
                if (pit != packageOfFile_.end()) module->packageName = pit->second;
            }
            // P7: 保存窗体描述 (此时moduleName已从Attribute VB_Name或文件名确定)
            // Fix 110: .ctl/.pag 与 .frm 同样需要设计期子控件元数据.
            // 背景: VB6 的 UserControl/PropertyPage 可放置设计期子控件
            // (Begin VB.Timer Timer1 / Begin VB.PictureBox Picture1 ...), 代码里
            // `Timer1.Interval = 100` 必须走"控件属性"路径
            // (vb6_SetTimerInterval(vb6_hwnd_Timer1, 100)). 若缺元数据,
            // knownFormControls_ 为空 → 落入 cgen_assign/cgen_expr_member 的
            // "Module.member" 回退 → vb6_Timer1_Interval 未声明标识符 (C2065).
            if (isFormModule || isControlModule || isPropertyPageModule) {
                frmFiles_[module->moduleName] = std::move(frmDesc);
            }

            modules_.push_back(std::move(module));
        }

        if (diag_->hasErrors()) {
            return false;
        }
    }

    // 类模块必须先于标准模块做语义分析。
    // 原因: 标准模块中引用项目类 (如 Dim WithEvents btn As Button) 时, 需要 Button
    // 的类符号已经在符号表中; 若类模块尚未分析, 类型解析会把 btn 退化为 void*,
    // 方法调用随之退化为 COM 后期绑定 (vb6_ComCall(btn, L"DoClick", ...)),
    // 而项目类实例实际是纯 C 结构体 (vb6_cls_Button*) → 运行期解引用 vtable 崩溃
    // (0xC0000005)。vbp 中 Module 完全可能排在 Class 之前 (如 test_events.vbp),
    // 故在此把类模块稳定前移 (同类之间、Form 与标准模块之间的相对顺序保持不变)。
    std::stable_sort(modules_.begin(), modules_.end(),
                     [](const auto& a, const auto& b) {
                         int rankA = a->isClassModule ? 0 : 1;
                         int rankB = b->isClassModule ? 0 : 1;
                         return rankA < rankB;
                     });

    return !diag_->hasErrors();
}

bool Driver::runTypeLibImport(const CompileOptions& options) {
    // P6.3: 编译期TypeLib导入, 提取COM类型信息用于前期绑定
    // 此阶段在语义分析之前运行, 将TypeLib中的coclass/接口注册为全局符号

    typelibParser_ = std::make_unique<TypeLibParser>(*diag_);

    // 1. 加载显式引用的TypeLib
    for (const auto& ref : options.typelibRefs) {
        // 判断是文件路径还是ProgID
        if (options.verbose) std::cerr << "C3: Loading TypeLib ref: " << ref << std::endl;
        if (ref.find('.') != std::string::npos && ref.find('\\') == std::string::npos && ref.find('/') == std::string::npos) {
            // 含点但不含路径分隔符 → ProgID
            typelibParser_->loadByProgId(ref);
        } else if (ref.find('{') != std::string::npos) {
            // 含花括号 → CLSID
            typelibParser_->loadByClsid(ref);
        } else {
            // 否则视为文件路径
            typelibParser_->loadByPath(ref);
        }
    }

    // 2. 自动加载常用TypeLib (可被--no-auto-typelib关闭)
    if (options.autoTypelib) {
        // 常用COM组件ProgID列表
        static const std::vector<std::string> commonProgIds = {
            "Scripting.FileSystemObject",   // Scripting Runtime
            "Scripting.Dictionary",         // Scripting Runtime (同一TypeLib)
            "ADODB.Connection",             // ADO
            "Excel.Application",            // Excel
            "Word.Application",             // Word
            "Shell.Application",            // Shell
            "WScript.Shell",                // WScript
            "MSXML2.DOMDocument",           // MSXML
        };

        for (const auto& progId : commonProgIds) {
            // 仅在缓存中不存在时加载; silent=true: 缺失是预期(用户未显式引用),
            // 不报VB4001避免污染c3-error.log
            if (!typelibParser_->findCachedCoClass(progId)) {
                typelibParser_->loadByProgId(progId, true);
            }
        }
    }

    // 2.5 Fix 160: ComLib= 免注册表收集
    // 只有 vbp 里 ComLib= 显式声明过的 DLL 进表 — 普通 Reference= / auto-typelib 加载的
    // typelib 一律不进。这是「未声明 ProgID 运行期零变化」的编译期半边: 表为空时
    // 产物不调 vb6_ComLibRegister, vb6_CreateObject 走原注册表路径, 一字不改。
    if (!comLibCanonMap_.empty()) {
        // canonical 键归一, 与 driver_compile 的 comLibCanonMap_ 键一致
        auto canonKey = [](const std::string& p) {
            std::string c = p;
            for (char& x : c) {
                if (x >= 'A' && x <= 'Z') x = (char)(x - 'A' + 'a');
                else if (x == '\\') x = '/';
            }
            return c;
        };

        // 跨 DLL 重复 ProgID 检测 (小写 ProgID → 首个声明它的 DLL)
        std::unordered_map<std::string, std::string> seenProgId;
        bool capWarned = false;
        for (auto& tl : typelibParser_->cachedResults()) {
            auto it = comLibCanonMap_.find(canonKey(tl->tlbPath));
            if (it == comLibCanonMap_.end())
                it = comLibCanonMap_.find(tl->canonPath);  // loadByPath 的 GetLongPathNameW 展开结果
            if (it == comLibCanonMap_.end()) continue;
            const std::string& relPath = it->second;

            for (auto& cc : tl->coclasses) {
                if (cc->progId.empty() || cc->clsidStr.empty()) continue;
                if (comLibRefs_.size() >= 256) {
                    if (!capWarned) {
                        diag_->warn(DiagnosticID::CodeGenUnsupportedFeature, SourceLocation{},
                                    "ComLib table full (256 entries); further components "
                                    "must be registry-registered");
                        capWarned = true;
                    }
                    continue;
                }
                std::string key = canonKey(cc->progId);
                auto ins = seenProgId.insert({key, relPath});
                if (!ins.second) {
                    // 两个 DLL 导出同一 ProgID: 病态但真实存在 (同名/同版本组件).
                    // 运行期按表序第一条命中即止, 会静默赢 — 必须在编译期告知.
                    diag_->warn(DiagnosticID::CodeGenUnsupportedFeature, SourceLocation{},
                                "Duplicate ProgID in ComLib components: " + cc->progId +
                                " in both " + ins.first->second + " and " + relPath +
                                "; first wins");
                }
                comLibRefs_.push_back({cc->progId, cc->clsidStr, cc->name, relPath});
            }
        }
    }

    if (options.verbose && !comLibRefs_.empty()) {
        std::cerr << "C3: ComLib= reg-free components: " << comLibRefs_.size() << "\n";
        for (const auto& e : comLibRefs_) {
            std::cerr << "  " << e[0] << "  " << e[1] << "  <- " << e[3] << "\n";
        }
    }

    // 2.6 Fix <vbeclipse> 2026-10-06: Object= 免注册 OCX 表收集
    // 只有 vbp 里 Object= 显式声明过的 OCX 进表 — 普通 Reference= / auto-typelib 加载的
    // typelib 一律不进。与 ComLib= 同构: 表为空时产物不调 vb6_OcxRefRegister,
    // 运行期 vb6_Form_ControlsAdd 走原注册表路径, 一字不改.
    if (!ocxCanonMap_.empty()) {
        auto canonKey = [](const std::string& p) {
            std::string c = p;
            for (char& x : c) {
                if (x >= 'A' && x <= 'Z') x = (char)(x - 'A' + 'a');
                else if (x == '\\') x = '/';
            }
            return c;
        };
        std::unordered_map<std::string, std::string> seenProgId;
        bool capWarned = false;
        for (auto& tl : typelibParser_->cachedResults()) {
            auto it = ocxCanonMap_.find(canonKey(tl->tlbPath));
            if (it == ocxCanonMap_.end())
                it = ocxCanonMap_.find(tl->canonPath);  // loadByPath 的 GetLongPathNameW 展开结果
            if (it == ocxCanonMap_.end()) continue;
            const std::string& relPath = it->second;
            for (auto& cc : tl->coclasses) {
                if (cc->progId.empty() || cc->clsidStr.empty()) continue;
                if (ocxLibRefs_.size() >= 256) {
                    if (!capWarned) {
                        diag_->warn(DiagnosticID::CodeGenUnsupportedFeature, SourceLocation{},
                                    "OCX reg-free table full (256 entries); further controls "
                                    "must be registry-registered");
                        capWarned = true;
                    }
                    continue;
                }
                std::string key = canonKey(cc->progId);
                auto ins = seenProgId.insert({key, relPath});
                if (!ins.second) {
                    diag_->warn(DiagnosticID::CodeGenUnsupportedFeature, SourceLocation{},
                                "Duplicate ProgID in OCX components: " + cc->progId +
                                " in both " + ins.first->second + " and " + relPath +
                                "; first wins");
                }
                ocxLibRefs_.push_back({cc->progId, cc->clsidStr, cc->name, relPath});
            }
        }
    }

    if (options.verbose && !ocxLibRefs_.empty()) {
        std::cerr << "C3: Object= reg-free OCX controls: " << ocxLibRefs_.size() << "\n";
        for (const auto& e : ocxLibRefs_) {
            std::cerr << "  " << e[0] << "  " << e[1] << "  <- " << e[3] << "\n";
        }
    }

    // 3. 输出加载结果 (verbose模式)
    if (options.verbose) {
        std::cerr << "C3: TypeLib import: ";
        int totalCoClasses = 0, totalIfaces = 0, totalModules = 0;
        for (auto& tl : typelibParser_->cachedResults()) {
            totalCoClasses += (int)tl->coclasses.size();
            totalIfaces += (int)tl->interfaces.size();
            totalModules += (int)tl->modules.size();
            // P24-04: verbose module details
            if (!tl->modules.empty()) {
                for (auto& m : tl->modules) {
                    std::cerr << "\n  MODULE: " << m->name << " dllPath=" << m->dllPath
                              << " funcs=" << m->functions.size() << " consts=" << m->constants.size();
                    for (auto& f : m->functions) {
                        std::cerr << "\n    func: " << f.realName << " memid=" << f.memid;
                    }
                }
            }
        }
        std::cerr << totalCoClasses << " coclasses, " << totalIfaces
                  << " interfaces, " << totalModules << " modules from "
                  << typelibParser_->cachedResults().size()
                  << " type libraries" << std::endl;
    }

    // P24-04: verbose - show each cached TypeLib info
    if (options.verbose) {
        for (auto& tl : typelibParser_->cachedResults()) {
            std::cerr << "  TL: " << tl->tlbPath << " cclasses=" << tl->coclasses.size() << " ifaces=" << tl->interfaces.size() << " mods=" << tl->modules.size() << std::endl;
        }
    }
    return true;  // TypeLib加载失败不阻断编译
}

} // namespace vb6c3
