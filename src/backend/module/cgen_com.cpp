#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_com.cpp: COM类工厂 + 接口vtable + RaiseEvent + 事件接收器 ---

void CCodeGen::emitClassFactory(Module& module) {
    std::string clsStruct = "vb6_cls_" + cIdent(moduleName_);  // Fix 013: VB_Name

    c_.emitBlank();
    c_.emitLine("// === 类工厂函数: " + module.moduleName + " ===");
    c_.emitBlank();

    // ai/084c: Class_Initialize 形参 → 带参工厂 _NewParams(实参)。
    // _New(void) 保留为默认值路径 —— 默认实例 (_Default)/COM 打包 (vb6_ComPack_*)/
    // 隐式 As New 等内部自动创建通路仍以零参 _New() 为契约, 不能因带参构造而消失。
    SubDecl* initSub = nullptr;
    for (auto& decl : module.declarations) {
        if (decl->kind == ASTNodeKind::SubDecl) {
            auto& sub = static_cast<SubDecl&>(*decl);
            if (sub.name == "Class_Initialize") { initSub = &sub; break; }
        }
    }
    const bool hasCtorParams = initSub && !initSub->params.empty();
    if (hasCtorParams) {
        for (auto& p : initSub->params) {
            if (!p->isByVal || p->isOptional) {
                diag_.error(DiagnosticID::CodeGenUnsupportedFeature, initSub->loc,
                    "带参构造 (Class_Initialize 带形参) 目前仅支持非 Optional 的 ByVal 参数");
                break;
            }
        }
    }

    // 工厂公共体: 分配 + iface/vtbl 初始化 + 字段默认值 + 事件指针 + Class_Initialize
    auto emitFactoryBody = [&](const std::string& ctorCallSuffix) {
    c_.emitLine(clsStruct + "* me = (" + clsStruct + "*)vb6_Alloc(sizeof(" + clsStruct + "));");
    c_.emitLine("if (!me) return NULL;");
    // ExeComBridge 01: 无条件初始化 (原先仅 DLL). __comObj 现在是所有类模块
    //   结构体的首字段 (见 cgen_base_generate_c_open.inc), EXE 的 _New() 与
    //   DLL 侧保持一致; 纯 EXE 下无人写读, 恒 NULL, 行为零变化.
    c_.emitLine("me->__comObj = NULL;  /* P6.6.3: no COM wrapper yet */");
    // Fix <vbeclipse> rev30: 把本实例登记进 RTL 的「裸工程类实例 → coclass 描述」表。
    //
    // 这张表**今天还没有读者** (账 #305 §B187 读数 B/G, 2026-09/10 两次普查都是 0 枚),
    // 所以下面那段"为什么必须有这一步"写的是一条**还没兑现的前提**, 不是既成事实:
    //   晚绑定调用点传进来的是**裸 `vb6_cls_X*`**, 首字段 `__comObj` 恒 NULL ⇒ 没有
    //   真 vtable ⇒ `vb6_ComIsDispatchable` 判否 ⇒ `vb6_getDispid` 返 DISPID_UNKNOWN
    //   ⇒ `vb6_ComCall` 的返回值永远是 **VT_EMPTY** —— 这一段实测成立;
    //   但"登记之后晚绑定就能答"不成立, 因为没人来查这张表 (实证: `.build/k81_run.bat`
    //   那台 EXE 夹具里 `Set v = o` 发成裸指针赋值、`v.Twice(21)` 至今空答)。
    // 订正一条旧说法: 先前记的"EXE 里没有 CoClass 块 ⇒ `_d` 恒 NULL ⇒ 工程内实例没进表"
    // 是**错的** —— EXE 也链一张完整 coclass 表 (`com_entry.c`, ExeComBridge 01), 收集判据
    // 是类的 Instancing (`cgen_util_dllentry_collect.inc`), 与有没有手写 CoClass 块无关。
    //
    // 2026-10-15 口径 (账 #305 §B187 第二格, 用户拍定): 不给"所有工程内类默认建 COM 元数据",
    // 而是**按工程类型分档** —— 只有 ActiveX 那一档 (`Type=DLL` / `Type=Control`, 后端读
    // `isDll_`, 与 `driver_compile.cpp` 里置位的那两条同一个口径) 才喂这张表; 标准 EXE /
    // OLE Exe 不付这份每实例一格的代价, 于是那条「裸实例 → desc」的读者接上之后也只服务
    // ActiveX 工程。非 ActiveX 撤掉的是**登记腿**, 类结构体首字段 `__comObj` 照旧无条件初始化。
    // 登记放在**工厂**里而不是别处: 每个实例必经此处; 而且 desc 按类模块名查
    // (`vb6_FindCoClassDesc("<模块名>")`), 不需要在类工厂里再存一份 desc。
    // 未进 coclass 表的类 (Instancing 不是 MultiUse/SingleUse、无 Public 成员等) 查不到
    // → 传 NULL → 不登记。
    if (isDll_) {
    c_.emitLine("{ const vb6_CoClassDesc* _d = vb6_FindCoClassDesc(\"" + cIdent(module.moduleName) + "\");");
    c_.emitLine("  if (_d) vb6_RegisterProjectClassInstance((void*)me, _d);  /* Fix <vbeclipse> rev30 */ }");
    }
    emitIfaceNewInit(module);  // tB Interface 契约 (B04): me->__iv_<I>.vt = &vb6_ivtbl_<I>_for_<C>
    emitClassVirtNewInit(module);  // tB Inherits (B08d): me->__cvtbl = &vb6_cvtbl_<本类>_impl

    // 初始化所有字段为默认值
    // 同时注册BSTR/Long类型成员到knownBstrVars_/knownLongVars_ (用于赋值时BSTR安全处理)
    for (VariableDecl* var_p : structFieldDecls(module)) {
        if (var_p) {
            auto& var = *var_p;
            std::string field = cIdent(var.name);
            // 注册到类型集合 (用于后续赋值时BSTR安全处理)
            std::string lower = var.name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            if (var.asType) {
                Vb6Type vtype = resolveArrayElemType(var.asType.get());
                if (vtype == Vb6Type::String) knownBstrVars_.insert(lower);
                else if (vtype == Vb6Type::Long || vtype == Vb6Type::Integer || vtype == Vb6Type::Boolean) knownLongVars_.insert(lower);
                else if (vtype == Vb6Type::Double || vtype == Vb6Type::Single) knownDoubleVars_.insert(lower);
            }
            if (var.isDynamicArray) {
                // 动态数组 `Dim x() As T`: ReDim 前不分配 (VB6 语义), NULL 起步
                c_.emitLine("me->" + field + " = NULL;");
            } else if (!var.dimensions.empty()) {
                // 定长数组成员: VB6 **每个实例创建时自动分配** (每个实例一份)。
                // 此前只置 NULL —— 类成员数组从未分配, Class_Initialize 里首次读写
                // (cDlg.cls 的 InitCustomColors) 解引用 NULL SafeArray → AV 读 0xC
                // (Task #44 SSTabEx "..." 按钮 cDlg.ShowColor 实测 0xC0000005)。
                // 模块级 (cgen_decl_var.cpp) / 过程局部 (cgen_localdecl.cpp) 都有
                // 创建逻辑, 唯类实例成员漏了 —— 补齐。Destroy 侧已有 vb6_SA_Destroy
                // (本文件 _Destroy 循环), 生命周期闭环。
                int dimCount44 = (int)var.dimensions.size();
                Vb6Type elemType44 = resolveArrayElemType(var.asType.get());
                std::string saElem44 = mapSaElemType(elemType44);
                std::string udtC44 = resolveArrayUdtElemCType(var.asType.get());
                if (dimCount44 == 1) {
                    auto& dim44 = var.dimensions[0];
                    std::string lb44 = "0", ub44 = "0";
                    if (dim44.lower) { emitExpr(*dim44.lower); lb44 = std::move(lastExpr_); }
                    if (dim44.upper) { emitExpr(*dim44.upper); ub44 = std::move(lastExpr_); }
                    if (!udtC44.empty()) {
                        // UDT 元素: 字段类型 SafeArray1D*, _Udt 版创建 (口径同
                        // cgen_localdecl.cpp:73), 销毁走 vb6_SA_Destroy 通用路径
                        c_.emitLine("me->" + field + " = vb6_SafeArrayReDim1D_Udt((int32_t)sizeof(" + udtC44 + "), " + lb44 + ", " + ub44 + ");");
                    } else {
                        c_.emitLine("me->" + field + " = vb6_SafeArrayCreate1D(" + saElem44 + ", " + lb44 + ", " + ub44 + ");");
                    }
                } else {
                    // 多维: 字段类型 SafeArrayND* (cgen_base_generate_c_open.inc:97)
                    std::string bounds44 = "_cbounds_" + field;
                    c_.emitLine("vb6_SafeArrayBound " + bounds44 + "[] = {");
                    c_.indent();
                    for (int d44 = 0; d44 < dimCount44; d44++) {
                        auto& dim44 = var.dimensions[d44];
                        std::string lb44 = "0", ub44 = "0";
                        if (dim44.lower) { emitExpr(*dim44.lower); lb44 = std::move(lastExpr_); }
                        if (dim44.upper) { emitExpr(*dim44.upper); ub44 = std::move(lastExpr_); }
                        std::string tr44 = (d44 < dimCount44 - 1) ? "," : "";
                        c_.emitLine("{" + lb44 + ", (" + ub44 + " - " + lb44 + " + 1)}" + tr44);
                    }
                    c_.dedent();
                    c_.emitLine("};");
                    if (!udtC44.empty()) {
                        c_.emitLine("me->" + field + " = vb6_SafeArrayReDimND_Udt((int32_t)sizeof(" + udtC44 + "), " + std::to_string(dimCount44) + ", " + bounds44 + ");");
                    } else {
                        c_.emitLine("me->" + field + " = vb6_SafeArrayCreateND(" + saElem44 + ", " + std::to_string(dimCount44) + ", " + bounds44 + ");");
                    }
                }
            } else if (var.asType) {
                Vb6Type fvt = resolveArrayElemType(var.asType.get());
                // Fix 084m: 字段C类型为指针 (void*/vb6_ComIface_X*/vb6_cls_X*/vb6_SafeArray1D*)
                // → 初始化为 NULL。resolveArrayElemType 对外部COM类型(ADODB.Connection等)返回
                // Variant, 若用 vb6_VariantEmpty() 初始化指针字段会 C2440
                // (cDataBase: me->Rs/Conn/pvWhereParams/Cmd = vb6_VariantEmpty())。
                std::string fieldCT = mapTypeRef(var.asType.get());
                if (!fieldCT.empty() && fieldCT.back() == '*') {
                    c_.emitLine("me->" + field + " = NULL;");
                } else if (fvt == Vb6Type::UserDefinedType) {
                    // Fix 038: UDT 字段不能用 = 0 初始化 (C2440), 改用 memset 零化
                    c_.emitLine("memset(&me->" + field + ", 0, sizeof(me->" + field + "));");
                } else {
                    c_.emitLine("me->" + field + " = " + defaultValue(fvt) + ";");
                }
            } else {
                // Fix 089j: VB6 无 As 类型声明的成员变量默认 Variant (C 字段
                // 类型 vb6_VARIANT), 不能 `= 0` 初始化 → C2440 (int→vb6_VARIANT),
                // 必须用 vb6_VariantEmpty() (与显式 `As Variant` 的 Tag 字段一致).
                c_.emitLine("me->" + field + " = vb6_VariantEmpty();");
            }
        }
    }

    // P6.5: 初始化事件接收器指针为NULL
    bool hasEvents = false;
    for (auto& decl : module.declarations) {
        if (decl->kind == ASTNodeKind::EventDecl) { hasEvents = true; break; }
    }
    if (hasEvents) {
        c_.emitLine("me->events = NULL;  /* P6.5: no event sink initially */");
    }

    // tB Inherits (B09): 祖先的 Class_Initialize **先**跑 (链上根→叶), 再跑自家那份。
    // 自家没有 Class_Initialize 时也要跑祖先的 —— 所以这条在下面的 initSub 判空之外。
    emitClassInitChain(module);

    // 084c: 自家 Class_Initialize —— 访问级别命名 + 带参构造 (_NewParams) 时透传实参
    if (initSub) {
        c_.emitLine(cProcName("Class_Initialize", initSub->access, isClassModule_ ? moduleName_ : "")
                    + "(me" + ctorCallSuffix + ");");
    }
    c_.emitLine("return me;");
};

    if (hasCtorParams) {
        // _New(void): 各形参默认值兜底 (内部自动创建路径产出"默认初始化"实例)
        c_.emitLine(clsStruct + "* " + clsStruct + "_New(void) {");
        c_.indent();
        std::string defaults;
        for (auto& p : initSub->params) {
            if (!defaults.empty()) defaults += ", ";
            Vb6Type pt = p->asType ? resolveArrayElemType(p->asType.get()) : Vb6Type::Variant;
            defaults += defaultValue(pt);
        }
        c_.emitLine("return " + clsStruct + "_NewParams(" + defaults + ");");
        c_.dedent();
        c_.emitLine("}");
        c_.emitBlank();
        // _NewParams: 真工厂, 形参映射与类方法 (makeParamList) 完全一致
        c_.emitLine(clsStruct + "* " + clsStruct + "_NewParams(" + makeParamList(initSub->params) + ") {");
        c_.indent();
        std::string argNames;
        for (auto& p : initSub->params) {
            if (!argNames.empty()) argNames += ", ";
            argNames += cIdent(p->name);
        }
        emitFactoryBody(", " + argNames);
        c_.dedent();
        c_.emitLine("}");
    } else {
        c_.emitLine(clsStruct + "* " + clsStruct + "_New(void) {");
        c_.indent();
        emitFactoryBody("");
        c_.dedent();
        c_.emitLine("}");
    }

    // _Destroy: 终止+释放
    // 账 #305 §B188 (第九十刀): terminate 那一步跑在自带帧里(`vb6_RunVbProtected`), 抛出来的号
    // 先接住, 等下面那三截尾巴(字段 BSTR / 数组字段 / vb6_Free(me))走完再交回外层帧。以前的做法
    // 是把整趟 destroyFunc 包进 ComObj_Release 的帧里(第八十刀), 那一趟 longjmp 会连尾巴一起跳掉:
    // 实测每枚抛错的出厂漏一份实例资源(x64 约 35 字节、x86 约 17 字节, 带一枚 String 字段约 120,
    // 出厂干净那枚恒 0; 尺 .build/k107_leak_probe.c / 读数 .build/k108_*.txt .build/k109_x64.txt)。
    // 出厂两腿都汇到 `_Destroy` 这一处(`instanceClaimRelease` 的末尾就直接调它), 所以边界装在这里
    // 是一份答案; 装在两条腿上就是两份。
    std::string destroyTermCall;
    for (auto& decl : module.declarations) {
        if (decl->kind == ASTNodeKind::SubDecl) {
            auto& sub = static_cast<SubDecl&>(*decl);
            if (sub.name == "Class_Terminate") {
                destroyTermCall = cProcName("Class_Terminate", sub.access,
                                            isClassModule_ ? moduleName_ : "");
                break;
            }
        }
    }
    const bool destroyHasTerm = !destroyTermCall.empty();
    c_.emitBlank();
    c_.emitLine("void " + clsStruct + "_Destroy(" + clsStruct + "* me) {");
    c_.indent();
    if (destroyHasTerm) {
        // 声明必须在 `if (!me) return;` 之前: 生成的 C 要按 C89 的"块首声明"排
        c_.emitLine("int32_t vb6_termErr = 0;  /* 账 305 B188: terminate 抛出来的号 */");
    }
    c_.emitLine("if (!me) return;");

    if (destroyHasTerm) {
        c_.emitLine("vb6_termErr = vb6_RunVbProtected((void(*)(void*))" + destroyTermCall +
                    ", me);  /* 账 305 B188: 体抛错也要把下面的尾巴走完 */");
    }

    // 释放BSTR字段
    for (VariableDecl* var_p : structFieldDecls(module)) {
        if (var_p) {
            auto& var = *var_p;
            Vb6Type varType = var.asType ? resolveArrayElemType(var.asType.get()) : Vb6Type::Variant;
            if (varType == Vb6Type::String) {
                c_.emitLine("vb6_BSTR_Free(me->" + cIdent(var.name) + ");");
            } else if (var.isDynamicArray || !var.dimensions.empty()) {
                c_.emitLine("if (me->" + cIdent(var.name) + ") vb6_SA_Destroy(me->" + cIdent(var.name) + ");");
            } else if (var.asType) {
                // 账 #305 §B218 (甲) 腿⑤: 字段那一枚引用**由字段自己握** —— 覆写(腿②)与置空(腿③)
                // 两站从第一刀起就问同一位主人, 只有出厂这一站缺着: 持有者拆掉时字段里那枚没人退,
                // 于是每枚被字段持有的实例永不走 Class_Terminate (漏, 不是崩)。
                // 类名问 `mapTypeRef` = 发这枚字段那一行用的同一个答案, 不开第二问;
                // 答出来不是 `vb6_cls_X*` 那一个形状 (As Object 是 void* / As Variant 是槽 /
                // 外部接口是裸指针) 就维持今天的"不动"。
                const std::string fieldCType = mapTypeRef(var.asType.get());
                const std::string prefix = "vb6_cls_";
                if (fieldCType.size() > prefix.size() + 1 &&
                    fieldCType.compare(0, prefix.size(), prefix) == 0 &&
                    fieldCType[fieldCType.size() - 1] == '*' &&
                    fieldCType.find(' ') == std::string::npos) {
                    const std::string fieldStruct = fieldCType.substr(0, fieldCType.size() - 1);
                    c_.emitLine("if (me->" + cIdent(var.name) + ") " + fieldStruct +
                                "_Release(me->" + cIdent(var.name) +
                                ");  /* 账 #305 (甲) 腿⑤: 字段那一枚引用出厂放手 */");
                }
            }
        }
    }

    c_.emitLine("vb6_Free(me);");
    if (destroyHasTerm) {
        // 尾巴拆完了, 号才交出去: COM 出厂那一腿由 ComObj_Release 的帧接住(照旧记一条带号诊断),
        // 进程内的 `Set x = Nothing` 那一腿由调用方自己的 On Error 接住 —— 两条都跟改前同形,
        // 区别只在实例这回真被拆掉了。
        c_.emitLine("if (vb6_termErr) vb6_RaiseError(vb6_termErr, vb6_ErrDescription());  /* 账 305 B188: 号在尾巴之后交回外层帧 */");
    }
    c_.dedent();
    c_.emitLine("}");

    // 账 #305 §B218 (甲, 第一刀) = 引用计数那一问的**唯一一位主人**: 每个类模块一对出入口,
    // 发在本类自己的 .c 里 (结构体完整类型可见), 声明在 .h (跨模块的槽要调它)。
    // 为什么是一对导出函数而不是每条腿各写三行: 这条判断有四个分支 (空实例 / 还有引用 /
    // 被 __comObj 接走 / 真该拆), 五腿各 inline 一遍 = 五位主人 (#167/#231/#245/#298/#443 那一族)
    // —— 今天那三处手抄 (接口 AddRef / Release / claim 桩) 就在这一刀改成转调它。
    c_.emitBlank();
    c_.emitLine("unsigned long " + clsStruct + "_AddRef(void* instance) {");
    c_.indent();
    c_.emitLine(clsStruct + "* me = (" + clsStruct + "*)instance;");
    c_.emitLine("if (!me) return 0UL;");
    c_.emitLine("me->__refcount += 1;");
    c_.emitLine("return (unsigned long)me->__refcount;");
    c_.dedent();
    c_.emitLine("}");
    c_.emitBlank();
    c_.emitLine("unsigned long " + clsStruct + "_Release(void* instance) {");
    c_.indent();
    c_.emitLine(clsStruct + "* me = (" + clsStruct + "*)instance;");
    c_.emitLine("if (!me) return 0UL;");
    c_.emitLine("me->__refcount -= 1;");
    c_.emitLine("if (me->__refcount > 0) return (unsigned long)me->__refcount;  /* 还有人握 */");
    // 实例已被 COM 包装器接管时 (__comObj 非空) 销毁权在包装器 (P6/B13 统一两套计数)
    c_.emitLine("if (me->__comObj != NULL) return 0UL;  /* wrapper owns teardown */");
    c_.emitLine(clsStruct + "_Destroy(me);  /* 0 引用: Class_Terminate + 释放 */");
    c_.emitLine("return 0UL;");
    c_.dedent();
    c_.emitLine("}");

    // ExeComBridge 03: COM 实参打包函数 (每个类模块一个).
    //   `New <本工程类>` 出现在 COM 调用实参位置时 (cgen_util_com.cpp comPackExpr),
    //   通用打包 vb6_ComPackValue 会把裸 vb6_cls_X* 当 VT_DISPATCH 传出去, 对端
    //   AddRef 时把结构体首字段当 vtable → 0xC0000005
    //   (VBMAN_DEMO: .Router.Reg "Demo", New bHello).
    //   本函数先经 __comObj 包装成真 IDispatch 再转 VARIANT (RTL:
    //   vb6_ComPackVB6InstanceRaw). 类名用原始名 —— desc 表的 classVariable 即
    //   VB_Name (cgen_util_dllentry_collect.inc), 查找按 _stricmp 大小写不敏感.
    c_.emitBlank();
    c_.emitLine("// ExeComBridge 03: 工程类实例 → COM 实参 (先包装成 IDispatch 再转 VARIANT)");
    c_.emitLine("void* vb6_ComPack_" + cIdent(moduleName_) + "(void* instance) {");
    c_.indent();
    c_.emitLine("return vb6_ComPackVB6InstanceRaw(\"" + moduleName_ + "\", instance);");
    c_.dedent();
    c_.emitLine("}");
}

// 账 #305 §B218 (甲) 腿②③: 「覆写 / 置空一枚 typed 工程类槽」收成一位主人。
// 为什么是一条函数而不是站点各写三行: 「这一格该不该放手」的判断跟着上面那对出入口走,
// 站点抄一遍就多一位主人 (#167/#231/#245/#298/#443 那一族)。valueExpr 传空串 = 置空。
// 顺序 = **先 claim 新、再 release 旧** —— 接口那族 emitIvrefQiSet 撞过的同一件事:
// 先放旧会让 `Set a = a` 在自己还没换掉之前把自己拆掉。
// rhsTransferred = 右侧那枚引用**已经带着对价交给这一格了**: 两种成因 (`New <类>` 与工程内过程
// 的调用, 见 rhsTransfersProjectRef) ⇒ 移交、不 claim。
void CCodeGen::emitClassSlotRefSet(const std::string& target, const std::string& className,
                                   const std::string& valueExpr, bool rhsTransferred, bool outSlot) {
    const std::string clsStruct = "vb6_cls_" + cIdent(className);
    if (outSlot) {
        // 腿②' = 写进一枚**活在本过程之外**的槽 (ByRef 出参那座桥的写回站)。这里的判断只有一条
        // 加法: 取走一枚引用。**不退旧值** —— 那种槽今天整个在计数系统外面, 退它 = 多退一次;
        // 「出参写回该不该放手旧值」与下标槽是同一格, 在册等腿⑤ 一起答。
        c_.emitLine("{  /* 账 #305 (甲) 腿②': Set 出参槽 " + target + " */");
        c_.indent();
        c_.emitLine(target + " = " + valueExpr + ";");
        if (!rhsTransferred) {
            c_.emitLine("if (" + target + ") " + clsStruct + "_AddRef(" + target + ");");
        }
        c_.dedent();
        c_.emitLine("}");
        return;
    }
    const std::string legTag = valueExpr.empty() ? "③" : "②";
    const std::string oldSlot = "__clsold" + std::to_string(ivrefSetSeq_++);
    c_.emitLine("{  /* 账 #305 (甲) 腿" + legTag + ": Set " + target + " */");
    c_.indent();
    c_.emitLine(clsStruct + "* " + oldSlot + " = " + target + ";");
    if (!valueExpr.empty()) {
        c_.emitLine(target + " = " + valueExpr + ";");
        if (!rhsTransferred) {
            c_.emitLine("if (" + target + ") " + clsStruct + "_AddRef(" + target + ");");
        }
    } else {
        c_.emitLine(target + " = NULL;  /* Set Nothing (项目类实例) */");
    }
    c_.emitLine("if (" + oldSlot + ") " + clsStruct + "_Release(" + oldSlot + ");");
    c_.dedent();
    c_.emitLine("}");
}

// 账 #305 §B218 (甲) 刀3: 「这一发右值是不是自带一枚移交过来的引用」—— 上面那位主人「不 claim」
// 那条臂的第二个成因, 与 `New` 同一条契约, 所以只在这一处问, 不在站点各抄一遍。
// 契约的实物 (三台工程同形, 尺 `.build/k1081_proj_emit.c`): 被调方 `Set <返回槽> = <局部/字段>`
// 走 腿② 取走一枚, 出口 腿④ 退的是它自己那枚**局部**, 而返回槽自带那道夹不放手 ⇒ 交出来的指针
// 本身带一枚对价引用; 调用方再 claim 就是每调用净 +1, 那一枚永远没人退 (实测 W6 / W9 `terms=0`)。
// 刻意只认**裸标识符 callee** 且模块符号表认到 Function / Property Get、返回类型是工程 Class:
// 带接收者的成员调用 (`obj.Method()`) 走的是另一套契约 (COM 打包 / 视图 QI), 本刀不碰。
bool CCodeGen::rhsTransfersProjectRef(const ASTNode& expr) {
    // 账 #305 §B218 (甲) 刀3 + 刀4: 「这一发右值自带一枚移交过来的引用」= 上面那位主人
    // 「不 claim」那条臂的**唯一**问句, 三条臂都收在这里 (别处再答一遍 = 第二套判据):
    //   形① `Set c = MakeOne()`               callee 是裸标识符 (刀3);
    //   形② `Set rb = d.RefOf()`              callee 是成员访问 (刀4);
    //   形③ `Set p = Peer` / `Set p = MyBase.Peer`  **无括号**的属性读 = 语义上就是一次调用 (刀4)。
    // 前提在三台工程上逐行核过 (尺 `.build/k1115_membercensus.py` + cls_inh 的 `--emit-c`):
    // 被调方在**自己的返回槽**那一格 claim 一枚, 出口 腿④ 退的是那枚局部、返回槽自带那道夹不放手,
    // 继承那几档包装器 `return vb6_<Base>_M((vb6_cls_<Base>*)me);` 原样转交不退 ⇒ 交出来的指针
    // 本身就带对价 ⇒ 调用方再 claim = 每调用净 +1 = 永不销毁。
    // 反面 = 数据字段: 那一枚仍由持有者握着, 接手会少一枚, 所以字段那一档继续 claim。
    const ASTNode* callee = nullptr;
    if (expr.kind == ASTNodeKind::IndexOrCallExpr) {
        callee = static_cast<const IndexOrCallExpr&>(expr).callee.get();
    } else if (expr.kind == ASTNodeKind::MemberAccessExpr) {
        callee = &expr;  // 形③: 整枚右值就是一次属性读
    }
    if (!callee) return false;

    if (callee->kind == ASTNodeKind::IdentifierExpr) {
        // 形① (刀3 那一臂, 一字未动)
        const std::string name = static_cast<const IdentifierExpr&>(*callee).name;
        const Symbol* fn = symTab_.lookupModule(name);
        if (!fn || (fn->kind != SymbolKind::Function && fn->kind != SymbolKind::PropertyGet)
            || fn->variableTypeName.empty()) {
            return false;
        }
        return !projectClassNameOf(fn->variableTypeName).empty();
    }
    if (callee->kind != ASTNodeKind::MemberAccessExpr) return false;
    const auto& ma = static_cast<const MemberAccessExpr&>(*callee);
    if (!ma.object) return false;

    // 接收者那一问只有两条答案: 普通表达式问那一位类型推断主人; `MyBase` 这一形按
    // cgen_expr_member_precheck.inc 那条**按名接管**的事实认 (parser 不动), 基名只问链表
    // 那一张表 (classViewByName), 不现拼第二套。
    std::string recvCls;
    const bool myBaseRecv =
        ma.object->kind == ASTNodeKind::IdentifierExpr
        && Symbol::toLower(static_cast<const IdentifierExpr&>(*ma.object).name) == "mybase";
    if (myBaseRecv) {
        const ClassChainView* self = classViewByName(moduleName_);
        const ClassChainView* base =
            (self && !self->baseKey.empty()) ? classViewByName(self->baseKey) : nullptr;
        if (!base || !base->mod) return false;
        recvCls = base->name;
    } else {
        recvCls = inferClassTypeOfExpr(*ma.object);
    }
    if (recvCls.empty()) return false;

    // 字段那一档不是调用 (这张表就是发那行字段声明用的同一位主人; `COM:` 前缀那几档同样
    // 算字段 —— 它们压根不是 vb6_cls_ 实例)。
    if (classTypedFieldMap_) {
        auto itCls = classTypedFieldMap_->find(recvCls);
        if (itCls != classTypedFieldMap_->end()
            && itCls->second.count(Symbol::toLower(ma.memberName)) != 0) {
            return false;
        }
    }
    // 过程返回类: 只认 Function / PropertyGet 符号, 与形① 同一位主人。返回的是工程内**接口**
    // 那一档刻意不走这条腿 (薄指针不是实例, 它的所有权由接口那一位主人管)。
    const std::string retCls = getClassMethodReturnType(recvCls, ma.memberName);
    if (retCls.empty() || !ivrefCType(retCls).empty()) return false;
    return true;
}

// P6.4+: 类模块默认实例 (VB_PredeclaredId=True) 惰性单例访问器.
// VB6 为 PredeclaredId 类生成隐藏的全局默认实例, frm 里裸类名成员访问经过
// knownClassVars_ 解析, 对象表达式生成 vb6_cls_X_Default() 调用. 这里生成:
//   static vb6_cls_cTT* g_p_vb6_cls_cTT_Default = NULL;
//   vb6_cls_cTT* vb6_cls_cTT_Default(void) { ... _New() on demand ... }
void CCodeGen::emitClassDefaultInstance(Module& module) {
    if (defaultInstanceClassName(moduleName_).empty()) return;
    std::string clsStruct = "vb6_cls_" + cIdent(moduleName_);
    c_.emitBlank();
    c_.emitLine("// 默认实例 (VB_PredeclaredId=True): 惰性创建共享单例");
    c_.emitLine("static " + clsStruct + "* g_p_" + clsStruct + "_Default = NULL;");
    c_.emitLine(clsStruct + "* " + clsStruct + "_Default(void) {");
    c_.indent();
    c_.emitLine("if (!g_p_" + clsStruct + "_Default)");
    c_.emitLine("    g_p_" + clsStruct + "_Default = " + clsStruct + "_New();");
    c_.emitLine("return g_p_" + clsStruct + "_Default;");
    c_.dedent();
    c_.emitLine("}");
    c_.emitBlank();
}

// ============================================================
// Fix 099: Public 字段的 COM 读写访问器 (ActiveX DLL)
// ============================================================
// 真 VB6 把类模块的 `Public X As T` 暴露为 property get/let 对 (TypeLib 里同为
// Property). 此前 C3 完全不暴露: cHttpServer 的 Router/RouteBefore/RouteAfter/
// SSE/Database/Statistics 等 Public 字段既不进 IDispatch 方法表也不进 TypeLib,
// 客户端 GetIDsOfNames("Router") 直接失败 → vb6_ComGetProp 返回 NULL → 调用方
// 拿 NULL 当对象解引用 → 0xC0000005 (段错误).
//
// 访问器必须生成在**类模块自己的 .c** 里: dll_entry.c 只有 `struct vb6_cls_X;`
// 前向声明, 属不完整类型, 无法 `me->Fld`. 命名 vb6_cls_<cls>_field_get_<fld> /
// _field_let_<fld>, 由 dll_entry.c 的字段桥接 extern 调用.
//
// 只暴露能直接映射到 vb6_VARIANT 的字段类型; Variant / UDT / 数组 / 外部 COM
// 接口指针字段一律不生成 (宁缺勿错 —— 暴露了但类型映射不对会写出编译不过或
// 语义错误的 C). 收集条件与 semantic_analyzer 的 publicFieldNames 严格一致.
void CCodeGen::emitClassFieldAccessors(Module& module) {
    // ExeComBridge 01: 放开 EXE (原先 if (!isDll_ || ...) return).
    //   EXE 工程也生成 com_entry.c (driver_codegen_dll_entry.inc), 其字段桥接
    //   extern 引用 _field_get_/_let_ —— 访问器实体必须无条件生成, 否则
    //   Instancing 非 Private 且有 Public 字段的 EXE 类 LNK2019.
    //   依赖的 RTL 符号 (vb6_ComObject_FromInstance 等) 由 vb6comserver_obj.c
    //   提供, driver_link.cpp 已无条件链入; 无 COM 引用时为死代码, 链接器剔除.
    if (!isClassModule_) return;

    const std::string clsStruct = "vb6_cls_" + cIdent(moduleName_);

    struct FieldEmit {
        std::string name;                    // 声明原名
        std::vector<std::string> getLines;   // getter 体 (写 r)
        std::vector<std::string> letLines;   // setter 体 (读 value 写 me->fld)
    };
    std::vector<FieldEmit> fields;

    for (auto& decl : module.declarations) {
        if (decl->kind != ASTNodeKind::VariableDecl) continue;
        auto& var = static_cast<VariableDecl&>(*decl);
        // 与 semantic_analyzer.cpp 的 publicFieldNames 收集条件一致 (Fix 187: 含 As New)
        if (var.access != AccessLevel::Public) continue;
        if (var.isWithEvents || var.isDynamicArray || !var.dimensions.empty()) continue;

        const std::string fld = cIdent(var.name);
        const std::string fieldCT = mapTypeRef(var.asType.get());
        const Vb6Type fvt = var.asType ? resolveArrayElemType(var.asType.get()) : Vb6Type::Variant;

        FieldEmit fe;
        fe.name = var.name;

        const bool isClassPtr = fieldCT.size() > 9
            && fieldCT.compare(0, 8, "vb6_cls_") == 0
            && fieldCT[fieldCT.size() - 1] == '*';

        // Fix 187: As New 字段. 真 VB6 把它暴露为 PropertyGet/PutRef; 字段 C 侧是
        // void* (惰性实例化的裸实例指针), 所以 getter 必须自己补实例化, 否则客户端
        // 首次读取拿到空对象. 只有"工程内 MultiUse/SingleUse 类"与"外部 COM 类"
        // 两种目标能安全映射; 其余 (内建集合/未知类型) 退化为空体存根, 宁缺勿错.
        if (var.isNew && var.asType && var.asType->kind == ASTNodeKind::SimpleTypeRef) {
            std::string nTarget = static_cast<SimpleTypeRef*>(var.asType.get())->name;
            size_t nDot = nTarget.find_last_of('.');
            if (nDot != std::string::npos) nTarget = nTarget.substr(nDot + 1);
            Symbol* nSym = lookupModuleDotted(nTarget);
            // 账 #305 §B187 第二腿: 判据问 Symbol::isCoClassTableEntry() (与建表那侧、
            // 装箱腿那份名单同一把尺子) —— 这里发的正是
            // vb6_ComObject_FromBorrowedInstance(vb6_FindCoClassDesc(...))，类不在表里
            // 就会装箱出 NULL 包装器 (= Nothing)。
            const bool nIsCoclass = nSym && nSym->isCoClassTableEntry();
            if (nIsCoclass) {
                const std::string tId = cIdent(nSym->name);
                fe.getLines.push_back("if (!me->" + fld + ") me->" + fld + " = (void*)vb6_cls_"
                                      + tId + "_New();  /* As New auto-instantiate */");
                fe.getLines.push_back("r = vb6_VariantObject(vb6_ComObject_FromBorrowedInstance("
                                      "vb6_FindCoClassDesc(\"" + tId + "\"), (void*)me->" + fld + "));");
                fe.letLines.push_back("me->" + fld + " = (value.vt == vb6_vtDispatch) ? "
                                      "vb6_ComObject_GetInstance(value.pdispVal) : NULL;");
            } else if (nSym && nSym->kind == SymbolKind::ComClass) {
                const std::string nProgId = nSym->comProgId.empty() ? nSym->name : nSym->comProgId;
                fe.getLines.push_back("if (!me->" + fld + ") me->" + fld
                                      + " = (void*)vb6_NewObject(L\"" + nProgId
                                      + "\");  /* As New auto-instantiate */");
                fe.getLines.push_back("r = vb6_VariantObject((void*)me->" + fld + ");");
                fe.letLines.push_back("me->" + fld + " = (value.vt == vb6_vtDispatch) ? "
                                      "value.pdispVal : NULL;");
            }
            fields.push_back(fe);
            continue;
        }

        if (isClassPtr) {
            // (a) 工程类实例指针字段 → VT_DISPATCH: 裸实例交 RTL 包装成 IDispatch
            // Fix 188: 借出型包装 —— 实例归宿主类所有, 客户端释放包装器不能销毁它.
            std::string target = (var.asType && var.asType->kind == ASTNodeKind::SimpleTypeRef)
                ? static_cast<SimpleTypeRef*>(var.asType.get())->name : std::string();
            size_t dot = target.find_last_of('.');
            if (dot != std::string::npos) target = target.substr(dot + 1);
            if (target.empty()) continue;
            fe.getLines.push_back("r = vb6_VariantObject(vb6_ComObject_FromBorrowedInstance("
                                  "vb6_FindCoClassDesc(\"" + target + "\"), (void*)me->" + fld + "));");
            fe.letLines.push_back("me->" + fld + " = (" + fieldCT + ")("
                                  "(value.vt == vb6_vtDispatch) ? "
                                  "vb6_ComObject_GetInstance(value.pdispVal) : NULL);");
        } else if (fvt == Vb6Type::String) {
            // (b) 字符串 → VT_BSTR, 必须深拷贝 (BSTR 所有权归实例)
            fe.getLines.push_back("vb6_BSTR_Assign(&r.bstrVal, me->" + fld + ");");
            fe.getLines.push_back("r.vt = vb6_vtBSTR;");
            fe.letLines.push_back("vb6_BSTR_Assign(&me->" + fld + ", "
                                  "(value.vt == vb6_vtBSTR) ? value.bstrVal : NULL);");
        } else if (fvt == Vb6Type::Boolean) {
            // (c) 布尔 → VT_BOOL (VB6 True = -1)
            fe.getLines.push_back("r = vb6_VariantBool(me->" + fld + " ? -1 : 0);");
            fe.letLines.push_back("me->" + fld + " = (value.lVal != 0) ? -1 : 0;");
        } else if (fvt == Vb6Type::Integer) {
            // 账 #305 §B184: 写回收到的 `value.lVal` 恒是 int32, 存进比它窄的这一档
            // 以前发的是裸 C cast ⇒ 40000 静默变 -25536, 而真 VB6 交 error 6 且保住
            // 旧值。范围闸只从 narrowCheckFnForVbType 那一处答 (改前这条链根本不知道
            // 有那张表); 闸的返回类型正是字段类型 (int16_t / uint8_t) ⇒ cast 一并撤掉。
            // ⚠ Long / Boolean 两臂刻意不动: 前者源本就 32 位不截, 后者是归一不是截断。
            fe.getLines.push_back("r = vb6_VariantInt(me->" + fld + ");");
            fe.letLines.push_back("me->" + fld + " = "
                                  + narrowCheckFnForVbType(fvt, nullptr) + "(value.lVal);");
        } else if (fvt == Vb6Type::Byte) {
            fe.getLines.push_back("r = vb6_VariantByte(me->" + fld + ");");
            fe.letLines.push_back("me->" + fld + " = "
                                  + narrowCheckFnForVbType(fvt, nullptr) + "(value.lVal);");
        } else if (fvt == Vb6Type::Long || fieldCT == "int32_t") {
            // (d) Long / 枚举 (mapTypeRef 对 EnumType 返回 int32_t, 底层即 Long)
            fe.getLines.push_back("r = vb6_VariantLong(me->" + fld + ");");
            fe.letLines.push_back("me->" + fld + " = (int32_t)value.lVal;");
        } else if (fvt == Vb6Type::Double) {
            fe.getLines.push_back("r = vb6_VariantDouble(me->" + fld + ");");
            fe.letLines.push_back("me->" + fld + " = (double)value.dblVal;");
        } else if (fvt == Vb6Type::Single) {
            fe.getLines.push_back("r.vt = vb6_vtSingle;");
            fe.getLines.push_back("r.fltVal = me->" + fld + ";");
            fe.letLines.push_back("me->" + fld + " = (float)value.dblVal;");
        } else if (fvt == Vb6Type::Date) {
            fe.getLines.push_back("r.vt = vb6_vtDate;");
            fe.getLines.push_back("r.dblVal = me->" + fld + ";");
            fe.letLines.push_back("me->" + fld + " = (double)value.dblVal;");
        } else if (fvt == Vb6Type::Currency) {
            fe.getLines.push_back("r.vt = vb6_vtCurrency;");
            fe.getLines.push_back("r.cyVal = me->" + fld + ";");
            fe.letLines.push_back("me->" + fld + " = (int64_t)value.cyVal;");
        } else if (fvt == Vb6Type::Object || fieldCT == "void*") {
            // (e) 外部 COM 对象 / As Object 字段 → VT_DISPATCH. 字段里存的就是接口
            //     指针 (Dictionary/Collection 等), 直接传出, 不做包装.
            fe.getLines.push_back("r = vb6_VariantObject((void*)me->" + fld + ");");
            fe.letLines.push_back("me->" + fld + " = (value.vt == vb6_vtDispatch) ? value.pdispVal : NULL;");
        }
        // 说明: 判不出可映射 C 类型的字段 (UDT / 接口值类型 / Variant / 未知外部类型)
        // 不跳过, 而是生成**空体存根** —— getter 返回 r 的初值 (Empty), setter 忽略
        // 写入. 必须生成定义的原因: dll_entry.c 在另一个编译单元里为 publicFieldNames
        // 无条件建桥接 (那边看不到这里的类型判定), 少一个定义就是 LNK2019 (实测 96 个).
        fields.push_back(fe);
    }

    if (fields.empty()) return;

    c_.emitBlank();
    c_.emitLine("// === Fix 099: Public 字段的 COM 访问器 (Property Get/Let) ===");
    c_.emitBlank();

    for (auto& fe : fields) {
        const std::string fldId = cIdent(fe.name);

        // getter: 返回 vb6_VARIANT (与 dll_entry.c 桥接的 result 缓冲布局兼容)
        c_.emitLine("vb6_VARIANT " + clsStruct + "_field_get_" + fldId + "(" + clsStruct + "* me) {");
        c_.indent();
        c_.emitLine("vb6_VARIANT r = vb6_VariantEmpty();");
        c_.emitLine("if (!me) return r;");
        for (auto& ln : fe.getLines) c_.emitLine(ln);
        c_.emitLine("return r;");
        c_.dedent();
        c_.emitLine("}");

        // setter: 从 vb6_VARIANT 取值写入字段
        c_.emitLine("void " + clsStruct + "_field_let_" + fldId + "(" + clsStruct + "* me, vb6_VARIANT value) {");
        c_.indent();
        c_.emitLine("if (!me) return;");
        for (auto& ln : fe.letLines) c_.emitLine(ln);
        c_.dedent();
        c_.emitLine("}");
        c_.emitBlank();
    }
}

// ============================================================
// P6.4: 接口 vtable + 包装类型生成 (Implements 代码生成)
// ============================================================

void CCodeGen::emitInterfaceVtable(Module& module) {
    std::string clsStruct = "vb6_cls_" + cIdent(moduleName_);  // Fix 013: VB_Name

    for (auto& impl : module.implements) {
        const std::string& ifaceName = impl->interfaceName;
        // 登记在册的接口 (tB `Interface` 块, 或 Pass A2 认进来的 VB6 .cls 宿主) 由
        // `emitIfaceImplTables` 那条 canonical 路径发: 槽序由接口自己定, `me` 是统一的
        // `vb6_ivref_I*`, 表结构带 guard 只发一份。这里按实现类 harvest 再发一遍就是同名
        // 结构 C2011 撞车 (实测 290 次), 所以只发登记表里查不到的接口。
        if (ivLookupIface(ifaceName)) continue;
        std::string ifaceId = cIdent(ifaceName);

        // 收集接口方法信息: 从实现类中查找 IFoo_MethodName 方法
        struct IfaceMethodInfo {
            std::string methodName;   // 原始方法名 (如 "Bar")
            std::string implFuncName; // 实现函数C名 (如 "vb6_Class1_IFoo_Bar")
            std::string retType;      // 返回C类型
            std::string params;       // 参数列表 (不含me, 如 "int32_t x")
            bool isSub;               // Sub vs Function
        };
        std::vector<IfaceMethodInfo> methods;

        for (auto& decl : module.declarations) {
            if (decl->kind == ASTNodeKind::SubDecl) {
                auto& sub = static_cast<SubDecl&>(*decl);
                // 检查是否为 Implements 实现方法 (IFoo_MethodName 格式)
                if (sub.name.size() > ifaceName.size() + 1 &&
                    Symbol::toLower(sub.name.substr(0, ifaceName.size() + 1)) ==
                    Symbol::toLower(ifaceName + "_")) {
                    std::string methodName = sub.name.substr(ifaceName.size() + 1);
                    std::string params = makeParamList(sub.params);
                    methods.push_back({methodName, cProcName(sub.name, sub.access, isClassModule_ ? moduleName_ : ""),
                                       "void", params, true});
                }
            } else if (decl->kind == ASTNodeKind::FunctionDecl) {
                auto& func = static_cast<FunctionDecl&>(*decl);
                if (func.name.size() > ifaceName.size() + 1 &&
                    Symbol::toLower(func.name.substr(0, ifaceName.size() + 1)) ==
                    Symbol::toLower(ifaceName + "_")) {
                    std::string methodName = func.name.substr(ifaceName.size() + 1);
                    std::string params = makeParamList(func.params);
                    std::string retType = mapTypeRef(func.returnType.get());
                    methods.push_back({methodName, cProcName(func.name, func.access, isClassModule_ ? moduleName_ : ""),
                                       retType, params, false});
                }
            } else if (decl->kind == ASTNodeKind::PropertyDecl) {
                auto& prop = static_cast<PropertyDecl&>(*decl);
                if (prop.name.size() > ifaceName.size() + 1 &&
                    Symbol::toLower(prop.name.substr(0, ifaceName.size() + 1)) ==
                    Symbol::toLower(ifaceName + "_")) {
                    std::string methodName = prop.name.substr(ifaceName.size() + 1);
                    std::string params = makeParamList(prop.params);
                    // Property Get → Function, Property Let/Set → Sub
                    // 使用prop_get_/prop_let_/prop_set_前缀
                    std::string propPrefix;
                    if (prop.propKind == ProcKind::PropertyGet) {
                        propPrefix = "prop_get_";
                        std::string retType = mapTypeRef(prop.returnType.get());
                        methods.push_back({methodName, cProcName(propPrefix + prop.name, prop.access, isClassModule_ ? moduleName_ : ""),
                                           retType, params, false});
                    } else {
                        propPrefix = (prop.propKind == ProcKind::PropertySet) ? "prop_set_" : "prop_let_";
                        methods.push_back({methodName, cProcName(propPrefix + prop.name, prop.access, isClassModule_ ? moduleName_ : ""),
                                           "void", params, true});
                    }
                }
            }
        }

        if (methods.empty()) continue;

        // 1. 生成 vtable 结构体 (函数指针表)
        std::string vtblName = "vb6_vtbl_" + ifaceId;
        h_.emitLine("// Interface vtable: " + ifaceName);
        h_.emitLine("typedef struct " + vtblName + " {");
        for (auto& m : methods) {
            std::string paramList = classMeParam();
            if (m.params != "void") {
                paramList += ", " + m.params;
            }
            h_.emitLine("    " + m.retType + " (*" + cIdent(m.methodName) + ")(" + paramList + ");");
        }
        h_.emitLine("} " + vtblName + ";");
        h_.emitBlank();

        // 2. 生成接口引用包装类型 (vtable指针 + 对象指针)
        std::string ifaceTypeName = "vb6_iface_" + ifaceId;
        h_.emitLine("typedef struct " + ifaceTypeName + " {");
        h_.emitLine("    " + vtblName + "* vtbl;");
        h_.emitLine("    void* obj;");
        h_.emitLine("} " + ifaceTypeName + ";");
        h_.emitBlank();

        // 3. 生成全局 vtable 实例 (指向实现类的接口方法)
        std::string vtblInstance = vtblName + "_for_" + cIdent(moduleName_);  // Fix 013: VB_Name
        c_.emitBlank();
        c_.emitLine("// Interface vtable instance: " + ifaceName + " for " + module.moduleName);
        c_.emitLine("static " + vtblName + " " + vtblInstance + " = {");
        c_.indent();
        for (size_t i = 0; i < methods.size(); i++) {
            std::string entry = "." + cIdent(methods[i].methodName) + " = " + methods[i].implFuncName;
            if (i < methods.size() - 1) entry += ",";
            c_.emitLine(entry);
        }
        c_.dedent();
        c_.emitLine("};");

        // 4. 生成包装函数: vb6_iface_IFoo_wrap(obj) → 创建接口引用
        h_.emitLine(ifaceTypeName + " " + ifaceTypeName + "_wrap(" + clsStruct + "* obj);");
        h_.emitBlank();
        c_.emitBlank();
        c_.emitLine(ifaceTypeName + " " + ifaceTypeName + "_wrap(" + clsStruct + "* obj) {");
        c_.indent();
        c_.emitLine(ifaceTypeName + " iface;");
        c_.emitLine("iface.vtbl = &" + vtblInstance + ";");
        c_.emitLine("iface.obj = obj;");
        c_.emitLine("return iface;");
        c_.dedent();
        c_.emitLine("}");
    }
}


// ============================================================
// P6.6: ActiveX DLL代码生成
// ============================================================


} // namespace vb6c3

