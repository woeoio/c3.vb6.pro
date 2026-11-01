#include "backend/cgen.hpp"
#include "common/event_param_abi.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

namespace {
// 账 #305 第二百五十六刀: 「这一串递出去能不能取地址」—— 认 标识符 / 成员链(. 与 ->) /
// 尾随一层下标 `[…]`；其余（带括号的调用、逗号、字面量、类型转换…）一律答否，
// 由调用方落一枚临时量。刻意保守：答否只是少一次直递，答错才会发出非法 C。
bool evtArgIsLvalue(const std::string& s) {
    size_t p = 0;
    auto eatChar = [&s, &p](bool first) {
        if (p >= s.size()) return false;
        const unsigned char c = static_cast<unsigned char>(s[p]);
        const bool ok = first ? (isalpha(c) || c == '_' || c >= 0x80)
                              : (isalnum(c) || c == '_' || c >= 0x80);
        if (ok) p++;
        return ok;
    };
    auto eatIdent = [&]() {
        const size_t st = p;
        if (!eatChar(true)) return false;
        while (eatChar(false)) {}
        return p > st;
    };
    if (!eatIdent()) return false;
    while (p < s.size()) {
        if (s.compare(p, 2, "->") == 0) { p += 2; if (!eatIdent()) return false; continue; }
        if (s[p] == '.') { p += 1; if (!eatIdent()) return false; continue; }
        if (s[p] == '[') {
            const size_t close = s.find(']', p);
            return (close != std::string::npos && close + 1 == s.size());
        }
        return false;
    }
    return true;
}

// 账 #305 第三百八刀: 「处理器写回的那枚槽怎么搬回实参自己的存储」—— 只有声明档搬得回去的
// 那几形有答案，出口一律问 src/rtl/core/vb6rtl/vb6rtl_variant.h 那一族 To{String,Bool,Long,
// LongPtr,Double}。串那一发用 AssignMove: vb6_VariantToString 交的是**新分配**的一枚
// (vb6_Format 的每个 return 都是 vb6_BSTR_FromStr)，搬进目标后所有权归目标，槽随后 clear
// 不会双 Free。数组载体 / 对象 / 64 位那几档交回空串 = 不回写 —— 装箱那一发照做，
// 「往标量存储里写整枚槽」的崩已经修掉，写回欠的是另一座桥（跟 ByRef As Object 同一族）。
std::string evtVariantWriteBack(const Vb6Type ty, const std::string& dst, const std::string& box,
                                const std::string& dstCType) {
    const uint16_t raw = static_cast<uint16_t>(ty);
    if ((raw & static_cast<uint16_t>(Vb6Type::Array)) != 0) return std::string();
    switch (static_cast<Vb6Type>(raw & ~static_cast<uint16_t>(Vb6Type::ByRef))) {
        case Vb6Type::String:
            return "vb6_BSTR_AssignMove(&(" + dst + "), vb6_VariantToString(" + box + "));";
        case Vb6Type::Boolean:
            return "(" + dst + ") = vb6_VariantToBool(" + box + ");";
        case Vb6Type::LongPtr:
            return "(" + dst + ") = vb6_VariantToLongPtr(" + box + ");";
        case Vb6Type::Single:
        case Vb6Type::Double:
        case Vb6Type::Date:
        case Vb6Type::Currency:
            return "(" + dst + ") = (" + dstCType + ")vb6_VariantToDouble(" + box + ");";
        case Vb6Type::Integer:
        case Vb6Type::Long:
        case Vb6Type::ULong:
        case Vb6Type::Byte:
        case Vb6Type::SByte:
        case Vb6Type::UInteger:
        case Vb6Type::Error:
            return "(" + dst + ") = (" + dstCType + ")vb6_VariantToLong(" + box + ");";
        default:
            return std::string();
    }
}
}  // namespace

// --- cgen_com_events.cpp: RaiseEvent 分发 + 事件接收器表 + 外部 COM vtable sink ---
// 由 src/backend/module/cgen_com.cpp 拆出（2026-09-17），纯搬移、零行为改动。
// 本文件内的 static comEventParamExpr 只被 emitComVtableSinks 使用，同文件内可见。


// ============================================================
// P6.5: RaiseEvent语句 → 生成事件回调分发代码
// ============================================================

void CCodeGen::visit(RaiseEventStmt& node) {
    // RaiseEvent EventName(args...)
    // 1. EXE内部回调: if (me->events && me->events->onEventName) { ... }
    // 2. DLL模式: vb6_FireEvent(comObj, dispid, args, argc) 广播给COM连接的接收器
    std::string evtId = cIdent(node.eventName);
    std::string callbackName = "on" + evtId;  // 事件接收器中的回调字段名

    // Fix 029: 查找当前模块的 EventDecl 以获取事件参数类型.
    // 必须 按 Variant 参数声明对实参做包装: 否则当事件参数声明 "As Variant"
    // (或无类型默认 Variant) 时, 直接回调 typedef 期望 vb6_VARIANT, 而实参
    // 是 int32_t/BSTR 等标量, 触发 C2440 "scalar → vb6_VARIANT" 错误.
    const EventDecl* evtDecl = nullptr;
    if (currentModule_) {
        for (auto& decl : currentModule_->declarations) {
            if (!decl) continue;
            if (decl->kind != ASTNodeKind::EventDecl) continue;
            auto& evt = static_cast<EventDecl&>(*decl);
            if (evt.name == node.eventName) { evtDecl = &evt; break; }
        }
    }

    c_.emitLine("if (me->events && me->events->" + callbackName + ") {");
    c_.indent();
    // 生成回调调用
    std::string call = "me->events->" + callbackName + "(me->events->handler";
    std::vector<std::string> evtRefTemps256;   // 不可取址的按地址实参：先落一枚临时量
    std::vector<std::string> evtPostCall308;   // 账 #305 第三百八刀：装箱槽的出站回写 + 清槽
    for (size_t i = 0; i < node.args.size(); i++) {
        emitExpr(*node.args[i]);
        std::string argVal = lastExpr_;
        // 账 #305 第二百五十六刀: 这一位形参按不按地址，问 src/common/event_param_abi.hpp
        // (与 emitEventSink 的 typedef / prelude 的 thunk / 处理器形参同源)。按地址那一档
        // 不走下面任何一条装箱 —— 那些出口发的是**右值**，取不了地址。
        const ParameterDecl* prm256 =
            (evtDecl && i < evtDecl->params.size()) ? evtDecl->params[i].get() : nullptr;
        const bool byAddr256 = prm256 && eventParamByAddress(
            prm256->isByRef,
            prm256->asType && prm256->asType->kind == ASTNodeKind::ArrayTypeRef);
        if (byAddr256) {
            /* 账 #305 第三百八刀: 形参声明的是**槽** (As Variant / 无型) 而实参在 C 侧不是槽 ——
               直递 `&(<标量>)` 把「标量的地址」交给了一枚 `vb6_VARIANT*` 形参，处理器一发赋值
               就往 4/8 字节的标量存储里写整枚槽 (x86 24B) ⇒ 栈被写穿。实测 ComCtlsDemo
               VirtualControls Demo 起窗后 0xC000041D，现场 ListView.c:13721
               `onGetVirtualItem(…, &(CallbackText))`（CallbackText = `Dim Text As String`），
               紧接着 13724 读它 ⇒ AV read 0x8 in wcscmp。VB6 在这一格交的就是一枚临时 Variant
               （装箱进去 + 调用后按声明档回写），照做。
               「是不是槽」两条判据**都要点头才留在直递**（inferExprType ∧ cExprIsVariant）：
               少装箱只是少一层保护，错装箱会把本来就是槽的那份复制再回写，改坏已正确的组合。 */
            const std::string prmBase256 =
                prm256->asType ? mapTypeRef(prm256->asType.get()) : std::string("vb6_VARIANT");
            const Vb6Type argTy308 = inferExprType(*node.args[i]);
            if (prmBase256 == "vb6_VARIANT" && argTy308 != Vb6Type::Variant &&
                !cExprIsVariant(argVal)) {
                const std::string src308 = argVal;
                const std::string box308 = "_evt_box_" + std::to_string(i);
                evtRefTemps256.push_back("vb6_VARIANT " + box308 + " = vb6_VariantEmpty();");
                evtRefTemps256.push_back("vb6_VariantAssign(&" + box308 + ", vb6_VariantFromValue(" +
                                         src308 + "));");
                if (evtArgIsLvalue(src308)) {
                    const std::string back308 =
                        evtVariantWriteBack(argTy308, src308, box308, mapType(argTy308));
                    if (!back308.empty()) evtPostCall308.push_back(back308);
                }
                evtPostCall308.push_back("vb6_VariantClear(&" + box308 + ");");
                call += ", &" + box308;
                continue;
            }
            if (evtArgIsLvalue(argVal)) {
                argVal = "&(" + argVal + ")";
            } else {
                // 实参不是左值（函数调用 / 带括号的下标 / 字面量）：VB6 那里递进去的本来
                // 就是一枚匿名临时量，处理器写回的值递出方同样看不见 ⇒ 落一枚同形临时量。
                const std::string baseTy256 =
                    prm256->asType ? mapTypeRef(prm256->asType.get()) : std::string("vb6_VARIANT");
                const std::string tmp256 = "_evt_ref_" + std::to_string(i);
                evtRefTemps256.push_back(baseTy256 + " " + tmp256 + " = " + argVal + ";");
                argVal = "&" + tmp256;
            }
            call += ", " + argVal;
            continue;
        }
        /* Fix 296 / 4b: 事件实参位上的 Extender 成员读也走实例入口 (CoolBar.ctl
           `RaiseEvent HeightChanged(Extender.Height)` 实测 ComCtlsDemo 1 处)。
           **这是第四条消费路径**: RaiseEvent 自带一条发码, 不经过 cgen_expr_call_arg_emit,
           所以那里那道闸够不着 —— 4b 前三道都装上了唯独这处漏掉, 就是哨兵少一条负控的代价。
           换道放在**捕获的那一刻**, 下面的 VariantFromValue / ObjectVal 装箱才套得住。 */
        argVal = extenderMemberReadExpr(argVal);
        // Fix 029: Variant 形参 → 用 vb6_VariantFromValue 包装标量实参.
        // 其它具体类型 (Long/String/...)形参期望已由 emitExpr 生成匹配类型,
        // 不做额外包装以避免破坏现有正确调用.
        // Fix 084k: 类名事件参数 — 回调typedef现用 mapTypeRef → vb6_cls_X*,
        // 直接传对象指针即可; 旧代码 resolveTypeName 失败回退为 Variant 导致
        // 这里错误包装成 vb6_VariantFromValue → 与 vb6_cls_X* 形参 C2440.
        if (evtDecl && i < evtDecl->params.size()) {
            auto& param = evtDecl->params[i];
            std::string pCType = (param->asType ? mapTypeRef(param->asType.get()) : std::string("vb6_VARIANT"));
            if (pCType == "vb6_VARIANT") {
                argVal = "vb6_VariantFromValue(" + argVal + ")";
            }
            // Fix 228: 事件形参声明为**项目类** (Fix 084k 的口径: mapTypeRef → vb6_cls_X*,
            // 回调 typedef 用真实类指针) 而实参在 C 层是 **VARIANT 值** (后期绑定属性读 /
            // 集合 .Item(i) 的结果, 形如 vb6_VariantFromComResult(vb6_ComCall(…, L"Item", …)))
            // 时, VB6 的语义是把 VARIANT 里装的那个对象取出来传给形参。
            // 修复前这里**什么都不做** ⇒ VARIANT 结构体直接塞给 vb6_cls_X* 形参 ⇒
            // C2440 "无法从 vb6_VARIANT 转换为 vb6_cls_LlbLink *"
            // (ComCtlsDemo LinkLabel.c:2997/3076/3091 `RaiseEvent LinkGetTipText(Links.Item(…))`
            //  及其余同形处, 共约 47 条)。
            // 判据刻意收窄到"必然编译报错"的那一格 (形参类指针 ∧ 实参 VARIANT 值),
            // 故不会改变任何当前能编过的组合。
            else if (pCType.rfind("vb6_cls_", 0) == 0 && cExprIsVariant(argVal)) {
                argVal = "(" + pCType + ")vb6_VariantToObjectVal(" + argVal + ")";
            }
            // Fix 267: 事件实参槽是**具体标量**而实参是**无型**的晚绑定 COM 调用
            // (`vb6_ComCall` 返回 void*) 时, 按目标槽选同型出口 —— 否则
            // C2440 "void* → float"。
            // 来时路: `With Me` 里的 `.ScaleY(...)` 在 UserControl 上落到通用 COM 方法
            // 调用 (comTypedCallExpr 的兜底: 签名表里没有 ScaleX/ScaleY ⇒ 发 vb6_ComCall),
            // 而同文件里 `Me.ScaleY(...)` 走 host-pseudo 表发出的是 vb6_UserControl_ScaleY。
            // 实测 ComCtlsDemo Pager.ctl `RaiseEvent Scroll(..., vb6_ComCall(...), ...)`
            // (Pager.c 2202/2208, typedef 第 4/6 位形参是 float)。
            else if (argVal.compare(0, 12, "vb6_ComCall(") == 0) {
                if (pCType == "float" || pCType == "double")
                    argVal = "vb6_ComCallDouble(" + argVal.substr(12);
                else if (pCType == "int32_t" || pCType == "int16_t" || pCType == "intptr_t")
                    argVal = "vb6_ComCallInt(" + argVal.substr(12);
                else if (pCType == "BSTR")
                    argVal = "vb6_ComCallBSTR(" + argVal.substr(12);
            }
        }
        call += ", " + argVal;
    }
    call += ");";
    for (const auto& d256 : evtRefTemps256) c_.emitLine(d256);
    c_.emitLine(call);
    for (const auto& p308 : evtPostCall308) c_.emitLine(p308);
    c_.dedent();
    c_.emitLine("}");

    // P6.6.3: DLL模式 - 广播事件给COM连接的接收器
    if (isDll_) {
        // 使用事件在eventNames中的索引+1作为DISPID (与driver.cpp TypeLib构建一致)
        // 因为driver中方法DISPID先分配(方法数), 然后事件DISPID从方法数+1开始递增
        // 但为简化, 我们在driver和cgen中都使用: 方法总数 + 事件索引 + 1
        // 这里使用 eventNames 索引位置即可, 因为 DISPID 只需在同一个 source dispinterface 内唯一
        int32_t evtDispid = 0;
        {
            auto* modScope = symTab_.moduleScope();
            if (modScope) {
                for (auto& [key, symPtr] : modScope->symbols()) {
                    if (!symPtr || symPtr->kind != SymbolKind::Class) continue;
                    if (symPtr->instancing == VBInstancing::Private) continue;
                    // 先尝试comEventDispids (从driver回写或从TypeLib导入)
                    std::string evtLower = node.eventName;
                    std::transform(evtLower.begin(), evtLower.end(), evtLower.begin(), ::tolower);
                    auto itDispId = symPtr->comEventDispids.find(evtLower);
                    if (itDispId != symPtr->comEventDispids.end()) {
                        evtDispid = itDispId->second;
                        break;
                    }
                    // 回退: 使用eventNames中的索引+1
                    for (size_t ei = 0; ei < symPtr->eventNames.size(); ei++) {
                        if (Symbol::toLower(symPtr->eventNames[ei]) == evtLower) {
                            evtDispid = (int32_t)(ei + 1);
                            break;
                        }
                    }
                    if (evtDispid != 0) break;
                }
            }
        }

        // 收集参数到VARIANT数组
        if (!node.args.empty()) {
            c_.emitLine("{ // vb6_FireEvent args scope");
            c_.indent();
            c_.emitLine("VARIANT __evt_args[" + std::to_string(node.args.size()) + "];");
            for (size_t i = 0; i < node.args.size(); i++) {
                c_.emitLine("VariantInit(&__evt_args[" + std::to_string(i) + "]);");
                emitExpr(*node.args[i]);
                // P15.1: 根据参数类型设置VARIANT
                Vb6Type argType = inferExprType(*node.args[i]);
                std::string idx = std::to_string(i);
                if (argType == Vb6Type::Long || argType == Vb6Type::Integer || argType == Vb6Type::Boolean) {
                    c_.emitLine("__evt_args[" + idx + "].vt = VT_I4; __evt_args[" + idx + "].lVal = (int32_t)(" + lastExpr_ + ");");
                } else if (argType == Vb6Type::Single) {
                    c_.emitLine("__evt_args[" + idx + "].vt = VT_R4; __evt_args[" + idx + "].lVal = *(int32_t*)&(float){" + lastExpr_ + "};");
                } else if (argType == Vb6Type::Double
                           || argType == Vb6Type::Currency
                           || argType == Vb6Type::Date) {
                    // Fix 123: Currency/Date 实参此前落 else → VBSTR 打包
                    // vb6_BSTR_FromStr(double) C2440 (cZipArchive RaiseEvent
                    // Progress(lIdx, uLocal.Ext.USize, .Size, ...) 的 18 处).
                    // 事件形参 As Currency 的 sink typedef 是 double, 统一 VT_R8.
                    c_.emitLine("__evt_args[" + idx + "].vt = VT_R8; __evt_args[" + idx + "].dblVal = (double)(" + lastExpr_ + ");");
                } else if (argType == Vb6Type::Object) {
                    c_.emitLine("__evt_args[" + idx + "].vt = VT_DISPATCH; __evt_args[" + idx + "].pdispVal = (IDispatch*)(" + lastExpr_ + ");");
                } else {
                    // Fix 084o-5: 仅当实参确实是 vb6_VARIANT (字符串级检测或已知
                    // Variant 变量) 才先提取 BSTR 值, 否则 vb6_BSTR_FromStr(vb6_VARIANT)
                    // 触发 C2440 (如 cZipArchive 的 RaiseEvent BeforeExtract(..., vFileName...)
                    // 中 vFileName As Variant). 注意: 对象指针 (vb6_cls_*/IDispatch* 等)
                    // 不能包 vb6_VariantToString — 那会把 C4047 警告升级为 C2440 错误,
                    // 因此只用 cExprIsVariant + knownVariantVars_ 判定, 不用 inferExprType
                    // (其 Variant 判定对 UDT 字段/对象变量不可靠).
                    bool evtArgIsVariant = cExprIsVariant(lastExpr_);
                    if (!evtArgIsVariant && i < node.args.size()
                        && node.args[i]->kind == ASTNodeKind::IdentifierExpr) {
                        auto& idArg = static_cast<IdentifierExpr&>(*node.args[i]);
                        std::string argLower = idArg.name;
                        std::transform(argLower.begin(), argLower.end(), argLower.begin(), ::tolower);
                        if (knownVariantVars_.count(argLower)) evtArgIsVariant = true;
                    }
                    if (evtArgIsVariant) {
                        c_.emitLine("__evt_args[" + idx + "].vt = VT_BSTR; __evt_args[" + idx + "].bstrVal = vb6_BSTR_FromStr(vb6_VariantToString(" + lastExpr_ + "));");
                    } else {
                        c_.emitLine("__evt_args[" + idx + "].vt = VT_BSTR; __evt_args[" + idx + "].bstrVal = vb6_BSTR_FromStr(" + lastExpr_ + ");");
                    }
                }
            }
            c_.emitLine("vb6_FireEvent((vb6_ComObject*)me->__comObj, " + std::to_string(evtDispid) + ", __evt_args, " + std::to_string(node.args.size()) + ");");
            for (size_t i = 0; i < node.args.size(); i++) {
                c_.emitLine("VariantClear(&__evt_args[" + std::to_string(i) + "]);");
            }
            c_.dedent();
            c_.emitLine("}");
        } else {
            c_.emitLine("vb6_FireEvent((vb6_ComObject*)me->__comObj, " + std::to_string(evtDispid) + ", NULL, 0);");
        }
    }
}

void CCodeGen::visit(BeepStmt& node) {
    c_.emitLine("vb6_Beep();");
}

void CCodeGen::visit(DoEventsStmt& node) {
    c_.emitLine("vb6_DoEvents();");
}

// ============================================================
// P6.5: 事件接收器表生成 (Event Sink Table)
// ============================================================

void CCodeGen::emitEventSink(Module& module) {
    std::string clsStruct = "vb6_cls_" + cIdent(moduleName_);  // Fix 013: VB_Name
    std::string sinkName = "vb6_events_" + cIdent(moduleName_);  // Fix 013: VB_Name

    // 收集所有Event声明
    struct EventInfo {
        std::string name;           // 事件名(原始)
        std::string cName;          // 安全C标识符
        std::vector<ParameterInfo> params;  // 事件参数
        std::vector<std::string> ctypes;    // Fix 084k: 每个参数的C类型+名字
    };
    std::vector<EventInfo> events;

    for (auto& decl : module.declarations) {
        if (decl->kind == ASTNodeKind::EventDecl) {
            auto& evt = static_cast<EventDecl&>(*decl);
            EventInfo info;
            info.name = evt.name;
            info.cName = cIdent(evt.name);
            for (auto& param : evt.params) {
                ParameterInfo pi;
                pi.name = param->name;
                // 用mapTypeRef获取C类型, 同时从TypeSystem获取Vb6Type
                std::string cType = mapTypeRef(param->asType.get());
                pi.type = (param->asType && param->asType->kind == ASTNodeKind::SimpleTypeRef)
                    ? typeSys_.resolveTypeName(static_cast<SimpleTypeRef*>(param->asType.get())->name)
                    : Vb6Type::Variant;
                if (pi.type == Vb6Type::Unknown || pi.type == Vb6Type::Empty)
                    pi.type = Vb6Type::Variant;
                // 账 #305 第二百五十六刀: 「按不按地址」问的是**事件声明里那一位形参**，
                // 主人 = src/common/event_param_abi.hpp。从前这里无条件 isByVal=true
                // (注释「事件参数总是ByVal传递(跨对象边界)」)，把声明里明写的 ByRef 出参
                // 一起按值发了；今天只有**没写 ByRef** 的那些维持按值 (语料里那几千处形状不动)。
                const bool byAddr256 = eventParamByAddress(
                    param->isByRef,
                    param->asType && param->asType->kind == ASTNodeKind::ArrayTypeRef);
                pi.isByVal = !byAddr256;
                info.params.push_back(pi);
                // Fix 084k: 回调typedef须与包装函数签名一致 — 类名参数用 mapTypeRef
                // → vb6_cls_X* (ByVal单指针), 不能回退为 vb6_VARIANT (导致 C2440)
                info.ctypes.push_back(eventParamCDecl(cType, byAddr256, cIdent(param->name)));
            }
            events.push_back(std::move(info));
        }
    }

    if (events.empty()) return;

    // 1. 生成事件回调函数指针typedef
    // Fix 010: typedef名称包含类名前缀, 避免不同类同名事件(但不同签名)的typedef冲突 (C2370/C2040/C2371)
    h_.emitLine("// P6.5: Event callback function pointer types");
    for (auto& evt : events) {
        std::string cbName = "vb6_evt_" + cIdent(moduleName_) + "_" + evt.cName + "_cb";
        std::string sig = "void (*" + cbName + ")(void* handler";
        for (auto& pt : evt.ctypes) {
            sig += ", " + pt;
        }
        sig += ")";
        h_.emitLine("typedef " + sig + ";");
    }
    h_.emitBlank();

    // 2. 生成事件接收器表结构体
    h_.emitLine("// Event sink table: " + module.moduleName);
    h_.emitLine("typedef struct " + sinkName + " {");
    h_.emitLine("    void* handler;  /* event handler object (consumer) */");
    for (auto& evt : events) {
        std::string cbName = "vb6_evt_" + cIdent(moduleName_) + "_" + evt.cName + "_cb";
        h_.emitLine("    " + cbName + " on" + evt.cName + ";  /* Event " + evt.name + " */");
    }
    h_.emitLine("} " + sinkName + ";");
    h_.emitBlank();
}

void CCodeGen::visit(ParameterDecl& node) {
    // 由makeParamList内部处理
}

// ============================================================
// P6.6: ActiveX DLL代码生成
// 为ActiveX DLL工程生成COM服务端代码:
//   1. coclass描述表 (g_vb6_coclasses[])
//   2. IDispatch方法描述 (g_vb6_disp_<Class>Methods[])
//   3. IDispatch方法调用桥接 (vb6_disp_<Class>_<Method>_invoke)
//   4. DllGetClassObject / DllCanUnloadNow
//   5. DllRegisterServer / DllUnregisterServer
//   6. .def导出文件
// ============================================================
// P13.23: 外部COM vtable source interface 事件接收器生成
// ============================================================

void CCodeGen::emitComVtableSinkDecls() {
    for (auto& [varLower, srcClassName] : knownWithEventsVars_) {
        auto* srcClsSym = symTab_.lookup(srcClassName);
        if (!srcClsSym || srcClsSym->kind != SymbolKind::ComClass) continue;
        if (!srcClsSym->comHasSourceIface) continue;
        if (srcClsSym->comSourceIfaceIsDispOnly) continue;  // dispinterface uses vb6_CreateEventSink
        if (srcClsSym->comSourceMethods.empty()) continue;
        if (srcClsSym->comSourceIfaceIid.empty()) continue;
        std::string createFn = "vb6_vsink_" + varLower + "_create";
        // Fix 095: 声明须与定义/调用一致 (void* handler)。原声明 (void) 导致
        // 调用点 (传 handler) 报 C2197 参数太多。
        h_.emitLine("void* " + createFn + "(void* handler);");
        h_.emitBlank();
    }
}

static std::string comEventParamExpr(const std::string& comName, Vb6Type uhType) {
    // 将 COM vtable 参数转换为 VB6 用户处理器参数表达式
    bool isArray = (static_cast<uint16_t>(uhType) & static_cast<uint16_t>(Vb6Type::Array)) != 0;
    Vb6Type baseType = isArray
        ? static_cast<Vb6Type>(static_cast<uint16_t>(uhType) & ~static_cast<uint16_t>(Vb6Type::Array))
        : uhType;
    if (isArray) {
        if (baseType == Vb6Type::Byte) {
            return "(" + comName + " ? (uint8_t*)" + comName + "->pvData : NULL)";
        } else {
            return "(" + comName + " ? (void*)" + comName + "->pvData : NULL)";
        }
    }
    return comName;
}

void CCodeGen::emitComVtableSinks() {
    for (auto& [varLower, srcClassName] : knownWithEventsVars_) {
        auto* srcClsSym = symTab_.lookup(srcClassName);
        if (!srcClsSym || srcClsSym->kind != SymbolKind::ComClass) continue;
        if (!srcClsSym->comHasSourceIface) continue;
        if (srcClsSym->comSourceIfaceIsDispOnly) continue;  // dispinterface uses vb6_CreateEventSink
        if (srcClsSym->comSourceMethods.empty()) continue;
        if (srcClsSym->comSourceIfaceIid.empty()) continue;

        std::string iidStr = srcClsSym->comSourceIfaceIid;
        std::string sinkType = "vb6_vsink_" + varLower;
        std::string sinkPrefix = "vb6_vsink_" + varLower;
        std::string iidConst = sinkPrefix + "_iid";
        std::string guidInit = emitGuidInitializer(iidStr);
        if (guidInit.empty()) continue;

        // 查找原始变量名(保留大小写)
        std::string varName = varLower;
        auto* varSym = symTab_.lookup(varLower);
        if (varSym) varName = varSym->name;

        // 1. GUID 常量
        c_.emitBlank();
        c_.emitLine("// P13.23: vtable event sink for " + varName + " (" + srcClassName + ")");
        c_.emitLine("static const IID " + iidConst + " = " + guidInit + ";");

        // 2. Sink 结构体
        c_.emitLine("typedef struct " + sinkType + " {");
        c_.emitLine("    void** vtable;");
        c_.emitLine("    LONG refCount;");
        c_.emitLine("    void* handler;");  // 宿主实例(类模块 WithEvents 时为 me)
        c_.emitLine("} " + sinkType + ";");

        // Forward declarations (used by QI before definition)
        c_.emitBlank();
        c_.emitLine("static ULONG __stdcall " + sinkPrefix + "_AddRef(void* This);");
        c_.emitLine("static ULONG __stdcall " + sinkPrefix + "_Release(void* This);");

        // 3. IUnknown methods
        // Note: QI does NOT respond to IID_IDispatch because this sink has no
        // Invoke implementation — only IUnknown and the source IID are supported.
        c_.emitBlank();
        c_.emitLine("static HRESULT __stdcall " + sinkPrefix + "_QI(void* This, REFIID riid, void** ppv) {");
        c_.indent();
        c_.emitLine("if (IsEqualIID(riid, \u0026IID_IUnknown) || IsEqualIID(riid, \u0026" + iidConst + ")) {");
        c_.indent();
        c_.emitLine("*ppv = This;");
        c_.emitLine(sinkPrefix + "_AddRef(This);");
        c_.emitLine("return S_OK;");
        c_.dedent();
        c_.emitLine("}");
        c_.emitLine("*ppv = NULL;");
        c_.emitLine("return E_NOINTERFACE;");
        c_.dedent();
        c_.emitLine("}");

        c_.emitBlank();
        c_.emitLine("static ULONG __stdcall " + sinkPrefix + "_AddRef(void* This) {");
        c_.indent();
        c_.emitLine("return InterlockedIncrement(\u0026((" + sinkType + "*)This)->refCount);");
        c_.dedent();
        c_.emitLine("}");

        c_.emitBlank();
        c_.emitLine("static ULONG __stdcall " + sinkPrefix + "_Release(void* This) {");
        c_.indent();
        c_.emitLine("ULONG c = InterlockedDecrement(\u0026((" + sinkType + "*)This)->refCount);");
        c_.emitLine("if (c == 0) CoTaskMemFree(This);");
        c_.emitLine("return c;");
        c_.dedent();
        c_.emitLine("}");

        // 4. Source interface 方法
        // Fix 096: 未实现的事件也要生成空 stub — vtable 数组引用全部事件名,
        // 而 VB6 语义允许 WithEvents 只实现感兴趣的事件 (其余静默).
        // 此前 if(!handlerSym) continue 跳过函数体生成, vtable 引用悬空
        // → C2065 (cHttpClient.c 684/685 OnResponseStart/OnResponseDataAvailable)
        // 及关联 C2099 (vtable 初始化器不是常量).
        for (auto& evtName : srcClsSym->eventNames) {
            std::string evtLower = Symbol::toLower(evtName);
            auto itSig = srcClsSym->comSourceMethods.find(evtLower);
            if (itSig == srcClsSym->comSourceMethods.end()) continue;
            auto& sig = itSig->second;

            std::string handlerName = varName + "_" + evtName;
            auto* handlerSym = symTab_.lookup(handlerName);

            std::string methodName = sinkPrefix + "_" + cIdent(evtName);
            std::string methodSig = "static HRESULT __stdcall " + methodName + "(void* This";
            std::string callArgs = "(";
            for (size_t i = 0; i < sig.params.size(); i++) {
                std::string comType = mapComType(sig.params[i].type);
                std::string comName = "com_" + cIdent(sig.params[i].name);
                methodSig += ", " + comType + " " + comName;
                if (i > 0) callArgs += ", ";
                callArgs += comEventParamExpr(comName, sig.params[i].type);
            }
            methodSig += ")";
            callArgs += ")";

            c_.emitBlank();
            c_.emitLine(methodSig + " {");
            c_.indent();
            if (!handlerSym) {
                // Fix 096: 未实现的事件 → 空 stub, 仅保持 sink 接口完整
                c_.emitLine("(void)This;  /* event not implemented by user (VB6 semantics) */");
                c_.emitLine("return S_OK;");
                c_.dedent();
                c_.emitLine("}");
                continue;
            }
            std::string procCall = cProcName(handlerName, handlerSym->access,
                                             handlerSym->isExternal ? handlerSym->sourceModule : (isClassModule_ ? moduleName_ : ""));
            if (isClassModule_) {
                // 类模块: 处理器是类方法(带 me 形参), 需补宿主实例 (sink->handler = me)
                std::string argsInner = callArgs;
                if (argsInner.size() >= 2 && argsInner.front() == '(' && argsInner.back() == ')') {
                    argsInner = argsInner.substr(1, argsInner.size() - 2);
                }
                std::string callFull = procCall + "(((" + sinkType + "*)This)->handler";
                if (!argsInner.empty()) callFull += ", " + argsInner;
                c_.emitLine(callFull + ");");
            } else {
                c_.emitLine(procCall + callArgs + ";");
            }
            c_.emitLine("return S_OK;");
            c_.dedent();
            c_.emitLine("}");
        }

        // 5. vtable 数组
        c_.emitBlank();
        c_.emitLine("static void* " + sinkPrefix + "_vtable[] = {");
        c_.indent();
        c_.emitLine(sinkPrefix + "_QI, " + sinkPrefix + "_AddRef, " + sinkPrefix + "_Release,");
        for (auto& evtName : srcClsSym->eventNames) {
            std::string evtLower = Symbol::toLower(evtName);
            if (srcClsSym->comSourceMethods.find(evtLower) == srcClsSym->comSourceMethods.end()) continue;
            c_.emitLine(sinkPrefix + "_" + cIdent(evtName) + ",");
        }
        c_.dedent();
        c_.emitLine("};");

        // 6. create 函数
        c_.emitBlank();
        c_.emitLine("void* " + sinkPrefix + "_create(void* handler) {");
        c_.indent();
        c_.emitLine(sinkType + "* s = (" + sinkType + "*)CoTaskMemAlloc(sizeof(" + sinkType + "));");
        c_.emitLine("if (!s) return NULL;");
        c_.emitLine("memset(s, 0, sizeof(" + sinkType + "));");
        c_.emitLine("s->vtable = " + sinkPrefix + "_vtable;");
        c_.emitLine("s->refCount = 1;");
        c_.emitLine("s->handler = handler;");
        c_.emitLine("return s;");
        c_.dedent();
        c_.emitLine("}");
    }
}

} // namespace vb6c3
