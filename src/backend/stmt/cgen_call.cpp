#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_call.cpp: 调用与数组语句生成 (Call / ReDim / Erase) ---

void CCodeGen::visit(CallStmt& node) {
    if (node.callee) {
        // Fix <vbeclipse>: 无括号语句式 `Parent.<方法>` (ucSplitBar.ctl:138
        // `Parent.Refresh`) 的 callee 是**裸 MemberAccessExpr** (无实参的 paren-less
        // 调用不进 IndexOrCallExpr, 见下方 Fix 090g 注释), 走到 emitExpr 时
        // asCallCallee_=true 只能交出函数名 → 落到 class_module.inc 优先级3 的
        // "模块名.成员" 回退 → `vb6_Parent_Refresh();` C2065 (ucSplitBar.c:205)。
        // `Parent` 是 VB6 内建的 As Object 宿主对象, 成员调用本就是 IDispatch
        // 后期绑定, 这里直接发 vb6_ComCall(vb6_UC_ParentObject(), L"<成员>", NULL, 0)。
        // (带实参的形态走 cgen_expr_call_callee_member.inc 的 IndexOrCallExpr 分支。)
        if (node.callee->kind == ASTNodeKind::MemberAccessExpr && isDesignerModule_) {
            auto& maParCs = static_cast<MemberAccessExpr&>(*node.callee);
            if (maParCs.object && maParCs.object->kind == ASTNodeKind::IdentifierExpr
                && Symbol::toLower(
                       static_cast<IdentifierExpr&>(*maParCs.object).name) == "parent"
                && !knownLocalVars_.count("parent")) {
                c_.emitLine("vb6_ComCall(vb6_UC_ParentObject(), L\""
                            + escapeWideCString(maParCs.memberName)
                            + "\", NULL, 0);  /* Parent.<method> */");
                return;
            }
        }
        // 检测 Debug.Print 调用: 特殊处理多参数输出
        if (node.callee->kind == ASTNodeKind::IndexOrCallExpr) {
            auto& call = static_cast<IndexOrCallExpr&>(*node.callee);
            if (call.callee && call.callee->kind == ASTNodeKind::MemberAccessExpr) {
                auto& member = static_cast<MemberAccessExpr&>(*call.callee);
                if (member.object && member.object->kind == ASTNodeKind::IdentifierExpr) {
                    auto& objIdent = static_cast<IdentifierExpr&>(*member.object);
                    std::string objLower = objIdent.name;
                    std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
                    std::string memLower = member.memberName;
                    std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);

                    // Fix <vbeclipse>: 语句式 `Debug.Assert <expr>` —— VB6 **编译版**语义是
                    // 整句被移除且**条件不求值** (只有 IDE 里才求值). 这正是"用一个带 ByRef
                    // 出参的函数把函数自身返回值置 True"的经典写法所依赖的:
                    //   Private Function Subclass_InIDE() As Boolean
                    //       Debug.Assert zSetTrue(Subclass_InIDE)   ' zSetTrue 置 True, 返回 True
                    //   End Function
                    // 编译版必须返回 False (MagneticWnd 据此选择 SetWindowLongA 子类化路径);
                    // 若照常求值 → 返回 True → 走 IDE 分支用 vba6!EbMode 的地址(NULL, 因为
                    // 独立 EXE 没加载 vba6)去 patch 机器码桩 → 桩内 call 0 → 运行期
                    // 0xC0000005(实测读 0x784000) + 堆损坏 (VbEclipse play78.exe).
                    // 故此处只落一条注释, 实参一个都不 emit (求值即产生副作用).
                    if (objLower == "debug" && memLower == "assert") {
                        c_.emitLine("/* Debug.Assert <expr> removed (compiled-mode semantics) */");
                        return;
                    }

                    // Fix 161: Console.WriteLine/.Write 复用 Debug.Print 的"逐参转 BSTR
                    // 后输出"路径 (两者都是语句式输出调用, 参数需按 BSTR 转换; 区别仅在
                    // 换行/输出函数)。Console.Write 不换行, WriteLine 换行。
                    const bool isConsoleWrite = (objLower == "console" &&
                        (memLower == "writeline" || memLower == "write"));
                    const bool isConsoleWriteLine = (objLower == "console" && memLower == "writeline");

                    if ((objLower == "debug" && memLower == "print") || isConsoleWrite) {
                        // Debug.Print: 逐参数输出, 最后换行
                        // 每个参数转为BSTR后用vb6_DebugWriteBSTR输出
                        if (call.positional.empty()) {
                            c_.emitLine("vb6_DebugWriteNewline();");
                        } else {
                            // 已知返回BSTR的内置函数前缀
                            static const std::vector<std::string> bstrFuncs = {
                                "vb6_BSTR_FromStr", "vb6_Left", "vb6_Right", "vb6_Mid",
                                "vb6_UCase", "vb6_LCase", "vb6_UCase_str", "vb6_LCase_str",
                                "vb6_Trim", "vb6_LTrim", "vb6_RTrim", "vb6_Chr",
                                "vb6_Str", "vb6_CStr", "vb6_Format", "vb6_Hex", "vb6_Oct",
                                "vb6_Replace", "vb6_Space", "vb6_String", "vb6_StrReverse",
                                "vb6_BSTR_Concat", "vb6_BSTR_Empty", "vb6_App_Path", "vb6_App_EXEName", "vb6_App_HelpFile", "vb6_Command", "vb6_CurDir", "vb6_Environ", "vb6_Dir", "vb6_IIfBSTR", "vb6_GetControlText", "vb6_GetControlCaption",
                                "vb6_ComCallBSTR",  // Fix 160-com-byref: 早期绑定COM方法返回BSTR (StringOf/StringAt 等) — 按BSTR打印
                                "vb6_Console_ReadLine", "vb6_Console_ReadKey"  // Fix 161: Console 输入返回 BSTR
                            };
                            // <vbeclipse>: 名字前缀匹配必须紧跟 `(` 才算命中。
                            // 旧写法 expr.compare(0, size, prefix)==0 让短名字吞掉长名字:
                            //   "vb6_Str" 命中 "vb6_StrComp(" ⇒ Debug.Print StrComp(a,b)
                            //   被发成 vb6_DebugWriteBSTR(返回 1) → 把 1 当 BSTR 指针
                            //   → 段错误 (实测, 改动前后两版都崩; Len/InStr/Asc 因无短前缀
                            //   而正常, 所以整族只崩 StrComp 这一条)。
                            // 同一形状也保护 doubleFuncs ("vb6_Int" vs "vb6_Integer…")。
                            auto callNameHit = [](const std::string& expr,
                                                  const std::string& name) -> bool {
                                if (expr.compare(0, name.size(), name) != 0) return false;
                                return expr.size() > name.size() && expr[name.size()] == '(';
                            };
                            auto isBstrExpr = [&](const std::string& expr) -> bool {
                                for (auto& prefix : bstrFuncs) {
                                    if (callNameHit(expr, prefix)) return true;
                                }
                                // vb6_BSTR_ 开头的都是 BSTR
                                if (expr.compare(0, 8, "vb6_BSTR") == 0) return true;
                                // VB6_SA_AT(BSTR, ...) 也是 BSTR
                                if (expr.find("VB6_SA_AT(BSTR,") != std::string::npos) return true;
                                // 已知BSTR变量名 (小写匹配)
                                std::string lower = expr;
                                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                                if (knownBstrVars_.count(lower)) return true;
                                return false;
                            };

                            // 已知返回double的内置函数前缀
                            static const std::vector<std::string> doubleFuncs = {
                                "vb6_Sin", "vb6_Cos", "vb6_Tan", "vb6_Atn",
                                "vb6_Log", "vb6_Exp", "vb6_Sqr", "vb6_Rnd",
                                "vb6_Round", "vb6_Fix", "vb6_Int",
                                "vb6_CDbl", "vb6_CSng", "vb6_Val",
                                "vb6_Abs"
                            };
                            auto isDoubleExpr = [&](const std::string& expr) -> bool {
                                for (auto& prefix : doubleFuncs) {
                                    if (callNameHit(expr, prefix)) return true;
                                }
                                // 已知double变量名 (小写匹配)
                                std::string lower = expr;
                                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                                if (knownDoubleVars_.count(lower)) return true;
                                // 包含浮点字面量 (如 3.14)
                                // 检查是否包含小数点且不是函数调用
                                if (expr.find('.') != std::string::npos && expr.find('(') == std::string::npos) return true;
                                return false;
                            };

                            // Fix 091r: Debug.Print 参数是否为 Variant 值 —
                            // cExprIsVariant 只认 C 表达式前缀, 不认"Variant 变量名"
                            // (局部/模块级 knownVariantVars_ / 类字段 classVariantFields_).
                            // Demo 671: For Each 循环变量 x As Variant 被赋
                            // vb6_VariantFromStackVARIANT 后 Debug.Print x →
                            // DebugWriteLong((int32_t)(x)) C2440 (vb6_VARIANT→int32_t).
                            auto isVariantVal091r = [&](const std::string& e) -> bool {
                                if (cExprIsVariant(e)) return true;
                                std::string n091r = e;
                                std::string suffix091r;
                                size_t cmt091r = n091r.find("/*");
                                if (cmt091r != std::string::npos) {
                                    suffix091r = n091r.substr(cmt091r);
                                    n091r = n091r.substr(0, cmt091r);
                                }
                                while (!n091r.empty() && (n091r.back() == ' ' || n091r.back() == '\t')) {
                                    n091r.pop_back();
                                }
                                if (n091r.rfind("me->", 0) == 0) n091r = n091r.substr(4);
                                else if (n091r.size() > 4 && n091r[0] == '(' && n091r[1] == '*'
                                         && n091r.back() == ')') {
                                    n091r = n091r.substr(2, n091r.size() - 3);
                                }
                                std::string ln091r = Symbol::toLower(n091r);
                                if (knownVariantVars_.count(ln091r)
                                    || classVariantFields_.count(ln091r)) {
                                    return true;
                                }
                                return false;
                            };

                            for (size_t j = 0; j < call.positional.size(); j++) {
                                emitExpr(*call.positional[j]);
                                std::string val = std::move(lastExpr_);

                                // COM属性读取 (P6.2): isComMarker_标志
                                if (isComMarker_) {
                                    isComMarker_ = false;
                                    // 使用一体化函数, 内部处理VARIANT清理
                                    std::string comPropCall = "vb6_ComGetStringProp(" + comObjExpr_ + ", L\"" + comMemberName_ + "\")";
                                    c_.emitLine("{");
                                    c_.indent();
                                    c_.emitLine("wchar_t* _dbg_com_bstr = " + comPropCall + ";");
                                    c_.emitLine("vb6_DebugWriteBSTR(_dbg_com_bstr);");
                                    c_.emitLine("vb6_BSTR_Free(_dbg_com_bstr);");
                                    c_.dedent();
                                    c_.emitLine("}");
                                    comObjExpr_.clear();
                                    comMemberName_.clear();
                                    continue;
                                }

                                if (isBstrExpr(val)
                                    || inferExprType(*call.positional[j]) == Vb6Type::String) {
                                    // 已经是BSTR, 直接输出
                                    // Fix 176: 前缀表只认内置函数名, **用户 Function 返回
                                    // String** (Debug.Print "x="; MyFunc()) 捡不到 → 落到
                                    // 下面的 DebugWriteLong((int32_t)(BSTR)) → 指针截断成
                                    // int32, 打印出 -150012728 这类垃圾数 (不是崩溃, 静默错)。
                                    // 补一条按 AST 的返回类型判定 (inferExprType 会查
                                    // symTab_ 里 Function 符号的返回类型)。
                                    c_.emitLine("vb6_DebugWriteBSTR(" + val + ");");
                                } else if (inferExprType(*call.positional[j]) == Vb6Type::Date) {
                                    // Fix 175: Debug.Print d (d As Date) —— Date 变量同时
                                    // 登记在 knownDoubleVars_, isDoubleExpr 会抢先命中并
                                    // 打成序列号 46023; 按 VB6 应是短日期串。
                                    c_.emitLine("vb6_DebugWriteBSTR(vb6_CStrDate((double)(" + val + ")));");
                                } else if (isDoubleExpr(val)
                                    || inferExprType(*call.positional[j]) == Vb6Type::Double
                                    || inferExprType(*call.positional[j]) == Vb6Type::Single) {
                                    // 浮点数, 用DebugWriteDouble输出
                                    // <vbeclipse>: isDoubleExpr 是一张**函数名前缀清单**, 清单外的
                                    // 浮点表达式以前一律落到 DebugWriteLong((int32_t)(x)) ——
                                    // 实测 `Debug.Print CDbl(v)` (发的是 vb6_CDblV) 把 1.5 打成 1,
                                    // VB6 打 1.5。判定改按 AST 类型走 (同上面 BSTR/Date/Boolean
                                    // 三档已有的口径), 名字清单只当补充。
                                    c_.emitLine("vb6_DebugWriteDouble((double)(" + val + "));");
                                } else if (isVariantVal091r(val)) {
                                    // Fix 090x: Debug.Print x (x As Variant 变量 /
                                    // Variant 表达式) — 运行时值按字符串输出. 此前落入
                                    // DebugWriteLong((int32_t)(x)) → C2440 (无法从
                                    // vb6_VARIANT 转换 int32_t).
                                    // Fix 091r: 判定改用 isVariantVal091r (含
                                    // knownVariantVars_/类字段 兜底).
                                    c_.emitLine("vb6_DebugWriteBSTR(vb6_VariantToString(" + val + "));");
                                } else if (inferExprType(*call.positional[j]) == Vb6Type::Boolean) {
                                    // Fix 198: Debug.Print b (b As Boolean) 落到下面的
                                    // DebugWriteLong 会打成 -1/0 —— VB6 打 True/False,
                                    // 与 CStr / `&` 拼接同一口径 (同一个值两条路读数不同,
                                    // 正是 ai/022 待拍板 5 的那处不一致)。
                                    c_.emitLine("vb6_DebugWriteBSTR(vb6_CStrBool(" + val + "));");
                                } else {
                                    // 整数/布尔值, 用DebugWriteLong输出
                                    c_.emitLine("vb6_DebugWriteLong((int32_t)(" + val + "));");
                                }
                            }
                            // Fix 161: Console.Write 不换行; Debug.Print / Console.WriteLine 换行
                            if (!isConsoleWrite || isConsoleWriteLine) {
                                c_.emitLine("vb6_DebugWriteNewline();");
                            }
                        }
                        return;
                    }
            }
        }
    }

        // Fix 015: 标记 callee 上下文, 让 visit(MemberAccessExpr) 的 Fix 015 路径
        // 把链式对象参数通过 pendingChainObj_ 交付, 而不是直接合成 func(wrappedObj)
        // (那样会让下面的 bare-call 判定 callExpr.find('(') != npos 错过 padding).
        pendingChainObj_.clear();
        bool savedAsCallCallee = asCallCallee_;
        asCallCallee_ = true;
        emitExpr(*node.callee);
        asCallCallee_ = savedAsCallCallee;
        // 语句级调用: 确保表达式被求值(即使是void调用)
        // 如果结果是函数名(不含括号), 自动添加()调用
        std::string callExpr = lastExpr_;

        // Fix 150 (Fix 146 回归补全 / T2 ExtShow): 无括号无参语句 `Frm.Show` —
        // parser 对行尾无参数的成员调用直接交付 CallStmt(callee=MemberAccessExpr)
        // (parser_stmt_assign 收尾), 不经过 IndexOrCallExpr; 而 form_builtin 的
        // external form show 转发只交付 vb6_form_show_<Form>(NULL) 半成品
        // (modal 实参由 com_bind 拆包路径追加), 语句形态原样 emit → 1 参调用
        // 对 Fix 146 的 2 参签名 (void* hMDIClient, int modal) → C2198.
        // 无参 Show 即 modeless, 此处补默认 modal=0.
        if (callExpr.size() > 6
            && callExpr.compare(0, 14, "vb6_form_show_") == 0
            && callExpr.compare(callExpr.size() - 6, 6, "(NULL)") == 0) {
            callExpr = callExpr.substr(0, callExpr.size() - 1) + ", 0)";
        }

        // COM调用检测 (P6.2): isComMarker_标志
        if (isComMarker_) {
            isComMarker_ = false;
            // P20-39: ImageList 原生复刻 —— **无实参**的 `ImageList1.ListImages.Clear`
            // 走不到 visit(IndexOrCallExpr&), 是在这条语句路径上收尾的。这里必须同样拦一道,
            // 否则掉回 vb6_ComCall(vb6_ComGetObjectProp(...), L"Clear", NULL, 0) 的假 IDispatch。
            // (cgen_expr_call_com_bind.inc 那条管带实参的 Add/Remove, 两条并行, 别只挂一头。)
            if (imageListNameOfExpr(comObjExpr_) != ""
                && Symbol::toLower(comMemberName_) == "clear") {
                std::string ilSlot = "vb6_com_" + imageListNameOfExpr(comObjExpr_);
                comObjExpr_.clear();
                comMemberName_.clear();
                c_.emitLine("vb6_ImageList_ClearImages((void*)" + ilSlot + ");");
                return;
            }
            // C29-Data: `Data1.Recordset.Refresh` 等 —— comObjExpr_ 是 vb6_Data_Self( 透传形态
            // (recordset 走控件属性分支不设裸控件标记, 所以这里认前缀而不是查表)。
            // 不拦就落 vb6_ComCall(vb6_Data_Self(...), L"Refresh") —— 对 ODBC 状态当 IDispatch 用。
            if (comObjExpr_.find("vb6_Data_Self(") == 0) {
                std::string dhD = dataSelfHwndExpr(comObjExpr_);
                std::string mD = Symbol::toLower(comMemberName_);
                if (!dhD.empty()) {
                    std::string fnD;
                    if (mD == "refresh")      fnD = "vb6_Data_Refresh";
                    else if (mD == "movenext")     fnD = "vb6_Data_MoveNext";
                    else if (mD == "moveprevious") fnD = "vb6_Data_MovePrevious";
                    else if (mD == "movefirst")    fnD = "vb6_Data_MoveFirst";
                    else if (mD == "movelast")     fnD = "vb6_Data_MoveLast";
                    if (!fnD.empty()) {
                        comObjExpr_.clear();
                        comMemberName_.clear();
                        c_.emitLine(fnD + "((void*)" + dhD + ");");
                        return;
                    }
                }
            }
            // P20-40: 同款 —— `StatusBar1.Panels.Clear` 无实参, 也必须在这条语句路径收尾。
            if (statusBarNameOfExpr(comObjExpr_) != ""
                && Symbol::toLower(comMemberName_) == "clear") {
                std::string sbHwnd = "vb6_hwnd_" + statusBarNameOfExpr(comObjExpr_);
                comObjExpr_.clear();
                comMemberName_.clear();
                c_.emitLine("vb6_StatusBar_ClearPanels((void*)" + sbHwnd + ");");
                return;
            }
            // D6 / C29-9: 无括号的 `CommonDialog1.ShowOpen` —— 与 List1.Clear 同一条
            // 语句路。不接这里的话会落到下面 `vb6_ComCall(dl1, L"ShowOpen")`：既编不过
            // (裸控制名)，也正是本批要拆掉的 OCX 形状。
            {
                auto itCd = knownFormControls_.find(comObjExpr_);
                std::string mCd = Symbol::toLower(comMemberName_);
                if (itCd != knownFormControls_.end()
                    && itCd->second == FrmControlType::CommonDialog
                    && (mCd == "showopen" || mCd == "showsave" || mCd == "showcolor"
                        || mCd == "showfont" || mCd == "showprinter" || mCd == "showabout")) {
                    std::string hwndCd = cIdent(knownFormControlOriginalNames_.count(comObjExpr_)
                        ? knownFormControlOriginalNames_[comObjExpr_] : comObjExpr_);
                    std::string fnCd = "vb6_CdShow"
                        + std::string(1, (char)::toupper((unsigned char)mCd[4])) + mCd.substr(5);
                    comObjExpr_.clear();
                    comMemberName_.clear();
                    c_.emitLine(fnCd + "((void*)vb6_hwnd_" + hwndCd + ");"
                                "  /* CommonDialog." + mCd + " (原生 comdlg32) */");
                    return;
                }
            }
            // C29-WS-a/b: 无实参的 Winsock 方法（`Winsock1.Close` / `.Listen` / `Winsock1.Connect`）
            // 也走这条语句路 —— 与上面 CommonDialog.Show* 同族。带实参那四条（Bind / SendData /
            // GetData / PeekData / Accept）在 cgen_expr_call_callee_withm.inc 收尾，两边都缺一头就是
            // `vb6_ComCall(wsa, L"close", NULL, 0)`：裸控制名 ⇒ C2065，或者更糟 —— 编得过、
            // 运行期一声不响（本线踩过三次的那同一条）。
            {
                auto itWs = knownFormControls_.find(comObjExpr_);
                std::string mWs = Symbol::toLower(comMemberName_);
                const char* fnWs = nullptr;
                if (itWs != knownFormControls_.end() && itWs->second == FrmControlType::Winsock) {
                    if (mWs == "close")        fnWs = "vb6_Ws_Close";
                    else if (mWs == "listen")  fnWs = "vb6_Ws_Listen";
                    else if (mWs == "connect") fnWs = "vb6_Ws_Connect";
                }
                if (fnWs) {
                    std::string hwndWs = cIdent(knownFormControlOriginalNames_.count(comObjExpr_)
                        ? knownFormControlOriginalNames_[comObjExpr_] : comObjExpr_);
                    comObjExpr_.clear();
                    comMemberName_.clear();
                    c_.emitLine(std::string(fnWs) + "((void*)vb6_hwnd_" + hwndWs + ");"
                                "  /* Winsock." + mWs + " (原生 Winsock2) */");
                    return;
                }
            }
            // Fix 086: 无括号的控件方法调用 (List1.Clear) — 与 IndexOrCallExpr
            // 的 P13.3 处理一致, 生成 vb6_ClearList(vb6_hwnd_Listx), 而非
            // vb6_ComCall(list1,...) 裸控制名 (C2065).
            auto itCtrlCS = knownFormControls_.find(comObjExpr_);
            if (itCtrlCS != knownFormControls_.end()
                && (itCtrlCS->second == FrmControlType::ListBox
                    || itCtrlCS->second == FrmControlType::ComboBox)
                && Symbol::toLower(comMemberName_) == "clear") {
                std::string ctrlNameCS = cIdent(knownFormControlOriginalNames_.count(comObjExpr_)
                    ? knownFormControlOriginalNames_[comObjExpr_] : comObjExpr_);
                comObjExpr_.clear();
                comMemberName_.clear();
                c_.emitLine("vb6_ClearList((void*)vb6_hwnd_" + ctrlNameCS + ");  /* ListBox.Clear */");
                return;
            }
            // Fix 185: 同族处理 —— 无括号的 PictureBox 绘制方法 (Picture2.Cls)。
            // 带实参的那条走 cgen_expr_call_callee_withm.inc，两条路径都得覆盖，
            // 否则留下 vb6_ComCall(vb6_hwnd_x, L"Cls", NULL, 0) 这种运行期 no-op。
            if (itCtrlCS != knownFormControls_.end()
                && itCtrlCS->second == FrmControlType::PictureBox) {
                std::string memLowerCS = Symbol::toLower(comMemberName_);
                if (memLowerCS == "cls" || memLowerCS == "print") {
                    std::string ctrlNamePic = cIdent(knownFormControlOriginalNames_.count(comObjExpr_)
                        ? knownFormControlOriginalNames_[comObjExpr_] : comObjExpr_);
                    comObjExpr_.clear();
                    comMemberName_.clear();
                    if (memLowerCS == "cls") {
                        c_.emitLine("vb6_ControlCls((void*)vb6_hwnd_" + ctrlNamePic + ");  /* PictureBox.Cls */");
                    } else {
                        c_.emitLine("vb6_ControlPrint((void*)vb6_hwnd_" + ctrlNamePic + ", 0);  /* PictureBox.Print */");
                    }
                    return;
                }
            }
            // C29-SL-l（账 #143）: **不带括号**的控件零实参方法 —— `Text1.SetFocus` /
            // `Slider1.ClearSel` 这一形由 parser 直接交付 CallStmt(callee=MemberAccessExpr)，
            // 到不了上面那条表达式路，所以在这里用同一张表再拦一次（`controlZeroArgMethod`）。
            // 不接这头的形状是 `vb6_ComCall(裸 HWND, L"SetFocus", NULL, 0)`：对假 IDispatch 发
            // Invoke ⇒ 编得过、链接过、跑起来一声不响，零诊断。
            {
                std::string zaHwnd;
                FrmControlType zaType = FrmControlType::Unknown;
                std::string zaMem = Symbol::toLower(comMemberName_);
                if (formCtrlSlot(comObjExpr_, zaType, zaHwnd)) {
                    std::string zaFn = controlZeroArgMethod(zaType, zaMem);
                    if (!zaFn.empty()) {
                        comObjExpr_.clear();
                        comMemberName_.clear();
                        c_.emitLine(zaFn + "((void*)" + zaHwnd + ");  /* " + zaMem + " */");
                        return;
                    }
                }
            }
            // 无括号的COM方法调用: obj.Method → vb6_ComCall(obj, L"Method", NULL, 0)
            callExpr = "vb6_ComCall(" + comObjExpr_ + ", L\"" + comMemberName_ + "\", NULL, 0)";
            comObjExpr_.clear();
            comMemberName_.clear();
        } else if (callExpr == "0") {
            // M22: void function call returned 0 (no-value), discard entire statement
            return;
        } else if (callExpr.find('(') == std::string::npos) {
            // Fix 010m: Bare call (no parentheses) — build complete arg list
            // Handle: me-prepend for class methods, ParamArray, Optional padding, _has_ flags
            bool calleeHasPA = false;
            std::vector<ParameterInfo> calleeParams;
            // Fix 030b: 跟踪 builtin 状态 — builtin 跳过 padding/IsMissing 尾叜
            bool calleeIsBuiltin = false;

            if (node.callee && node.callee->kind == ASTNodeKind::IdentifierExpr) {
                auto& idExpr = static_cast<IdentifierExpr&>(*node.callee);
                Symbol* sym = symTab_.lookupModule(idExpr.name);
                if (!sym || (sym->kind != SymbolKind::Sub && sym->kind != SymbolKind::Function
                    && sym->kind != SymbolKind::PropertyGet && sym->kind != SymbolKind::PropertyLet
                    && sym->kind != SymbolKind::PropertySet)) {
                    sym = symTab_.lookup(idExpr.name);
                }
                if (sym && (sym->kind == SymbolKind::Sub || sym->kind == SymbolKind::Function
                    || sym->kind == SymbolKind::PropertyGet || sym->kind == SymbolKind::PropertyLet
                    || sym->kind == SymbolKind::PropertySet)) {
                    calleeParams = sym->params;
                    calleeIsBuiltin = sym->isBuiltin;
                }
            }
            // Fix 015: Call X.Y(args).Z (无尾括号) 形态下 node.callee 是 .Z MemberAccessExpr.
            // 此时 Fix 015 emit 出的 lastExpr_ 是裸函数名 "vb6_cDataBase_Exec",
            // 对象参数通过 pendingChainObj_ 传递. 这里需要按成员名查找参数签名,
            // 才能正确填充 Optional 默认参数.
            else if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr) {
                auto& maExpr = static_cast<MemberAccessExpr&>(*node.callee);
                // Fix 092z 姊妹路 (见 cgen_expr_call_callee_params.inc 同名清单) ——
                // VB6 内置全局对象的成员调用不来自当前模块，若拿 lookupModule(memberName)
                // 要形参表会命中同名模块符号 (VBFlexGrid.Public Sub Clear(Optional Where,
                // Optional What))，把 bare `vb6_Clipboard_Clear` 补成 `(0,0,0,0)` → C2197。
                // 两条 padding 路径共用同一张表；改动请同步另一侧的 builtinGlobalObjs092z。
                bool builtinGlobalObjCS = false;
                if (maExpr.object && maExpr.object->kind == ASTNodeKind::IdentifierExpr) {
                    static const std::unordered_set<std::string> builtinGlobalObjsCS = {
                        "clipboard", "screen", "printer", "forms", "debug", "err",
                        "app", "controls", "console"};
                    builtinGlobalObjCS = builtinGlobalObjsCS.count(
                        Symbol::toLower(static_cast<IdentifierExpr&>(*maExpr.object).name)) > 0;
                }
                if (builtinGlobalObjCS) {
                    calleeIsBuiltin = true;
                } else {
                Symbol* sym = symTab_.lookupModule(maExpr.memberName);
                if (sym && (sym->kind == SymbolKind::Sub || sym->kind == SymbolKind::Function
                    || sym->kind == SymbolKind::PropertyGet || sym->kind == SymbolKind::PropertyLet
                    || sym->kind == SymbolKind::PropertySet)) {
                    calleeParams = sym->params;
                    calleeIsBuiltin = sym->isBuiltin;
                }
                // Fix 084y-3: 链式调用对象方法时 (X.Y(args).Z), .Z 是类方法而非
                // 模块成员, lookupModule 必然失败 → calleeParams 为空 → 不填充
                // Optional 默认值 → C2198 参数太少 (如 Exec(Optional RecordsAffected,
                // Optional Options As Long = -1) 声明5参却只传对象1参).
                // 用 inferClassTypeOfExpr 推断对象类 (X.Y(args) → cDataBase),
                // 再按类方法签名取形参表.
                if (calleeParams.empty() && maExpr.object) {
                    std::string chainClass = inferClassTypeOfExpr(*maExpr.object);
                    if (!chainClass.empty()) {
                        std::vector<ParameterInfo> clsParams;
                        bool clsBuiltin = false;
                        if (findClassMemberCallParams(chainClass, maExpr.memberName,
                                                      clsParams, clsBuiltin)) {
                            calleeParams = clsParams;
                            calleeIsBuiltin = clsBuiltin;
                        }
                    }
                }
                }
            }
            // Fix 090s: With 块内无括号类方法调用 (.Start — callee=WithMemberExpr,
            // callExpr 裸函数名, 与 Fix 015 MAE 同协议). visit(WithMemberExpr)
            // asCallCallee_ 已不再拼完整调用而是交付 pendingChainObj_ → 此处
            // 解析形参表才能做 Optional padding, 否则 .Start 声明带 4 个
            // Optional 参只传 this → C2198 参数太少.
            else if (node.callee && node.callee->kind == ASTNodeKind::WithMemberExpr) {
                auto& wmExpr90s = static_cast<WithMemberExpr&>(*node.callee);
                if (!withObjectInfoStack_.empty()) {
                    const auto& wmInfo90s = withObjectInfoStack_.back();
                    if (wmInfo90s.kind == WithObjKind::ClassInstance
                        && !wmInfo90s.className.empty()) {
                        std::vector<ParameterInfo> wmParams90s;
                        bool wmBuiltin90s = false;
                        if (findClassMemberCallParams(wmInfo90s.className,
                                                      wmExpr90s.memberName,
                                                      wmParams90s, wmBuiltin90s)) {
                            calleeParams = wmParams90s;
                            calleeIsBuiltin = wmBuiltin90s;
                        }
                    }
                }
            }

            // Check for ParamArray
            int paIndex = -1;
            for (size_t i = 0; i < calleeParams.size(); i++) {
                if (calleeParams[i].isParamArray) { paIndex = (int)i; calleeHasPA = true; break; }
            }

            // P6.6: 类模块中调用同类方法, 需要自动添加me作为第一个参数
            std::string bareArgList;
            if (isClassModule_ && currentProc_) {
                std::string modPrefix = "vb6_" + cIdent(moduleName_) + "_";
                if (callExpr.find(modPrefix) == 0) {
                    bareArgList = "(void*)me";
                }
            }

            // Fix 015: 链式调用对象参数前置 — 由 visit(MemberAccessExpr).Fix015 交付
            if (!pendingChainObj_.empty()) {
                if (!bareArgList.empty()) bareArgList += ", ";
                bareArgList += pendingChainObj_;
                pendingChainObj_.clear();
            }

            if (calleeHasPA) {
                // ParamArray: pass NULL SAFEARRAY*
                if (!bareArgList.empty()) bareArgList += ", ";
                bareArgList += "NULL";
            } else if (calleeParams.size() > 0 && !calleeIsBuiltin) {
                // Pad all params with default values (bare call = 0 args)
                // Fix 030b: builtin 跳过 padding/IsMissing 路径 (RTL C 签名不接受尾叜)
                for (size_t i = 0; i < calleeParams.size(); i++) {
                    const auto& param = calleeParams[i];
                    if (!bareArgList.empty()) bareArgList += ", ";
                    std::string defVal;
                    if (param.hasDefaultValue && !param.defaultValueExpr.empty()) {
                        defVal = param.defaultValueExpr;
                    } else {
                        defVal = defaultValue(param.type);
                    }
                    if (param.isByVal) {
                        bareArgList += defVal;
                    } else {
                        // ByRef: pass address of compound literal
                        std::string cType = mapType(param.type);
                        if (param.type == Vb6Type::Variant || param.type == Vb6Type::Empty ||
                            param.type == Vb6Type::Null || param.type == Vb6Type::Object) {
                            bareArgList += "&(" + cType + "){0}";
                        } else {
                            bareArgList += "&(" + cType + "){" + defVal + "}";
                        }
                    }
                }
                // Append _has_ flags for Optional params (all 0 since none passed)
                for (size_t i = 0; i < calleeParams.size(); i++) {
                    if (calleeParams[i].isOptional && !calleeParams[i].isParamArray) {
                        if (!bareArgList.empty()) bareArgList += ", ";
                        bareArgList += "0";
                    }
                }
            }
            // Fix 034: Builtin Sub statements with Optional ByVal 参 — 硬编码补默认值.
            // calleeParams 未注册的 builtin (如 Randomize), bare-call 无参时补默认值.
            // Randomize([seed]) — RTL vb6_Randomize(double seed); 不传参时 seed=0.0.
            if (callExpr == "vb6_Randomize" && bareArgList.empty()) {
                bareArgList = "0.0";
            }
            callExpr += "(" + bareArgList + ")";
        }

        // Fix 090g: fallback 合成 (visit MemberAccessExpr asCallCallee_ → func(obj)
        // 完整调用) 产生的单 this 语句调用 — 无括号 MAE 语句 (如 Response.State403,
        // State403 声明 (Optional Say As String) → C 签名 (me, BSTR*, int _has_Say)),
        // callExpr 只含 this → C2198 参数太少. 检测括号内无顶层逗号 (单 this
        // 实参, 用户实参由 IndexOrCallExpr 承载不会以 MAE 形态到此) 后按形参表
        // 重建: this + 各形参默认值 + Optional _has_ 标志 (同 bare-call padding).
        if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr) {
            auto& maExpr090g = static_cast<MemberAccessExpr&>(*node.callee);
            std::string cls090g = inferClassTypeOfExpr(*maExpr090g.object);
            std::vector<ParameterInfo> params090g;
            bool builtin090g = false;
            // Fix <vbeclipse>: 类实例上不存在该方法 (ucPerspective.ctl:1837
            //   `l_ucFolder.ZOrder` — l_ucFolder As ucFolder, UserControl 类的
            //   struct 实例), 成员路径发成 `l_ucFolder->ZOrder /* class var ...
            //   field */` + 调用括号 → C2223 "-> 左侧必须指向结构/联合". 类里没有
            //   该成员函数就根本没有正确 C 形态 (没有 HWND, ZOrder 于容器 z-序不可
            //   建模), 发安全空操作并保留引用, 避免静默丢引用/未用告警.
            if (!cls090g.empty()
                && params090g.empty() && !builtin090g
                && resolveClassMemberCall(cls090g, maExpr090g.memberName).empty()) {
                emitExpr(*maExpr090g.object);
                c_.emitLine("(void)" + lastExpr_ + ";  /* vbeclipse: "
                            + maExpr090g.memberName
                            + " 未建模于 " + cls090g + ", 空操作 */");
                return;
            }
            if (!cls090g.empty()
                && findClassMemberCallParams(cls090g, maExpr090g.memberName,
                                             params090g, builtin090g)
                && !params090g.empty() && !builtin090g) {
                size_t p090g = callExpr.find('(');
                if (p090g != std::string::npos && callExpr.size() >= 3
                    && callExpr.back() == ')') {
                    bool topComma090g = false;
                    int depth090g = 0;
                    for (size_t k090g = p090g; k090g < callExpr.size(); k090g++) {
                        char ch090g = callExpr[k090g];
                        if (ch090g == '(') depth090g++;
                        else if (ch090g == ')') {
                            depth090g--;
                            if (depth090g == 0) break;
                        } else if (ch090g == ',' && depth090g == 1) {
                            topComma090g = true;
                            break;
                        }
                    }
                    if (!topComma090g) {
                        std::string thisArg090g = callExpr.substr(p090g + 1,
                                                                  callExpr.size() - p090g - 2);
                        std::string fullArg090g = thisArg090g;
                        bool anyPad090g = false;
                        for (size_t i090g = 0; i090g < params090g.size(); i090g++) {
                            const auto& prm090g = params090g[i090g];
                            if (prm090g.isParamArray) continue;
                            std::string defVal090g;
                            if (prm090g.hasDefaultValue && !prm090g.defaultValueExpr.empty()) {
                                defVal090g = prm090g.defaultValueExpr;
                            } else {
                                defVal090g = defaultValue(prm090g.type);
                            }
                            if (prm090g.isByVal) {
                                fullArg090g += ", " + defVal090g;
                            } else {
                                std::string cT090g = mapType(prm090g.type);
                                if (prm090g.type == Vb6Type::Variant
                                    || prm090g.type == Vb6Type::Empty
                                    || prm090g.type == Vb6Type::Null
                                    || prm090g.type == Vb6Type::Object) {
                                    fullArg090g += ", &(" + cT090g + "){0}";
                                } else {
                                    fullArg090g += ", &(" + cT090g + "){" + defVal090g + "}";
                                }
                            }
                            anyPad090g = true;
                        }
                        for (size_t i090g = 0; i090g < params090g.size(); i090g++) {
                            if (params090g[i090g].isOptional
                                && !params090g[i090g].isParamArray) {
                                fullArg090g += ", 0";
                            }
                        }
                        if (anyPad090g) {
                            callExpr = callExpr.substr(0, p090g) + "("
                                       + fullArg090g + ")";
                        }
                    }
                }
            }
        }

        if (callExpr.find("vb6_ComCall(") == 0) {
            // ComCall返回可能含对象的VARIANT*, 用VarFree避免Release对象
            c_.emitLine("vb6_ComVarFree((void*)" + callExpr + ");  /* COM call, discard result */");
        } else if (callExpr.find("vb6_ComGetProp(") == 0) {
            c_.emitLine("vb6_ComVarClear((void*)" + callExpr + ");  /* COM prop get, discard result */");
        } else {
            c_.emitLine(callExpr + ";");
        }

        // M22: ANSI临时变量释放由emitStmtList统一处理, 此处不再单独清理
    }
}


} // namespace vb6c3
