// cgen_expr_call_builtin_argtype.inc - visit(IndexOrCallExpr&) 函数体片段：内建函数实参类型适配（Fix 108c）：裸 COM 结果解包 / BSTR 形参数值实参 / ObjPtr 接口取值
// 由 src/backend/detail/expr/cgen_expr_call_pad_conv.inc 拆出（2026-09-19），
// 在函数体内被 #include，非独立编译单元。内容 = 拆分前该文件第 437~546 行，逐行未改。

    // Fix 108c: 内建函数实参的类型适配 (走到此处说明没有专门分支处理).
    //  (a) 数值型内建函数收到"无类型 COM 调用结果" (vb6_ComCall → void*, 指向
    //      VARIANT) 时先解包为数值, 否则 void* 传入 double 形参报 C2440
    //      (Charts 2020 ucChartBar SumSerieValues: vb6_Abs(vb6_ComCall(...))).
    //  (b) 形参为 String 的内建函数收到数值实参时按 VB6 隐式转换转 BSTR,
    //      否则 float 传入 BSTR 形参报 C2440
    //      (Charts 2020 ucChartArea: vb6_Replace(m_LabelsFormats, "{V}", yRange)) .
    {
        static const std::unordered_set<std::string> numericBuiltins108c = {
            "vb6_Abs", "vb6_Int", "vb6_Fix", "vb6_Sgn", "vb6_Sqr", "vb6_Sin",
            "vb6_Cos", "vb6_Tan", "vb6_Atn", "vb6_Log", "vb6_Exp", "vb6_Round",
            "vb6_CDbl", "vb6_CSng", "vb6_CInt", "vb6_CLng", "vb6_CByte",
            "vb6_CBool", "vb6_Hex", "vb6_Oct",
        };
        static const std::unordered_map<std::string, std::vector<size_t>> strArgPos108c = {
            {"vb6_Replace",    {0, 1, 2}},
            {"vb6_StrComp",    {0, 1}},
            {"vb6_StrReverse", {0}},
            {"vb6_Trim",       {0}},
            {"vb6_LTrim",      {0}},
            {"vb6_RTrim",      {0}},
            {"vb6_UCase",      {0}},
            {"vb6_LCase",      {0}},
            {"vb6_StrConv",    {0}},
            // Fix 158t: Mid$/Left$/Right$ 的首参是 String, VB6 允许直接喂数值
            // (隐式 CStr). MainForm.frm:386 `DecStr = Mid$(1.1, 2, 1)` 取本地小数点
            // 分隔符是惯用写法 → 裸发 vb6_Mid(1.1, 2, 1) 报 C2440 "double→BSTR".
            {"vb6_Mid",        {0}},
            {"vb6_Left",       {0}},
            {"vb6_Right",      {0}},
        };
        bool argsFixed108c = false;
        // argList 在更早处已按 args 拼好, 其后还有"补默认参数"分支向 argList 追加
        // (如 vb6_Replace 的 ", 1, -1, 0"). 因此不能整体重建 argList —— 只能就地把
        // 被改写的实参文本替换掉, 否则会丢掉默认参数 (曾导致 vb6_Replace 只剩 3 参).
        size_t searchFrom108c = 0;
        auto replaceInArgList108c = [&](const std::string& oldVal,
                                        const std::string& newVal) {
            size_t p = argList.find(oldVal, searchFrom108c);
            if (p == std::string::npos) return;
            argList.replace(p, oldVal.size(), newVal);
            searchFrom108c = p + newVal.size();
        };
        if (numericBuiltins108c.count(callee)) {
            for (auto& a : args) {
                if (a.compare(0, 12, "vb6_ComCall(") == 0) {
                    std::string wrapped =
                        "vb6_VariantToDouble(vb6_VariantFromComResult(" + a + "))";
                    replaceInArgList108c(a, wrapped);
                    a = wrapped;
                    argsFixed108c = true;
                }
            }
        }
        // Fix 108c-2: vb6_Val 形参为 BSTR, 收到"无类型 COM 调用结果"
        // (vb6_ComCall → void*, 指向 VARIANT) 时解包为字符串, 而非数值 —
        // 若走 numericBuiltins108c 的 vb6_VariantToDouble 包裹, double 传入
        // BSTR 形参报 C2440 (BalloonTooltips cTT.cls: Val(colTooltips(sKey))).
        if (callee == "vb6_Val") {
            for (auto& a : args) {
                bool isComResult = (a.compare(0, 12, "vb6_ComCall(") == 0
                                    || a.compare(0, 12, "vb6_ComGetProp(") == 0);
                if (isComResult) {
                    std::string wrapped =
                        "vb6_VariantToString(vb6_VariantFromComResult(" + a + "))";
                    replaceInArgList108c(a, wrapped);
                    a = wrapped;
                    argsFixed108c = true;
                }
            }
        }
        auto itStrPos = strArgPos108c.find(callee);
        if (itStrPos != strArgPos108c.end()) {
            for (size_t pos : itStrPos->second) {
                if (pos >= args.size() || pos >= node.positional.size()) continue;
                if (!node.positional[pos]) continue;
                Vb6Type at108c = inferExprType(*node.positional[pos]);
                // Fix 158t: Date 也纳入 (Mid$(StartDate, ...) 之类) — CStr(Date)
                // 与 CStr(数值) 不同, 需走 vb6_CStrDate 才能得本地日期字符串.
                if (at108c != Vb6Type::Single && at108c != Vb6Type::Double &&
                    at108c != Vb6Type::Long && at108c != Vb6Type::Integer &&
                    at108c != Vb6Type::Currency && at108c != Vb6Type::Decimal &&
                    at108c != Vb6Type::Boolean && at108c != Vb6Type::Date) {
                    continue;  // 已是字符串/对象等, 不介入
                }
                if (args[pos].find("vb6_CStr") != std::string::npos) continue;
                std::string wrapped108c;
                if (at108c == Vb6Type::Date) {
                    wrapped108c = "vb6_CStrDate((double)(" + args[pos] + "))";
                } else if (at108c == Vb6Type::Long || at108c == Vb6Type::Integer ||
                           at108c == Vb6Type::Boolean) {
                    wrapped108c = "vb6_CStrLong((int32_t)(" + args[pos] + "))";
                } else {
                    wrapped108c = "vb6_CStrDbl((double)(" + args[pos] + "))";
                }
                replaceInArgList108c(args[pos], wrapped108c);
                args[pos] = wrapped108c;
                argsFixed108c = true;
            }
        }

        // Fix 108c-3: vb6_ObjPtr(ifaceVar) — 接口引用是值结构体, 不能直接作为
        // uintptr_t 实参 (C2172). 对象指针位于 .obj 字段: 改写为 (<var>).obj.
        // mdlSubclass.bas: vb6_ObjPtr(Subclass) → vb6_ObjPtr((Subclass).obj).
        // 仅精确匹配裸接口变量名, 不触碰其它形式 (LongPtr 变量等走别处转换).
        if (callee == "vb6_ObjPtr" && !args.empty() && !node.positional.empty()
            && node.positional[0] && node.positional[0]->kind == ASTNodeKind::IdentifierExpr) {
            std::string aLower = args[0];
            std::transform(aLower.begin(), aLower.end(), aLower.begin(), ::tolower);
            if (knownIfaceVars_.count(aLower)) {
                std::string newVal = "(" + args[0] + ").obj";
                replaceInArgList108c(args[0], newVal);
                args[0] = newVal;
                argsFixed108c = true;
            }
        }
        (void)argsFixed108c;
    }

    // Fix <vbeclipse> 2026-10-06: Declare 的 Err.LastDllError 快照**不在这里**做 ——
    // 调用点注入语句会把 `While Api() <> 0` / `IIf(Api(), ...)` 这类表达式位置的
    // 调用撕成 `(Api(); vb6_ErrSetLastDllError(...) != 0)` 这种非法 C。快照改由
    // cgen_decl_api.cpp 为每个 Declare 生成的 static __inline 包装函数承担
    // (调用点经 #define / declareAliasMap_ 已被重定向到该包装), 这里保持原样。
    lastExpr_ = callee + "(" + argList + ")";
