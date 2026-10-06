// cgen_expr_call_conv_cstr.inc - visit(IndexOrCallExpr&) 函数体片段：CStr 家族适配（按实参类型选 CStr / CStrSingle / CStrLong / CStrDbl）
// 由 src/backend/detail/expr/cgen_expr_call_pad_conv.inc 拆出（2026-09-19），
// 在函数体内被 #include，非独立编译单元。内容 = 拆分前该文件第 82~164 行，逐行未改。

    // M22-fix: CStr类型适配 — 根据参数类型选择正确的CStr变体
    if (callee == "vb6_CStr" && !node.positional.empty()) {
        auto& firstArg = node.positional[0];
        if (firstArg->kind == ASTNodeKind::IdentifierExpr) {
            auto& idArg = static_cast<IdentifierExpr&>(*firstArg);
            std::string argLower = idArg.name;
            std::transform(argLower.begin(), argLower.end(), argLower.begin(), ::tolower);
            // 参数是BSTR变量 → CStr是空操作, 直接使用参数
            if (knownBstrVars_.count(argLower)) {
                emitExpr(*firstArg);
                return;
            }
            // Fix 175: Date 变量 → vb6_CStrDate (必须先判, Date 同时登记在
            // knownDoubleVars_ 里, 否则 CStr(d) 会按 double 打成序列号)
            if (knownDateVars_.count(argLower)) {
                callee = "vb6_CStrDate";
            }
            // Fix 117c: 参数是Single(float)变量 → vt_R4 语义 (必须优先于 Double 判定,
            // 因为 Single 变量同时登记在 knownDoubleVars_ 里以复用既有取值路径)
            else if (knownSingleVars_.count(argLower)) {
                callee = "vb6_CStrSingle";
            }
            // 参数是int32_t/Long变量 → 用vb6_CStrLong
            // ai/022 W1: 布尔形参/局部此前落在 Long 组 → CStr(b) 打成 -1。
            else if (knownBoolVars_.count(argLower)) {
                callee = "vb6_CStrBool";
            }
            else if (knownLongVars_.count(argLower)) {
                callee = "vb6_CStrLong";
            }
            // 参数是double变量 → 用vb6_CStrDbl
            else if (knownDoubleVars_.count(argLower)) {
                callee = "vb6_CStrDbl";
            }
            // 参数是Variant变量 → 保留vb6_CStr(VARIANT)
        } else {


            // 非标识符表达式: 推断类型选择CStr变体
            Vb6Type argType = inferExprType(*firstArg);
            if (argType == Vb6Type::String) {
                // Fix 056b: CStr(String) 是空操作 → 直接使用参数
                // (否则生成 vb6_CStr(BSTR) → C2440: 无法从BSTR转换为vb6_VARIANT)
                if (!args.empty()) { lastExpr_ = args[0]; return; }
                lastExpr_ = "vb6_BSTR_Empty()";
                return;
            }
            if (argType == Vb6Type::LongPtr) {
                // 账 #180: 句柄档 (LongPtr) 的实参在 C 侧是 void* / intptr_t。落到下面
                // 那条 `vb6_CStr(vb6_VariantFromValue(x))` 会命中 _Generic 的
                // `default: vb6_VariantObject` (vb6rtl_variant.h:215) ⇒ CStr(句柄) 打成
                // **空串**, 而 VB6 里 ContainerHwnd / hWnd 是 Long、CStr 打的是数字。
                // 不能往 _Generic 补一档 void* : 真对象引用在 C 层也是 void*, 那样会把
                // 对象一律打成数字 (装箱那一路比句柄更常见) ⇒ 只能在**类型已知**的这里分诊。
                if (!args.empty()) {
                    lastExpr_ = "vb6_CStrLongLong((int64_t)(intptr_t)(" + args[0] + "))";
                    return;
                }
            }
            if (argType == Vb6Type::Long || argType == Vb6Type::Integer) {
                callee = "vb6_CStrLong";
            } else if (argType == Vb6Type::Boolean) {
                callee = "vb6_CStrBool";   // ai/022 W1: VB6 的 CStr(True) = "True"
            } else if (argType == Vb6Type::Date) {
                callee = "vb6_CStrDate";   // Fix 175: 日期按 VB6 短日期串输出
            } else if (argType == Vb6Type::Single) {
                callee = "vb6_CStrSingle";   // Fix 117c: VT_R4 语义
            } else if (argType == Vb6Type::Double) {
                callee = "vb6_CStrDbl";
            }
            // P25: COM属性取值已在args[0]中, 根据取值函数选择CStr变体或跳过
            if (!args.empty()) {
                const std::string& a0 = args[0];
                if (a0.find("vb6_ComGetStringProp") == 0 || a0.find("vb6_ComCallBSTR") == 0) {
                    // 已是BSTR, CStr是空操作
                    lastExpr_ = a0;
                    return;
                } else if (a0.find("vb6_ComGetIntProp") == 0 || a0.find("vb6_ComVtableGetInt") == 0) {
                    callee = "vb6_CStrLong";
                } else if (a0.find("vb6_ComGetDoubleProp") == 0 || a0.find("vb6_ComVtableGetDouble") == 0) {
                    callee = "vb6_CStrDbl";
                }
            }
        }
        // Fix 056b: callee 仍为 vb6_CStr 且实参非Variant → 用 vb6_VariantFromValue 包装
        // (如 CStr((int32_t)GetCurrentThreadId()) / CStr(模块级Long变量) →
        //  vb6_CStr(vb6_VariantFromValue(x)), 由 _Generic 按实参类型自动包装, 消除 C2440)
        if (callee == "vb6_CStr" && !args.empty()) {
            // Fix 121: 实参是**裸 COM 结果** (vb6_ComCall/ComCallByDispid/ComGetProp/
            // ComGetObjectProp 返回 void*, 承载 COM VARIANT*) 时必须先用
            // vb6_VariantFromComResult 解引用成 vb6_VARIANT — vb6_VariantFromValue
            // 会按 _Generic 匹配 void* → 得到空 Variant, 于是 CStr(Collection.Item(i))
            // 打印空串; 图表把系列值读成空/噪声 → 坐标轴量级完全错
            // (Charts 2020 柱图/面积图 Y 轴数字大 100 倍; ferock 0.10.6 同修复,
            //  此处合并为并集: 含 ComCallByDispid/ComGetObjectProp)。
            const std::string& a0 = args[0];
            bool comRes121 = (a0.compare(0, 12, "vb6_ComCall(") == 0)
                || (a0.compare(0, 21, "vb6_ComCallByDispid(") == 0)
                || (a0.compare(0, 15, "vb6_ComGetProp(") == 0)
                || (a0.compare(0, 21, "vb6_ComGetObjectProp(") == 0);
            if (comRes121) {
                lastExpr_ = "vb6_CStr(vb6_VariantFromComResult(" + a0 + "))";
            } else {
                lastExpr_ = "vb6_CStr(vb6_VariantFromValue(" + a0 + "))";
            }
            return;
        }
    }
