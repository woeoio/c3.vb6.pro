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
            // ai/032: SByte/UInteger/ULong/ULongLong 变量 —— 专属表按档分诊。
            // ULong/ULongLong 必须走无符号版, 否则落到下面 vb6_CStr(vb6_VariantFromValue(x))
            // 时 _Generic 的 `unsigned int: vb6_VariantLong` 会把 &HFFFFFFFF 打成 "-1"。
            else if (auto itNI = knownNarrowIntVars_.find(argLower);
                     itNI != knownNarrowIntVars_.end()) {
                switch (itNI->second) {
                    case Vb6Type::ULong:     callee = "vb6_CStrULong";     break;
                    case Vb6Type::ULongLong: callee = "vb6_CStrULongLong"; break;
                    case Vb6Type::LongLong:  callee = "vb6_CStrLongLong";  break;
                    // 值域 ⊆ int32_t ⇒ 有符号/无符号十进制同字面 ("-128" / "65535")
                    case Vb6Type::SByte:
                    case Vb6Type::UInteger:  callee = "vb6_CStrLong";      break;
                    default: break;
                }
            }
            // 参数是Variant变量 → 保留vb6_CStr(VARIANT)
            //
            // ai/032 兜底: 标识符不在任何 known 表里时 (模块级 ULong/ULongLong —— 它们
            // 的登记点是声明期那一张, 而专属表随过程入口 clear()), 用语义类型再分诊
            // 无符号两档。不做这一步就只能落 vb6_CStr(vb6_VariantFromValue(x))。
            if (callee == "vb6_CStr" && !args.empty()) {
                Vb6Type itNI2 = inferExprType(*firstArg);
                if (itNI2 == Vb6Type::ULong)          callee = "vb6_CStrULong";
                else if (itNI2 == Vb6Type::ULongLong) callee = "vb6_CStrULongLong";
                else if (itNI2 == Vb6Type::SByte
                         || itNI2 == Vb6Type::UInteger) callee = "vb6_CStrLong";
            }
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
                // 账 #230 第二十八刀 (§B136): 但「类型答 LongPtr」不等于「实参在 C 侧是标量」——
                // 属性读没登记具名出口时落 COM 兜底, 交出的是 vb6_VARIANT (实测
                // `CStr(soloT.hDC)` 发成 (int64_t)(intptr_t)(vb6_VariantFromComResult(…))
                // ⇒ C2440 + C2198)。所以这一档也要先问下面 Fix 225 那一句: 已是 VARIANT
                // 就什么也不转, 交给默认那条 vb6_CStr 按档案自己拆。
                if (!args.empty() && !cExprIsVariantCarrier(args[0])) {
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
            // ai/032: 无符号/窄整型表达式 (函数返回 ULong/ULongLong、位移结果等)。
            // ULong/ULongLong 必须走无符号版 —— 落 vb6_CStr(vb6_VariantFromValue(x))
            // 会被 _Generic 的有符号档打成负数。
            else if (argType == Vb6Type::ULong)     callee = "vb6_CStrULong";
            else if (argType == Vb6Type::ULongLong) callee = "vb6_CStrULongLong";
            else if (argType == Vb6Type::SByte
                     || argType == Vb6Type::UInteger) callee = "vb6_CStrLong";
            // P25: COM属性取值已在args[0]中, 根据取值函数选择CStr变体或跳过
            if (!args.empty()) {
                const std::string& a0 = args[0];
                if (a0.find("vb6_ComGetStringProp") == 0 || a0.find("vb6_ComCallBSTR") == 0) {
                    Vb6Type cstrArgType = inferExprType(*firstArg);
                    if (cstrArgType == Vb6Type::Date && a0.find("vb6_ComGetStringProp") == 0) {
                        std::string fixed = a0;
                        fixed.replace(0, 23, "vb6_ComGetDoubleProp");
                        lastExpr_ = "vb6_CStrDate(" + fixed + ")";
                        return;
                    }
                    lastExpr_ = a0;
                    return;
                } else if (a0.find("vb6_ComGetIntProp") == 0 || a0.find("vb6_ComVtableGetInt") == 0) {
                    callee = "vb6_CStrLong";
                } else if (a0.find("vb6_ComGetDoubleProp") == 0 || a0.find("vb6_ComVtableGetDouble") == 0) {
                    callee = "vb6_CStrDbl";
                }
            }

            // Fix 225: 实参在 C 层已经是 **VARIANT 值** (后期绑定属性读 / COM 调用结果, 形如
            // `vb6_VariantFromComResult(vb6_ComGetProp(_vb6_with_7, L"Width"))`) 时, 上面
            // "按推断类型选 vb6_CStrSingle / CStrLong / CStrDbl" 全部作废 —— 那几个函数吃的
            // 是**已拆箱**的 C 标量, 把 vb6_VARIANT 结构体直接交给 float 形参 ⇒ C2440
            // "无法从 vb6_VARIANT 转换为 float" (ComCtlsDemo PPCoolBarBands.c:321-…
            //  `TextCX.Text = CStr(.Width)` / CStr(.MinWidth) / CStr(.IdealWidth) 等 ×61)。
            // VB6 里 CStr(<Variant>) 的答案本就是 Variant 那一档 ⇒ 回到默认 vb6_CStr
            // (形参即 vb6_VARIANT, 见 vb6rtl_builtin.h), 由下方 Fix 056b 那条路收尾:
            //   vb6_CStr(vb6_VariantFromValue(a0)) —— _Generic 对 vb6_VARIANT 是 identity
            //   (vb6_VariantIdentity), 故实际就是 vb6_CStr(<VARIANT>), 正确。
            // 口径与上方 54 行"参数是 Variant 变量 → 保留 vb6_CStr(VARIANT)"一致, 只是那条
            // 分支只对 IdentifierExpr 生效; 这里补的是 MemberAccess / WithMember 等非标识符形态。
            if (!args.empty() && cExprIsVariant(args[0])) {
                callee = "vb6_CStr";
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
                || (a0.compare(0, 20, "vb6_ComCallByDispid(") == 0)
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
