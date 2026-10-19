// cgen_expr_call_conv_output.inc - visit(IndexOrCallExpr&) 函数体片段：输出侧适配：Format 首参装箱 + CStr 兜底包装（MsgBox 那一腿已并入 Fix 175, 见下）
// 由 src/backend/detail/expr/cgen_expr_call_pad_conv.inc 拆出（2026-09-19），
// 在函数体内被 #include，非独立编译单元。内容 = 拆分前该文件第 249~353 行，逐行未改。

    // 账 #305 §B175 (第六十六刀): 原此处的 MsgBox prompt 折算段 (六条文本文案连判 +
    // Fix 161e 那条 wrapToBSTR 兜底) 与紧随其后的 argList 重建整段撤掉 —— 那六是**同一问**
    // (「实参档位与形参档位不符 ⇒ 折算」) 的第二个主人, 而它吃的是 args[] 的文本前缀。
    // 现在这一问由 Fix 175 在 arg_emit 那一站答 (位号问 builtinNoParamTableStringSlot,
    // 形状走 wrapToBSTR), 且实测**覆盖面一字不差**: BASE(段在) vs OFF(段撤) 跑 tests/ 全
    // 422 份输入, 移动的 6 行全在 tests/msgbox_num.bas 且全是那六条文本案里的数值一档
    // (`MsgBox n` / `MsgBox 123` / `MsgBox (n+1)` / `MsgBox d` / `MsgBox True` / `MsgBox Len(...)`),
    // 而 COM 结果 / Variant / 已折过 那三条案在段撤后产物**逐字节相同** (= 那三条今天已由
    // 别的站办完, 这一段在语料上只承数值一档) —— 读数 `.build/k66_ab_report.txt`。
    // 顺带被这一段管不着的旧账: MsgBox 的**标题位**是 BSTR, 而它只管 prompt ⇒ 今天仍裸送
    // (`vb6_MsgBox(s1, 0, d2)`), 新站按名单 {0,2} 一起管。
    // 下面那三段 (MsgBox arity 改道 / Format 首参装箱 / CStr 兜底) 问的是**另一件事**
    // (补默认实参、把值装进 VARIANT 形参、按实参类型挑 CStr 出口), 与本问无关, 一字未动。

    // MsgBox(prompt) -> vb6_MsgBox1(prompt)
    // MsgBox(prompt, buttons) -> vb6_MsgBox(prompt, buttons, NULL)
    if (callee == "vb6_MsgBox") {
        if (node.positional.size() == 1) {
            callee = "vb6_MsgBox1";
        } else if (node.positional.size() == 2) {
            argList += ", NULL";
        }
    }
    // P26: Format 第一个参数需要包装为 vb6_VARIANT
    if (callee == "vb6_Format" && !args.empty()) {
        Vb6Type argType = inferExprType(*node.positional[0]);
        if (argType == Vb6Type::Long || argType == Vb6Type::Integer) {
            args[0] = "vb6_VariantLong(" + args[0] + ")";
        } else if (argType == Vb6Type::Double || argType == Vb6Type::Single) {
            args[0] = "vb6_VariantDouble(" + args[0] + ")";
        } else if (argType == Vb6Type::String) {
            args[0] = "vb6_VariantString(" + args[0] + ")";
        } else if (argType == Vb6Type::Boolean) {
            // Fix 198: 原为 vb6_VariantInt((int16_t)x) —— 那是把 VB 的 Boolean 当
            // C 的 short 装箱 (VT_I2), 于是 Format(True,"G") 打出 -1. 换真 VT_BOOL.
            args[0] = boxToVariant(node.positional[0].get(), args[0]);
        } else if (argType == Vb6Type::Byte) {
            args[0] = "vb6_VariantInt((int16_t)(" + args[0] + "))";
        } else if (argType == Vb6Type::Date) {
            // 账 #430（§B204）: 这一臂原来**手抄**了一张"VB 类型 → 装箱出口"的名单, 其中
            // Date 那一行写的是 vb6_VariantDouble((double)(...)) —— 与唯一权威 boxToVariant
            // 对同一问的答案 vb6_VariantDate(...) **相反**。两位主人两个答案, 而权威那条从
            // 这一站**从没被走到**。跟已结的 #428/#431 同一条教训: 表的格子就是口径, 抄表的
            // 人不会知道自己抄错了。
            // 实测后果: 一枚 Dim d As Date 局部发成 double, 每一发 Format(d, …) 都贴
            // VT_R8, 可同一枚 VarType(d) 答 7 ⇒ "这枚值真是日期"在运行期根本读不出来。
            // 改法 = 改问权威(与上面 Boolean 档 Fix 198 同一个做法, 一处、零新形状)。
            // ⚠ **只动这一臂**: Long/Integer/Double/Single/String/Byte 那几档与权威**也各不同**
            //   (它们多数走 VariantFromValue, 而这站刻意发 VariantLong/VariantDouble/…),
            //   但那些都是**已登记的既存形状**, 且数值臂还得留给真数值
            //   (Format(43831, "mmm") 那一族, 见 vb6rtl_format_parse.inc:37-44)——
            //   不在 #430 范围内。**别顺手塌成一句 boxToVariant**, 那会一次性移动六档形状。
            args[0] = boxToVariant(node.positional[0].get(), args[0]);
        }
        // Fix 036: Format fallback — 未匹配类型 (Currency/Unknown 等, 非 Variant) 用
        // vb6_VariantFromValue 包装. Variant 类型无需包装 (已是 vb6_VARIANT).
        else if (argType != Vb6Type::Variant) {
            // 装箱走 boxToVariant 权威 (上面 Boolean 档已经是它)
            args[0] = boxToVariant(node.positional[0].get(), args[0]);
        }
        argList.clear();
        for (size_t i = 0; i < args.size(); i++) {
            if (i > 0) argList += ", ";
            argList += args[i];
        }
        // <vbeclipse>: Format(expression[, format[, FirstDayOfWeek[, FirstWeekOfYear]]])
        // RTL 签名 vb6_Format(vb6_VARIANT, BSTR, int32_t, int32_t) 与 VB6 完整 4 形对齐
        // (VBFlexGrid.ctl:20425/20428 直接 `Format$(Text, Col.Format, vbUseSystemDayOfWeek, vbUseSystem)`,
        // 2 参签会 C2197 too many args)。缺省 FirstDayOfWeek=vbSunday(1), FirstWeekOfYear=vbFirstJan1(1)。
        // 省略 format 那一形补 NULL ("无格式串"，RTL 内部各处也传 NULL, vb6rtl_conv.c:181+)。
        while (args.size() < 4) {
            if (args.size() == 1) { args.push_back("NULL"); }
            else                 { args.push_back("1"); }
            if (!argList.empty()) argList += ", ";
            argList += args.back();
        }
    }
    // Fix 036: CStr fallback — 当特殊分支未匹配 (Boolean/Byte/Date/Currency/常量/
    // 未注册函数返回值/Const BSTR 等), callee 仍为 vb6_CStr (接 vb6_VARIANT). 用
    // vb6_VariantFromValue 包装 args[0], _Generic 按实参 C 类型自动选择 Variant ctor,
    // 避免 concrete → VARIANT C2440. 对已是 vb6_VARIANT 的实参为 identity (no-op), 安全.
    // 但需排除确定 Variant 的表达式 (如 Me.Segments(i) 返回 Variant): _Generic 宏
    // 对某些 Variant 表达式展开可能产生逗号问题 → C2197, 故用 isDefinitelyVariantExpr
    // 跳过, 保留 vb6_CStr(variantExpr) 原样 (vb6_CStr 接 VARIANT, 直接可用).
    if (callee == "vb6_CStr" && !args.empty()) {
        bool argIsDefVariant = false;
        if (!node.positional.empty()) {
            argIsDefVariant = isDefinitelyVariantExpr(*node.positional[0]);
        }
        // Fix 038b-4: 字符串级 Variant 检测 — 如果是 C 级 Variant 表达式,
        // 也跳过 VariantFromValue 包装 (vb6_CStr 直接接 VARIANT)
        if (!argIsDefVariant) {
            argIsDefVariant = cExprIsVariant(args[0]);
        }
        if (!argIsDefVariant) {
            // Fix 121b: 同 Fix 121 — 裸 COM 结果必须解引用 (VariantFromComResult),
            // 否则 _Generic 按 void* 包装得到空 Variant。
            bool comRes121b = (args[0].compare(0, 12, "vb6_ComCall(") == 0)
                || (args[0].compare(0, 20, "vb6_ComCallByDispid(") == 0)
                || (args[0].compare(0, 15, "vb6_ComGetProp(") == 0)
                || (args[0].compare(0, 21, "vb6_ComGetObjectProp(") == 0);
            args[0] = comRes121b
                ? "vb6_VariantFromComResult(" + args[0] + ")"
                : boxToVariant(node.positional.empty() ? nullptr
                                 : node.positional[0].get(), args[0]);
            argList.clear();
            for (size_t i = 0; i < args.size(); i++) {
                if (i > 0) argList += ", ";
                argList += args[i];
            }
        }
    }
