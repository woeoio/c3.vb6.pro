// cgen_expr_call_conv_varargs_lenb.inc - visit(IndexOrCallExpr&) 函数体片段：CallByName 变参打包 + LenB 类型分派（UDT / UDT 数组元素 / 数值）
// 由 src/backend/detail/expr/cgen_expr_call_pad_conv.inc 拆出（2026-09-19），
// 在函数体内被 #include，非独立编译单元。内容 = 拆分前该文件第 354~436 行，逐行未改。

    // Fix 040d: CallByName variadic args packing
    // C signature: vb6_CallByName(void* obj, const wchar_t* procName, int32_t callType,
    //                              void* args, int32_t argc)
    // First 3 VB6 args map to obj/procName/callType. Extra args (4+) are method
    // arguments, packed into (void*[]){vb6_ComPackValue(arg), ...} with argc.
    if (callee == "vb6_CallByName" && args.size() >= 3) {
        if (args.size() == 3) {
            // No extra method args: pass NULL, 0
            argList = args[0] + ", " + args[1] + ", " + args[2] + ", NULL, 0";
        } else {
            // Pack extra args [3..] into void*[] array using vb6_ComPackValue
            std::string argsArray = "(void*[]){";
            for (size_t i = 3; i < args.size(); i++) {
                if (i > 3) argsArray += ", ";
                // B5: 第 4 个起是方法实参; 带型本工程类实例变量按**借用型**打包
                // (packArgValue 的 B5 格), 其余仍是 ComPackValue 原路.
                argsArray += (i < node.positional.size())
                    ? packArgValue("vb6_ComPackValue", args[i], node.positional[i].get())
                    : ("vb6_ComPackValue(" + args[i] + ")");
            }
            argsArray += "}";
            int32_t extraArgc = (int32_t)args.size() - 3;
            argList = args[0] + ", " + args[1] + ", " + args[2] + ", "
                    + argsArray + ", " + std::to_string(extraArgc);
        }
    }
    // Fix 065: LenB(UDT) — MSVC _Generic 对自定义结构体类型匹配有问题,
    // 直接生成 sizeof(vb6_type_XXX) 而不经 vb6_LenB _Generic 宏
    if (callee == "vb6_LenB" && args.size() == 1) {
        // 检查参数是否是已知 UDT 变量
        std::string argLower = args[0];
        // 去除可能的 With 前缀: _vb6_with_N. → 提取变量名
        // 去除 me-> 前缀
        if (argLower.substr(0, 4) == "me->") argLower = argLower.substr(4);
        std::transform(argLower.begin(), argLower.end(), argLower.begin(), ::tolower);
        auto udtIt = knownUdtVars_.find(argLower);
        if (udtIt != knownUdtVars_.end()) {
            lastExpr_ = "(int32_t)sizeof(" + udtIt->second + ")";
            return;
        }
        // Fix 160w: LenB(UDT变量.字段) — arg 是字段链 "pbitmapinfo.bmiheader"
        // (VBFlexGrid.ctl: LenB(pBitmapInfo.BMIHeader)), 基名在 knownUdtVars_ 而
        // 字段名不在 (knownUdtVars_ 只登记变量级). Fix 065 精确全串匹配失败,
        // M22 之后落入 _Generic 宏 → MSVC 对所有分支做转换检查含 struct→BSTR
        // → C2440 "无法从 vb6_type_BITMAPINFOHEADER 转换到 BSTR". 取基名命中
        // 已知 UDT 即按 sizeof(基 UDT) 处理 (LenB 语义 = 该类型栈宽度, 足够).
        {
            size_t dot160w = argLower.find('.');
            if (dot160w != std::string::npos) {
                std::string base160w = argLower.substr(0, dot160w);
                auto itBase160w = knownUdtVars_.find(base160w);
                if (itBase160w != knownUdtVars_.end()) {
                    lastExpr_ = "(int32_t)sizeof(" + itBase160w->second + ")";
                    return;
                }
            }
        }
        // Fix 084h: UDT 数组元素 LenB(VB6_SA_AT(vb6_type_XXX, arr, idx)) →
        // sizeof(vb6_type_XXX)。_Generic 宏无法匹配自定义结构体类型, 且
        // 数组元素类型未知, 直接由类型名生成 sizeof。
        {
            std::string rawArg = args[0];
            std::string rawLower = rawArg;
            std::transform(rawLower.begin(), rawLower.end(), rawLower.begin(), ::tolower);
            const char* saPrefix = "vb6_sa_at(vb6_type_";
            if (rawLower.compare(0, 19, saPrefix) == 0) {
                size_t comma = rawArg.find(',');
                if (comma != std::string::npos) {
                    std::string typeName = rawArg.substr(19, comma - 19);
                    lastExpr_ = "(int32_t)sizeof(" + typeName + ")";
                    return;
                }
            }
        }
        // Fix 125: LenB(数值实参) — vb6_LenB 的 _Generic 宏在 MSVC 下对**所有**
        // 关联分支都做实参→形参的隐式转换检查 (double→BSTR 非法 → C2440, 即使
        // 该分支未被选中; 独立最小复现在 cl /std:c11 下确认). 数值/日期/Currency
        // 实参 (VB6 LenB 语义 = 该类型栈宽度) 必须在生成器侧直接生成 sizeof,
        // 不能落入 _Generic 宏. 例: cZipArchive LenB(uHdr.Ext.CSize) (Currency).
        if (!node.positional.empty()) {
            Vb6Type tLenB125 = inferExprType(*node.positional[0]);
            switch (tLenB125) {
                case Vb6Type::Currency:
                case Vb6Type::Double:
                case Vb6Type::Date:
                    lastExpr_ = "(int32_t)sizeof(double)";
                    return;
                case Vb6Type::Single:
                    lastExpr_ = "(int32_t)sizeof(float)";
                    return;
                case Vb6Type::Long:
                case Vb6Type::Integer:
                case Vb6Type::Boolean:
                case Vb6Type::Error:
                    lastExpr_ = "(int32_t)sizeof(int32_t)";
                    return;
                case Vb6Type::Byte:
                    lastExpr_ = "(int32_t)sizeof(uint8_t)";
                    return;
                default:
                    break;  // BSTR/Variant/UDT/未知 → 走既有路径
            }
        }
    }
