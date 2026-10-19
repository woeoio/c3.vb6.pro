// cgen_expr_call_builtin_pre.inc - visit(IndexOrCallExpr&) 函数体片段：内置函数前置拦截（IIf / VarPtr / TypeName(Me) / Choose / Array / IsMissing / Len）
// 由 src/backend/expr/cgen_expr_call.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 300~620 行，逐行未改。

    // C3 扩展 (ai/032): `If(cond, truePart, falsePart)` 三元运算符。
    //
    // 与紧邻下面的 IIf 是**两条不同的路**, 不要合并:
    //   · `IIf` 是 VB6 内置**函数** —— 实参在调用前全部求值, 两个分支都会跑;
    //   · `If` 是 VB.NET 的**运算符** —— 只对选中的那一支求值。
    // 差别是实打实的: `If(d <> 0, a / d, 0)` 在 d = 0 时不除零, 换成 IIf 会。
    // C 的三目 `?:` 天生短路, 所以这里直接发 `((cond) != 0 ? t : f)`, 连运行期
    // 辅助函数都不需要 —— 而 IIf 因为是函数语义, 必须走 vb6_IIfBSTR/Long/Double/…
    // 那一族 "两支都先算好再选" 的调用。
    if (node.callee && node.callee->kind == ASTNodeKind::IdentifierExpr &&
        node.positional.size() == 3) {
        auto& identIf = static_cast<IdentifierExpr&>(*node.callee);
        std::string ifLower = identIf.name;
        std::transform(ifLower.begin(), ifLower.end(), ifLower.begin(), ::tolower);
        if (ifLower == "if") {
            emitExpr(*node.positional[0]);
            std::string ifCond = std::move(lastExpr_);
            emitExpr(*node.positional[1]);
            std::string ifTrue = std::move(lastExpr_);
            emitExpr(*node.positional[2]);
            std::string ifFalse = std::move(lastExpr_);
            // VB6 的 True = -1, 判真按 "非 0", 与 IIf 那一支同一口径。
            // 两支的外层括号是必要的: 实参顺序影响 C 的求值顺序, 但更关键的是
            // 分支本身可能是逗号表达式外的低优先级运算 (`a = b`), 不括会改写结合性。
            lastExpr_ = "(((" + ifCond + ") != 0) ? (" + ifTrue + ") : (" + ifFalse + "))";
            return;
        }
    }

    // P14.2.4: IIf特殊处理 - 调用类型化RTL函数(函数调用语义确保两个分支都求值,符合VB6规范)
    if (node.callee && node.callee->kind == ASTNodeKind::IdentifierExpr && node.positional.size() == 3) {
        auto& ident = static_cast<IdentifierExpr&>(*node.callee);
        std::string iifLower = ident.name;
        std::transform(iifLower.begin(), iifLower.end(), iifLower.begin(), ::tolower);
        if (iifLower == "iif") {
            emitExpr(*node.positional[0]);
            std::string cond = std::move(lastExpr_);
            // VB6 True=-1, C needs explicit !=0
            cond = "((" + cond + ") != 0)";
            emitExpr(*node.positional[1]);
            std::string trueVal = std::move(lastExpr_);
            emitExpr(*node.positional[2]);
            std::string falseVal = std::move(lastExpr_);
            // Type dispatch: BSTR > double > long
            auto isBstrResult = [&](const std::string& e) -> bool {
                if (e.compare(0, 8, "vb6_BSTR") == 0) return true;
                if (e.find("VB6_SA_AT(BSTR,") != std::string::npos) return true;
                static const char* bstrPfx[] = {"vb6_BSTR_FromStr","vb6_Left","vb6_Right","vb6_Mid",
                    "vb6_UCase","vb6_LCase","vb6_Trim","vb6_LTrim","vb6_RTrim","vb6_Chr",
                    "vb6_Str","vb6_CStr","vb6_Format","vb6_Replace","vb6_Space","vb6_String",
                    "vb6_Command","vb6_CurDir","vb6_Environ","vb6_Dir",
                    "vb6_IIfBSTR","vb6_InputBox","vb6_App_Path","vb6_App_EXEName","vb6_App_HelpFile",
                    "vb6_Console_ReadLine","vb6_Console_ReadKey",  // Fix 161: Console 输入返回 BSTR
                    "vb6_GetControlText","vb6_GetControlCaption",nullptr};
                for (int i = 0; bstrPfx[i]; i++)
                    if (e.compare(0, strlen(bstrPfx[i]), bstrPfx[i]) == 0) return true;
                std::string low = e;
                std::transform(low.begin(), low.end(), low.begin(), ::tolower);
                return knownBstrVars_.count(low) > 0;
            };
            // Fix 046: IIf Variant arg extraction — when typed IIf function
            // is selected but an arg is a Variant, extract the concrete value.
            auto iifArgIsVariant = [&](int idx) -> bool {
                const std::string& val = (idx == 1) ? trueVal : falseVal;
                if (cExprIsVariant(val)) return true;
                auto& arg = node.positional[idx];
                if (arg->kind == ASTNodeKind::IdentifierExpr) {
                    auto& idArg = static_cast<IdentifierExpr&>(*arg);
                    std::string argLower = idArg.name;
                    std::transform(argLower.begin(), argLower.end(), argLower.begin(), ::tolower);
                    if (knownVariantVars_.count(argLower)) return true;
                }
                return false;
            };
            auto extractVariantArg = [&](std::string& val, int idx, const char* extractFn) {
                if (iifArgIsVariant(idx)) {
                    val = std::string(extractFn) + "(" + val + ")";
                }
            };
            if (isBstrResult(trueVal) || isBstrResult(falseVal)) {
                extractVariantArg(trueVal, 1, "vb6_VariantToString");
                extractVariantArg(falseVal, 2, "vb6_VariantToString");
                lastExpr_ = "vb6_IIfBSTR(" + cond + ", " + trueVal + ", " + falseVal + ")";
            } else if (trueVal.find('.') != std::string::npos || falseVal.find('.') != std::string::npos) {
                extractVariantArg(trueVal, 1, "vb6_VariantToDouble");
                extractVariantArg(falseVal, 2, "vb6_VariantToDouble");
                lastExpr_ = "vb6_IIfDouble(" + cond + ", " + trueVal + ", " + falseVal + ")";
            } else if (iifArgIsVariant(1) && iifArgIsVariant(2)) {
                // Both args are Variant — use vb6_IIfVariant (accepts VARIANT args)
                lastExpr_ = "vb6_IIfVariant(" + cond + ", " + trueVal + ", " + falseVal + ")";
            } else {
                extractVariantArg(trueVal, 1, "vb6_VariantToLong");
                extractVariantArg(falseVal, 2, "vb6_VariantToLong");
                lastExpr_ = "vb6_IIfLong(" + cond + ", " + trueVal + ", " + falseVal + ")";
            }
            return;
        }
    }

    // P18-D: VarPtr special handling - returns address of variable
    if (node.callee && node.callee->kind == ASTNodeKind::IdentifierExpr && node.positional.size() == 1) {
        auto& ident = static_cast<IdentifierExpr&>(*node.callee);
        std::string vpLower = ident.name;
        std::transform(vpLower.begin(), vpLower.end(), vpLower.begin(), ::tolower);
        if (vpLower == "varptr") {
            emitExpr(*node.positional[0]);
            /* Fix 082: VarPtr must return intptr_t, not int32_t.
               On x64, (int32_t)(intptr_t) truncates 8-byte pointers.
               Use (intptr_t) instead — safe on both x86 (4 bytes) and x64 (8 bytes). */
            /* Fix 084aa: VarPtr(函数调用) — 调用结果不是左值, &(call) → C2102.
               用复合字面量 &(void*){call} 包装 (VarPtr(.Glob(0)) 中的 .Glob 是 COM 属性
               调用 vb6_ComCall(...)); 变量/字段/数组元素保持原 &(expr). */
            bool vpIsCall = false;
            if (!lastExpr_.empty()
                && (std::isalpha(static_cast<unsigned char>(lastExpr_[0])) || lastExpr_[0] == '_')) {
                size_t vpParen = lastExpr_.find('(');
                if (vpParen != std::string::npos) {
                    vpIsCall = true;
                    for (size_t j = 0; j < vpParen; j++) {
                        char ch = lastExpr_[j];
                        if (!std::isalnum(static_cast<unsigned char>(ch)) && ch != '_') {
                            vpIsCall = false;
                            break;
                        }
                    }
                }
            }
            // Fix 092s: VB6_SA_AT / VB6_SA_ND_ATn 展开后是**左值** (SafeArray 元素),
            // 不能按"调用结果非左值"走复合字面量 — VarPtr(uOutput.Buffer(0)) 曾生成
            // (intptr_t)&(void*){VB6_SA_AT(...)} → C2440 (cTlsSocket 4093).
            // 注意 vb6_VariantArrayGet 按值返回, 不在白名单内.
            bool vpIsLvalue092s = lastExpr_.compare(0, 10, "VB6_SA_AT(") == 0
                || lastExpr_.compare(0, 12, "VB6_SA_ND_AT") == 0;
            // Fix 159-C: VarPtr(Variant 数组元素) — vb6_VariantArrayGet 按值返回
            // vb6_VARIANT 结构体, 复合字面量 &(void*){...} 把它塞进 void* 槽 →
            // C2440 "无法从 vb6_VARIANT 转换为 void *" (VTableHandle.c 181/187/267
            // ObjPtr(VarArgList(i)), VBFlexGrid.c 18504~18645 VarPtr(Value(i))).
            // 158f 的 vb6_VariantArrayElemPtr(vb6_VARIANT*, int32_t) 正是"取元素地址",
            // 与 vb6_VariantArrayGet 同参数形态 → 直接改名, 得到真左值地址.
            if (isVariantSlotElemRead(lastExpr_)) {
                lastExpr_ = "(intptr_t)" + variantSlotElemAddrRead(lastExpr_);
                return;
            }
            if (vpIsCall && !vpIsLvalue092s) {
                lastExpr_ = "(intptr_t)&(void*){" + lastExpr_ + "}";
            } else if (!lastExpr_.empty() && (std::isdigit(static_cast<unsigned char>(lastExpr_[0])) || lastExpr_[0] == '(')) {
                if (lastExpr_.size() > 1 && lastExpr_[0] == '(' && lastExpr_[1] == '*') {
                    // Fix 090h: (*name) / (*name).field / (*name)[idx] — ByRef 参数解引用是左值
                    // (如 VarPtr(File) → (*File), File As ByRef Variant; VarPtr(uFile.BufferArray)
                    // → (*uFile).BufferArray). 此前落入下方复合字面量分支, 生成
                    // (intptr_t)&(int32_t){(*File)} → C2440 (vb6_VARIANT → int32_t 初始化失败).
                    lastExpr_ = "(intptr_t)&(" + lastExpr_ + ")";
                } else {
                    // Fix 086: VarPtr(常量) — 常量被内联为字面量 (或括号表达式), 不可取址.
                    // 用 int32_t 复合字面量承载 (VB6 语义: 取常量临时副本地址).
                    lastExpr_ = "(intptr_t)&(int32_t){" + lastExpr_ + "}";
                }
            } else {
                lastExpr_ = "(intptr_t)&(" + lastExpr_ + ")";
            }
            return;
        }
    }

    // Fix 092t: TypeName(Me) — Me 在 C 侧是类指针 (vb6_cls_pvSubClass*), 而
    // vb6_TypeName 形参是 vb6_VARIANT → C2440 (pvSubClass 543/544,
    // clsSubClass.cls 497/498: App.LogEvent TypeName(Me) & ...).
    // 类名在编译期已知 (vbp 的 Class= 名 == moduleName_), 直接折叠为字符串字面量 —
    // 这也是 VB6 语义上唯一正确的形态: TypeName(对象实例) 返回其实类的名字,
    // 经 Variant 包装对象只会得到 "Object".
    if (node.callee && node.callee->kind == ASTNodeKind::IdentifierExpr
        && node.positional.size() == 1 && node.positional[0]
        && node.positional[0]->kind == ASTNodeKind::MeExpr && isClassModule_) {
        auto& tn092t = static_cast<IdentifierExpr&>(*node.callee);
        std::string tnLower092t = tn092t.name;
        std::transform(tnLower092t.begin(), tnLower092t.end(), tnLower092t.begin(), ::tolower);
        if (tnLower092t == "typename") {
            lastExpr_ = "vb6_BSTR_FromStr(L\"" + moduleName_ + "\")";
            return;
        }
    }

    // 账 #305 §B189 第二档 (第八十八刀): TypeName(<一枚 typed 工程类变量>) —— VB6 答的是
    // **声明档那个类的模块名**，而这一路改前把实参装进槽再去问 vb6_TypeName: 槽里是裸实例
    // 指针 (不是 §B189 第一档认得出的那枚包装器)，于是交回 "Object" (两台实测:
    // .build/k93_t 的 A 臂 42060)。真语料 9949 份里 65 处拿 TypeName 比工程内类名，56 处比的
    // 就是自己那个类 ⇒ 承面在行为，不在外观。
    // 闸门三条，都在**发码之前**问完 (发射实参只在决定折之后才做，免得拒折时留下发射副作用):
    //   · 实参必须是一枚裸标识符 —— 调用结果/成员读那一族的值不是"一枚变量"，声明档说了不算;
    //   · 名字被谁 Implements 过 ⇒ 这是接口视图，VB6 答实体的真身类名，声明档不是答案 ⇒ 不折
    //     (真语料那 9 处 CEmployeeCol 那一族就是这一形);
    //   · 控件那一族留在闸门外 (与 boxIntoLateSlot 同一条负条件): 窗体上的子控件在 C 侧的值是
    //     HWND，按声明档折会给句柄报出个类名 —— tests/setobj 钉的 TypeName(o) = "TextBox"
    //     那一档归宿主那一问 (Fix 112) 管。
    // NULL 那一支与第八十七刀同一条口径: 对象变量不指向对象 = "Nothing"。
    // 第四道闸 (§B147 那一条口径用在这里): 实参是**整枚槽** ⇒ 这一问是运行期的, 声明档不是答案。
    // 逼出这条闸的是 A/B 上唯一一处真语料移动 (尺 .build/k96_ab.py, 读数 .build/k96_ab.txt):
    // VBFlexGrid.ctl 的 `Public Property Let ColImageList(ByVal Index As Long, ByVal Value As Variant)`
    // 里 `TypeName(Value) = "ImageList"` 被折成了常量 —— 因为"这枚标识符是哪一枚类"按名查的那张表
    // (knownClassVars_)把**另一个过程**里那枚 `Property Let FlexDataSource(ByVal Value As
    // IVBFlexDataSource)` 的声明漏了过来 (两台产物对上: 改前那一行发的是 `vb6_TypeName(vb6_VariantFromValue(
    // vb6_VariantFromValue(Value)))`, 内层那一枚就是"Value 是槽"的既有答案)。折了就是把这个分支
    // 变成恒假 ⇒ 静默行为改动, 不是外观改动。
    if (node.callee && node.callee->kind == ASTNodeKind::IdentifierExpr
        && node.positional.size() == 1 && node.positional[0]
        && node.positional[0]->kind == ASTNodeKind::IdentifierExpr) {
        auto& tnId88 = static_cast<IdentifierExpr&>(*node.callee);
        std::string tnLow88 = tnId88.name;
        std::transform(tnLow88.begin(), tnLow88.end(), tnLow88.begin(), ::tolower);
        if (tnLow88 == "typename") {
            std::string tnCls88 = inferClassTypeOfExpr(*node.positional[0]);
            std::string tnSpell88 = tnCls88.empty() ? std::string() : projectClassNameOf(tnCls88);
            if (!tnSpell88.empty() && !classUsedAsInterface(tnCls88)
                && !variantOperandAST(node.positional[0].get())) {
                std::string tnKeep88 = lastExpr_;
                emitExpr(*node.positional[0]);
                std::string tnArg88 = lastExpr_;
                lastExpr_ = tnKeep88;
                if (tnArg88.find("vb6_hwnd_") == std::string::npos
                    && tnArg88.find("vb6_form_ctrl_") == std::string::npos) {
                    lastExpr_ = "(((" + tnArg88 + ") != NULL) ? vb6_BSTR_FromStr(L\""
                              + tnSpell88 + "\") : vb6_BSTR_FromStr(L\"Nothing\"))";
                    return;
                }
            }
        }
    }

    // P18-E: Choose special handling - nested ternary chain
    if (node.callee && !node.positional.empty()) {
        std::string csLower;
        if (node.callee->kind == ASTNodeKind::IdentifierExpr) {
            csLower = static_cast<IdentifierExpr&>(*node.callee).name;
        } else if (node.callee->kind == ASTNodeKind::MemberAccessExpr) {
            // Fix 137a: VBA.Choose(...) / VBA.Switch(...) — 源码以 VBA. 成员形态调用
            // (VBFlexGrid.ctl 3824, MainForm.frm 567/571, UserEditingForm.frm 242).
            // 需先约简 MemberAccessExpr 得到全局函数名再按 builtin 折叠, 否则 fallback
            // 生成裸 vb6_Choose(...) 直调 → 无签名/类型信息 → LNK2019 (VBFLXGRD).
            auto& maCs = static_cast<MemberAccessExpr&>(*node.callee);
            if (maCs.object && maCs.object->kind == ASTNodeKind::IdentifierExpr) {
                std::string objnm = Symbol::toLower(static_cast<IdentifierExpr&>(*maCs.object).name);
                if (objnm == "vba" || objnm == "vba5") {
                    csLower = maCs.memberName;
                }
            }
        }
        std::transform(csLower.begin(), csLower.end(), csLower.begin(), ::tolower);
        if (csLower == "choose" && node.positional.size() >= 2) {
            emitExpr(*node.positional[0]);
            std::string idx = std::move(lastExpr_);
            std::string result = "NULL";
            for (int i = (int)node.positional.size() - 1; i >= 1; i--) {
                emitExpr(*node.positional[i]);
                result = "((" + idx + ")==" + std::to_string(i) + " ? (" + lastExpr_ + ") : (" + result + "))";
            }
            lastExpr_ = result;
            return;
        }
        if (csLower == "switch" && node.positional.size() >= 2 && node.positional.size() % 2 == 0) {
            std::string result = "NULL";
            for (int i = (int)node.positional.size() - 2; i >= 0; i -= 2) {
                emitExpr(*node.positional[i]);
                std::string cond = "((" + lastExpr_ + ")!=0)";
                emitExpr(*node.positional[i + 1]);
                result = "(" + cond + " ? (" + lastExpr_ + ") : (" + result + "))";
            }
            lastExpr_ = result;
            return;
        }
    }

        // P21-B: Array(arglist) special handling - creates a Variant SafeArray
    // Array(1, "hello", 3.14) → vb6_ArrayCreate(3) + vb6_ArraySetLong/SetBSTR/SetDouble
    if (node.callee && node.callee->kind == ASTNodeKind::IdentifierExpr) {
        auto& arrIdent = static_cast<IdentifierExpr&>(*node.callee);
        std::string arrLower = arrIdent.name;
        std::transform(arrLower.begin(), arrLower.end(), arrLower.begin(), ::tolower);
        if (arrLower == "array" && node.positional.size() >= 0) {
            int count = (int)node.positional.size();
            std::string arrVar = "_arr_" + std::to_string(tempCounter_++);
            if (count > 0) {
                c_.emitLine("vb6_SafeArray1D* " + arrVar + " = vb6_ArrayCreate(" + std::to_string(count) + ");");
                for (int i = 0; i < count; i++) {
                    emitExpr(*node.positional[i]);
                    std::string argExpr = std::move(lastExpr_);
                    // Detect argument type for proper setter selection
                    bool isLongArg = false;
                    bool isDoubleArg = false;
                    if (node.positional[i]->kind == ASTNodeKind::LiteralExpr) {
                        auto& lit = static_cast<LiteralExpr&>(*node.positional[i]);
                        if (lit.literalKind == LiteralKind::Integer) isLongArg = true;
                        else if (lit.literalKind == LiteralKind::Double) isDoubleArg = true;
                    }
                    // Fix 038b-3: 检测 Variant 实参 — Array() 元素通过 vb6_ArraySetLong/
                    // SetBSTR/SetDouble 设置, 这些函数期望具体类型. 当实参是 Variant
                    // (如 vb6_VariantArrayGet, Variant 变量等) 时, 需要先提取具体值.
                    // 仅使用 cExprIsVariant (C 字符串级) 和 knownVariantVars_ 检测.
                    bool argIsVariant = cExprIsVariant(argExpr);
                    if (!argIsVariant && node.positional[i]->kind == ASTNodeKind::IdentifierExpr) {
                        auto& idArg = static_cast<IdentifierExpr&>(*node.positional[i]);
                        std::string argLower = idArg.name;
                        std::transform(argLower.begin(), argLower.end(), argLower.begin(), ::tolower);
                        if (knownVariantVars_.count(argLower)) argIsVariant = true;
                    }
                    // 也检测 BSTR 变量
                    bool argIsBstr = false;
                    if (node.positional[i]->kind == ASTNodeKind::IdentifierExpr) {
                        auto& idArg = static_cast<IdentifierExpr&>(*node.positional[i]);
                        std::string argLower = idArg.name;
                        std::transform(argLower.begin(), argLower.end(), argLower.begin(), ::tolower);
                        if (knownBstrVars_.count(argLower)) argIsBstr = true;
                    }
                    if (isLongArg) {
                        c_.emitLine("vb6_ArraySetLong(" + arrVar + ", " + std::to_string(i) + ", " + argExpr + ");");
                    } else if (isDoubleArg) {
                        c_.emitLine("vb6_ArraySetDouble(" + arrVar + ", " + std::to_string(i) + ", " + argExpr + ");");
                    } else if (argIsVariant) {
                        // Variant 实参: 无法确定具体类型, 统一用 VariantToLong 转换
                        // (Array() 的元素通常是数值索引). 若为 BSTR Variant 则需
                        // VariantToString, 但 Array() 上下文中 Long 是最常见的默认.
                        // 更安全: 使用 vb6_VariantFromValue 让 _Generic 选择, 然后
                        // 直接用 vb6_ArraySetVariant (如果存在) 或回退到 Long.
                        bool looksLikeBSTR = (argExpr.find("vb6_BSTR") != std::string::npos ||
                                              argExpr.find("L\"") != std::string::npos);
                        // Fix 081i: 也用 inferExprType 检测 BSTR 类型
                        if (!looksLikeBSTR) {
                            Vb6Type argVt = inferExprType(*node.positional[i]);
                            if (argVt == Vb6Type::String) looksLikeBSTR = true;
                        }
                        if (looksLikeBSTR) {
                            c_.emitLine("vb6_ArraySetBSTR(" + arrVar + ", " + std::to_string(i) + ", vb6_VariantToString(" + argExpr + "));");
                        } else {
                            c_.emitLine("vb6_ArraySetLong(" + arrVar + ", " + std::to_string(i) + ", vb6_VariantToLong(" + argExpr + "));");
                        }
                    } else if (argIsBstr) {
                        c_.emitLine("vb6_ArraySetBSTR(" + arrVar + ", " + std::to_string(i) + ", " + argExpr + ");");
                    } else {
                        bool looksLikeBSTR = (argExpr.find("vb6_BSTR") != std::string::npos ||
                                              argExpr.find("L\"") != std::string::npos);
                        // Fix 081i: 也用 inferExprType 检测, 避免内置函数返回 BSTR
                        // (如 vb6_ErrSource/vb6_ErrDescription) 被误判为 Long
                        if (!looksLikeBSTR) {
                            Vb6Type argVt = inferExprType(*node.positional[i]);
                            if (argVt == Vb6Type::String) looksLikeBSTR = true;
                        }
                        if (looksLikeBSTR) {
                            c_.emitLine("vb6_ArraySetBSTR(" + arrVar + ", " + std::to_string(i) + ", " + argExpr + ");");
                        } else {
                            // Default: treat as Long (covers variable references etc.)
                            c_.emitLine("vb6_ArraySetLong(" + arrVar + ", " + std::to_string(i) + ", " + argExpr + ");");
                        }
                    }
                }
            } else {
                c_.emitLine("vb6_SafeArray1D* " + arrVar + " = vb6_ArrayCreate(0);");
            }
            lastExpr_ = arrVar;
            return;
        }
    }

    // 函数调用路径 (原有逻辑)
    // P20-36: IsMissing(arg) -> (!_has_arg) for Optional params
    // Must intercept BEFORE callee mapping to vb6_IsMissing
    if (node.callee && node.callee->kind == ASTNodeKind::IdentifierExpr && node.named.empty()) {
        auto& ismIdent = static_cast<IdentifierExpr&>(*node.callee);
        std::string ismLower = ismIdent.name;
        std::transform(ismLower.begin(), ismLower.end(), ismLower.begin(), ::tolower);
        if (ismLower == "ismissing" && node.positional.size() == 1) {
            auto& arg = node.positional[0];
            if (arg->kind == ASTNodeKind::IdentifierExpr) {
                std::string argName = static_cast<IdentifierExpr&>(*arg).name;
                std::string argLower = argName;
                std::transform(argLower.begin(), argLower.end(), argLower.begin(), ::tolower);
                if (currentProc_) {
                    for (auto& p : currentProc_->params) {
                        std::string pLower = p.name;
                        std::transform(pLower.begin(), pLower.end(), pLower.begin(), ::tolower);
                        if (pLower == argLower && p.isOptional && !p.isParamArray) {
                            // 账 #257 第二十层: 必须交 **VB6 布尔 (-1/0)**, 不能交 C 布尔 (0/1)。
                            // VB6 的 `Not` 在 cgen_expr.cpp 发成 `(~(int32_t)(<操作数>))` —— 按位
                            // 取反只有在操作数是 -1/0 时才等于逻辑非: ~0=-1 (True), ~(-1)=0 (False)。
                            // 交 0/1 时 ~0=-1 与 ~1=-2 **都非零** ⇒ `Not IsMissing(x)` 恒真。
                            // 实测 ComCtlsDemo TvwNodes.Add: `If Not IsMissing(Relative) Then` 在
                            // Relative 省略 (`TreeView1.Nodes.Add , , , "1", 1`) 时也成立, 于是走进
                            // `Select Case VarType(Relative)` 对 vbEmpty 落 Case Else ⇒ Err.Raise 13。
                            // 直接写 `If IsMissing(x)` 那一形不受影响 (0 与 -1 都按假/真成立)。
                            lastExpr_ = "(-(!_has_" + cIdent(p.name) + "))";
                            return;
                        }
                        // Fix 091k: ParamArray 整体 — RTL vb6_IsMissing(SAFEARRAY*)
                        // 语义为"未传任何实参" (psa == NULL).
                        if (pLower == argLower && p.isParamArray) {
                            lastExpr_ = "vb6_IsMissing(" + cIdent(p.name) + ")";
                            return;
                        }
                    }
                }
                // Not an Optional param - IsMissing returns False (0)
                lastExpr_ = "(0)";
                return;
            }
            // Fix 091k: 实参是复合表达式 (数组元素 / ParamArray 元素 / 函数结果).
            // VB6 IsMissing 仅对 Optional Variant 形参可能为 True, 其余恒 False;
            // 而 RTL vb6_IsMissing 形参是 SAFEARRAY* (ParamArray 专用), 直接透传
            // Variant 值会 C2440 (cToolsArray: IsMissing(Vars(i)), IsMissing(OutVars(i))).
            lastExpr_ = "(0)";
            return;
        }
    }

    // Fix 038: Len(udt) → sizeof(udt) — VB6 Len on UDT returns size in bytes.
    // vb6_Len is declared as int32_t vb6_Len(BSTR), so passing a UDT causes C2440.
    // Must intercept BEFORE callee mapping to vb6_Len.
    if (node.callee && node.callee->kind == ASTNodeKind::IdentifierExpr && node.named.empty()) {
        auto& lenIdent = static_cast<IdentifierExpr&>(*node.callee);
        std::string lenLower = lenIdent.name;
        std::transform(lenLower.begin(), lenLower.end(), lenLower.begin(), ::tolower);
        if (lenLower == "len" && node.positional.size() == 1) {
            // 账 #118 ①: **定长串的 Len 恒为声明长度 N**, 与当前装着几个字符无关
            // (`Dim a As String * 5: a = "ab"` 在真 VB6 里 Len(a) = 5)。
            // 此前落通用 vb6_Len 路 ⇒ SysStringLen 给的是**值的长度** (2)。
            // 这一条必须排在下面所有 sizeof 兜底**之前**: 模块级定长串不进
            // knownBstrVars_ (那张表每过程入口 clear), 会被"兜底桶"当数值发成
            // sizeof(指针) —— x64 读 8 (LEN-DECL 实测), 而定长串要给的就是 N。
            {
                const std::string fsLen118 = fixedStrLenOfExpr(*node.positional[0]);
                if (!fsLen118.empty()) {
                    lastExpr_ = "((int32_t)(" + fsLen118 + "))";
                    return;
                }
            }
            // VB6 口径 (真 VB6 实测): Len 的实参不是 String 就必须**能落到变量上** ——
            // 非 String 的 Len 取的是"存储宽度", 字面量没有存储宽度可言, 真 VB6 直接
            // 编译错误 "Variable required - can't assign to this expression"
            // (Len(1234) / Len(234.5) 均报错; Len("1234")=4 合法)。
            // 此前 C3 静默把它字符串化再数字符 (Len(1234)=4 纯属巧合, Len(234.5)=5)。
            if (node.positional[0]->kind == ASTNodeKind::LiteralExpr) {
                auto& lit44 = static_cast<LiteralExpr&>(*node.positional[0]);
                if (lit44.literalKind != LiteralKind::String) {
                    // 复用既有的 CodeGenUnsupportedFeature (4001): 不新增诊断号位。
                    diag_.error(DiagnosticID::CodeGenUnsupportedFeature, node.positional[0]->loc,
                        "Len argument must be a variable or a String expression; "
                        "a numeric/date literal has no storage width");
                    lastExpr_ = "(0)";
                    return;
                }
            }
            std::string udtCType = inferUdtTypeOfExpr(*node.positional[0]);
            if (!udtCType.empty()) {
                emitExpr(*node.positional[0]);
                lastExpr_ = "((int32_t)sizeof(" + lastExpr_ + "))";
                return;
            }
            // Task #44 探针发现: Len(数值) = **该类型的存储宽度** (Long→4, Integer→2, ...),
            // 不是"十进制位数"。此前落进 vb6_Len(BSTR) 通用路径 → 实参先被 vb6_CStrLong
            // 转成字符串再数字符, `Len(h)` (h=33489972) 打出 8, VB6 语义是 4。
            // 与 Fix 125 (LenB 数值→sizeof) / Fix 038 (Len UDT→sizeof) 同族, 在实参
            // 字符串化之前拦截。Variant/字符串实参不在此列 (走原路径)。
            // VB6 口径 (用户拿真 VB6 实测, 2026-09-27): 非 String 的 Len **必须吃变量**
            // —— 它要的是"存储宽度", 数值字面量/表达式在真 VB6 里直接编译错误
            // "变量要求 - 不能赋值到该表达式" (Len(1234) 报错; Len("1234")=4 合法,
            // 因为 String 走的是字符数那条语义)。
            // 所以下面两条只对**变量形态** (标识符 / 成员访问) 生效。
            // 复合表达式 (Len(a+1) / Len(Foo())) 落回原 vb6_Len 路径 —— 与本次改动前
            // 逐字相同, 零回归。真 VB6 对这些也报错, 但"函数调用返回数值"与"数组元素"
            // 在 AST 层不易区分 (Len(mCustomColors(0)) 是合法的), 误报会打断编译,
            // 故只把**字面量**这一 unambiguous 的形状做成硬报错 (见上, 3043)。
            const bool lenArgIsVar44 =
                node.positional[0]->kind == ASTNodeKind::IdentifierExpr
                || node.positional[0]->kind == ASTNodeKind::MemberAccessExpr;
            Vb6Type tLen44 = inferExprType(*node.positional[0]);
            // Task #44 (SSTabEx "..."): 数组元素下标访问 (Len(mCustomColors(0)))
            // 同样是"变量形态" —— VB6 语义 Len(数组元素) = **元素类型的存储宽度**
            // (Long → 4)。此前 IndexOrCallExpr 被排除在 lenArgIsVar44 之外, 落回
            // vb6_Len 通用路径: 实参先被 vb6_CStrLong 字符串化再**数字符** →
            // 元素值 0 打出 "0" → Len=1。真实场景: cDlg.ShowColor 的
            // lCustomColorSize = Len(mCustomColors(0))*16 → GlobalAlloc(16),
            // 而 ChooseColorA 按 16 个 Long 写 64 字节 → 堆越界 48 字节 →
            // 0xC000041D (SSTabEx "..." 按钮一点即退, 实测)。
            // 判"下标访问的是数组"而非函数调用, 双查口径同 For Each 源判定
            // (cgen_control.cpp): knownArrays_ 命中, 或符号表 Variable+isArray。
            // 普通函数调用 Len(Foo()) 两者皆不满足 → 维持落回原路径的既有口径。
            Vb6Type tArrElem44 = Vb6Type::Unknown;
            if (node.positional[0]->kind == ASTNodeKind::IndexOrCallExpr) {
                auto& call44 = static_cast<IndexOrCallExpr&>(*node.positional[0]);
                if (call44.callee && call44.callee->kind == ASTNodeKind::IdentifierExpr) {
                    std::string arrName44 = Symbol::toLower(
                        static_cast<IdentifierExpr&>(*call44.callee).name);
                    if (knownArrays_.count(arrName44)) {
                        tArrElem44 = arrayElemTypes_[arrName44];
                    } else {
                        Symbol* sym44 = symTab_.lookup(arrName44);
                        if (!sym44) sym44 = symTab_.lookupModule(arrName44);
                        if (sym44 && sym44->kind == SymbolKind::Variable && sym44->isArray) {
                            // 数组符号的 type 带 Array 位 (mapType 依它生成
                            // SafeArray), 元素类型 = 去 Array 位的基础类型
                            // (口径同 cgen_base_type.cpp mapType)。
                            tArrElem44 = static_cast<Vb6Type>(
                                static_cast<uint16_t>(sym44->type)
                                & ~static_cast<uint16_t>(Vb6Type::Array));
                        }
                    }
                }
            }
            const bool lenArrElemIsNumeric44 =
                tArrElem44 == Vb6Type::Byte     || tArrElem44 == Vb6Type::Boolean ||
                tArrElem44 == Vb6Type::Integer  || tArrElem44 == Vb6Type::Long ||
                tArrElem44 == Vb6Type::LongPtr  || tArrElem44 == Vb6Type::LongLong ||
                tArrElem44 == Vb6Type::Single   || tArrElem44 == Vb6Type::Double ||
                tArrElem44 == Vb6Type::Currency || tArrElem44 == Vb6Type::Date;
            if (lenArrElemIsNumeric44) {
                emitExpr(*node.positional[0]);
                // 数组元素表达式 (VB6_SA_AT 展开成 ((T*)data)[i-lBound] 或
                // Byte 数组的指针下标) 已是左值形态, sizeof 取其 C 存储宽度 ——
                // 与标量分支同口径; sizeof 不求值, 下标副作用不被重复执行。
                lastExpr_ = "((int32_t)sizeof(" + lastExpr_ + "))";
                return;
            }
            // 只拦**标量数值**; String / Variant / Object / 未知一律走原 vb6_Len 路径。
            const bool lenIsNumeric44 =
                lenArgIsVar44 && (
                tLen44 == Vb6Type::Byte     || tLen44 == Vb6Type::Boolean ||
                tLen44 == Vb6Type::Integer  || tLen44 == Vb6Type::Long ||
                tLen44 == Vb6Type::LongPtr  || tLen44 == Vb6Type::LongLong ||
                tLen44 == Vb6Type::Single   || tLen44 == Vb6Type::Double ||
                tLen44 == Vb6Type::Currency || tLen44 == Vb6Type::Date);
            if (lenIsNumeric44) {
                emitExpr(*node.positional[0]);
                // 用 sizeof(实参) 而**不是**按 Vb6Type 查字节数表:
                //   knownLongVars_ 把 Integer/Long/Byte 并进同一个登记桶 (C 型分别是
                //   int16_t/int32_t/uint8_t 但登记串同), inferExprType 对 `Dim i As
                //   Integer` 也返回 Long → 查表会把 Len(i) 判成 4 (VB6 是 2)。
                //   sizeof 直接取 C 侧真实存储宽度, 与 VB6 "Len(数值)=存储字节数" 同义。
                // 顺带修掉既有崩溃: `Dim b As Byte` 不落任何登记桶 → 此前漏拦 →
                //   vb6_Len(uint8_t) 把 200 当 BSTR 指针解引用 → 0xC0000005 (LEN-BYTE 实测)。
                // sizeof 不求值其操作数, 故 Len(Foo()) 不会重复调用 (VB6 会求值一次;
                // 数值实参求值无可观测副作用的场景占绝大多数, 这里取不崩优先)。
                lastExpr_ = "((int32_t)sizeof(" + lastExpr_ + "))";
                return;
            }
            // Byte / Currency / LongLong / 枚举 之类**不落任何登记桶**
            // (cgen_localdecl.cpp:160-186 只登记 BSTR / int32_t|int16_t|VBABOOL /
            // intptr_t / ComIface / vb6_VARIANT), inferExprType 对它们返回 Unknown,
            // 上面拦不住 → `Dim b As Byte: b = 200: Len(b)` 生成 vb6_Len(uint8_t),
            // 把 200 当 BSTR 指针解引用 → 0xC0000005 (LEN-BYTE 实测, 既有 bug)。
            // 对"确认不是字符串/Variant/对象"的裸标识符兜底 sizeof。
            // 2026-09-28 订正：上面那句"字符串侧 knownBstrVars_ 在声明处必登记"**只对局部成立** ——
            // 那张表在每过程入口 clear (cgen_decl_func.cpp:43 / cgen_decl_proc.cpp:51)，
            // 模块级 `Dim gS As String` 压根不进表 ⇒ 兜底桶把模块级 String 也吞了，
            // `Len(gS)` 发成 (int32_t)sizeof(gS) = 指针宽度（x64 读 8、x86 读 4，实测）。
            // 改判两把尺，都不动数值那一半的既有口径：
            //   · knownFixedStringLen_ —— 定长串在**模块级声明处也登记**(cgen_decl_var.cpp:199)，
            //     它的 C 型不是 BSTR，Len 要的就是存储宽度，继续走 sizeof；
            //   · tLen44 == String —— 符号表认得的字符串（模块级变量、vbCrLf 这类常量）走
            //     通用 vb6_Len 路，那才是字符数。
            if (node.positional[0]->kind == ASTNodeKind::IdentifierExpr) {
                std::string low44 = Symbol::toLower(
                    static_cast<IdentifierExpr&>(*node.positional[0]).name);
                if (!knownBstrVars_.count(low44) && !knownVariantVars_.count(low44)
                    && !knownObjectVars_.count(low44)
                    && !knownFixedStringLen_.count(low44)
                    && tLen44 != Vb6Type::String) {
                    emitExpr(*node.positional[0]);
                    lastExpr_ = "((int32_t)sizeof(" + lastExpr_ + "))";
                    return;
                }
            }
        }
    }
