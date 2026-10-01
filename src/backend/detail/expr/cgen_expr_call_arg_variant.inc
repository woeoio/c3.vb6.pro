// cgen_expr_call_arg_variant.inc - visit(IndexOrCallExpr&) 函数体片段：位置实参发射（二）—— ByVal Variant 包装与 Variant 提取（位于实参 for 循环内）
// 由 src/backend/expr/cgen_expr_call.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 2137~2453 行，逐行未改。

        // Fix 160w: Declare 形参 As Any (type==Unknown) 且调用处 ByVal 覆盖
        // (如 OleCreatePropertyFrame ..., ByVal Pages.pElems, ...) — 声明里 As Any
        // 没有 ByVal 关键字, 故 calleeParams[i].isByVal==false, 上方 ByVal Variant
        // 分支与 Fix 029 都不匹配 → 实参 (vb6_VariantFromComResult(...) 等 COM
        // 属性读取的 vb6_VARIANT 值) 会裸传 → C2172 "实参不是函数" (参 8).
        // VB6 语义: As Any ByVal 即按值传整型载荷; Variant 的整型载荷
        // (LongPtr/Long 指针值, 如 pElems 存 IFontDisp** 数组头) 用
        // vb6_VariantToLongPtr 提取再压回 void* 实参.
        if (!isByRef && i < calleeParams.size()
            && calleeParams[i].type == Vb6Type::Unknown
            && argHasByValOverride && cExprIsVariant(argVal)) {
            argVal = "(void*)(intptr_t)vb6_VariantToLongPtr(" + argVal + ")";
        }

        // 处理 callee 声明 "ByVal x As Variant" 而实参是标量/BSTR/SafeArray/class ptr 等情形.
        // _Generic 在编译期按实类型选择 ctor: 标量->VariantLong/Int/Double, BSTR->String,
        // SafeArray1D*->Array, void*/class ptr->Object, vb6_VARIANT->Identity(no-op).
        // 仅对 ByVal Variant 生效 (ByRef Variant 走上面复合字面量路径, 取地址需左值).
        // <vbeclipse>: 但 "p.type == Variant" 里有第二种槽 —— 工程类形参 (`ByVal h As
        // DemoShape`): 类名只留在 typeRefName, C 侧形参是 vb6_cls_X*, 装箱喂进去就是
        // C2440。这一种先按**槽**交付, 不参与下面的装箱。判据与 dll 入口 extern 共用
        // cParamClassPtrType 一处, 原型和调用点不再各说一套。
        if (!isByRef && i < calleeParams.size() && calleeParams[i].isByVal) {
            const std::string cp091t = cParamClassPtrType(calleeParams[i]);
            if (!cp091t.empty()) argVal = deliverToClassSlot(cp091t, argVal);
        }
        if (!isByRef && i < calleeParams.size() && calleeParams[i].isByVal
            && calleeParams[i].type == Vb6Type::Variant
            && cParamClassPtrType(calleeParams[i]).empty()
            // <vbeclipse> <VBFlexGridDemo> 同 arg_emit / Fix 086: 判据是 mapTypeRef
            // 出的 C 类型 (不是"名字是否非空")。"As Variant" 会得 vb6_VARIANT, 装箱
            // 照旧; "As IOleControl"/接口 等得 void*/vb6_ComIface_X*, 不装。
            && (calleeParams[i].typeRefName.empty()
                || cTypeForDeclaredTypeName(calleeParams[i].typeRefName) == "vb6_VARIANT")) {
            // Fix 084d: 实参若是类对象表达式 (如 me->Database, 类方法返回类对象),
            // 保持对象指针直传, 不要套 vb6_VariantFromValue — 符号表可能把类类型
            // 参数误记为 Variant, 若包装成 vb6_VARIANT 传给 vb6_cls_cXxx* 参数
            // 会触发 C2440 (如 cHttpServer.c LoadFromDatabase(..., me->Database)).
            bool argIsClassObj = false;
            if (i < node.positional.size()) {
                std::string acName091s = inferClassTypeOfExpr(*node.positional[i]);
                if (!acName091s.empty()) {
                    // Fix 091s: 只有"项目类"(SymbolKind::Class) 实参才跳过打包
                    // (084d 的类指针直传语义). Object/COM 类实参 (void*) 传给
                    // ByVal Variant 形参时仍需 vb6_VariantFromValue 包装 —
                    // Demo 857/859: Debug.Print ToolsJsonVba.ConvertToJson(Json)
                    // (Json As Object; ConvertToJson 第 1 参 ByVal JsonValue As
                    // Variant) 原样传 void* → C2440 (void* → vb6_VARIANT).
                    auto* acSym091s = symTab_.lookup(acName091s);
                    argIsClassObj = (acSym091s && acSym091s->kind == SymbolKind::Class);
                }
            }
            if (!argIsClassObj) {
                // Fix 198: 布尔实参按 VT_BOOL 装箱 (int16_t 走 _Generic 会得 VT_I2).
                argVal = boxToVariant(node.positional[i].get(), argVal);
            } else if (i < node.positional.size()
                       && node.positional[i]->kind == ASTNodeKind::IdentifierExpr) {
                // Fix 091s2: 实参是 Object/COM 变量 (void*) 而形参是 ByVal Variant
                // → 仍需打包 (inferClassTypeOfExpr 对 Object 变量可能返回内建
                // "Object" 类名而误判为项目类实参). Demo 857/859:
                // ToolsJsonVba.ConvertToJson(Json) (Json As Object).
                std::string on091s2 = Symbol::toLower(
                    static_cast<IdentifierExpr&>(*node.positional[i]).name);
                if (knownObjectVars_.count(on091s2) || knownTypedComVars_.count(on091s2)) {
                    // 同 arg_emit 那一处: 装箱只走 boxToVariant 这一权威 (见该处注释)
                    argVal = boxToVariant(node.positional[i].get(), argVal);
                }
            }
        }
        // Fix 029: 反向强制 — ByVal 具体类型参数 + 实参确定为 Variant: 自动调用提取函数.
        // 与 Fix 024 P2 互补: P2 处理 V(callee)=Variant,V(arg)=scalar; 此处处理
        // V(callee)=concrete,V(arg)=Variant. 覆盖 C2440 子类 to_int32/to_BSTR/to_SafeArray.
        // 严格判断: 用 isDefinitelyVariantExpr 避免 inferExprType 默认回退到 Variant
        // 造成对内置函数 (LenB 等) / UDT 字段访问 (uAddr.sin_addr) 的错误包装.
        if (!isByRef && i < calleeParams.size() && calleeParams[i].isByVal
            && calleeParams[i].type != Vb6Type::Variant) {
            // 处理 Array 标志位 (例如 Variant() 参数对应 Vb6Type::Variant|Array)
            bool paramIsArray = (static_cast<uint16_t>(calleeParams[i].type)
                                  & static_cast<uint16_t>(Vb6Type::Array)) != 0;
            Vb6Type paramBase = static_cast<Vb6Type>(
                static_cast<uint16_t>(calleeParams[i].type)
                & ~static_cast<uint16_t>(Vb6Type::Array));
            bool argIsVariantArr = false;
            bool argIsVariant = isDefinitelyVariantExpr(*node.positional[i], &argIsVariantArr);
            // Fix 038b-2: 字符串级 Variant 检测回退 — 补充 isDefinitelyVariantExpr
            // 无法识别的 C 级 Variant 表达式 (vb6_VariantArrayGet, vb6_VariantFromComResult 等)
            if (!argIsVariant && !argIsVariantArr) {
                argIsVariant = cExprIsVariant(argVal);
            }
            // Fix 113i: 形参是**具体类型** (String/数值/对象) 而实参是**裸 COM 结果**
            // (vb6_ComCall / vb6_ComGetProp / vb6_ComGetObjectProp 返回 void*,
            // 承载一个 COM VARIANT*) — 语义即 Variant, 必须判为 Variant 才能触发
            // 下方按 paramBase 的提取 (VariantToString / VariantToLong ...).
            // 否则 void* 被裸传给 BSTR 形参, GDI+ 把垃圾指针当字符串 → 全落空,
            // 渲染成一排"豆腐块"空框.
            //   Charts 2020 ucChartBar.DrawVertical X 轴月份标签:
            //     DrawText hGraphics, cAxisItem.Item(i), ... (第 2 参 ByVal As String)
            //     → 原生成 vb6_ucChartBar_DrawText((void*)me, hGraphics,
            //        vb6_ComCall(me->cAxisItem, L"Item", ..., 1), ...) 乱码.
            // 与 Fix 113c(运行时/内置函数参数分支) 同一检测; Fix 113 把
            // vb6_ComCall( 从 cExprIsVariant 前缀表回退后, 需要在两条实参发射
            // 路径上都补回该判断 —— 本条补的是"已知形参类型"这条路径.
            if (!argIsVariant && !argIsVariantArr) {
                bool argIsComRes113i =
                    (argVal.compare(0, 12, "vb6_ComCall(") == 0)
                    || (argVal.compare(0, 15, "vb6_ComGetProp(") == 0)
                    || (argVal.compare(0, 21, "vb6_ComGetObjectProp(") == 0);
                if (argIsComRes113i) argIsVariant = true;
            }
            if (argIsVariant || argIsVariantArr) {
                // Fix 090c1: 实参顶层是 COM 结果 (void* 承载 Variant) 且判为
                // Variant — 提取前先 vb6_VariantFromComResult (同 090c0).
                // cLogs MakeLogContent: Join(CacheDatas(i), Span) 的 CacheDatas(i)
                // = vb6_ComCall(...) → 原生成 VariantToSafeArray1D(ComCall) C2440.
                bool isComResultVal2 =
                    (argVal.find("vb6_ComCall(") == 0)
                    || (argVal.find("vb6_ComGetProp(") == 0)
                    || (argVal.find("vb6_ComGetObjectProp(") == 0);
                if (isComResultVal2) {
                    argVal = "vb6_VariantFromComResult(" + argVal + ")";
                }
                // Fix 091k: 实参是当前过程的 ParamArray 参数时, C 侧已是 SAFEARRAY*
                // (cgen 参数映射), 不能再按 Variant 数组提取 SafeArray1D* — 否则
                // UBound(ParamArray) 生成 vb6_PA_UBound(vb6_VariantToSafeArray1D(OutVars))
                // C2440 (cToolsArray DeArray).
                bool argIsPA091k = false;
                if (i < node.positional.size()
                    && node.positional[i]->kind == ASTNodeKind::IdentifierExpr && currentProc_) {
                    std::string paLower091k = Symbol::toLower(
                        static_cast<IdentifierExpr&>(*node.positional[i]).name);
                    for (auto& p091k : currentProc_->params) {
                        if (Symbol::toLower(p091k.name) == paLower091k) {
                            argIsPA091k = p091k.isParamArray;
                            break;
                        }
                    }
                }
                // Fix 092g: 实参是字面量数组 (Array(...) → 匿名 _arr_N) — C 侧
                // 本身已是 vb6_SafeArray1D* (由 vb6_ArraySetBSTR 逐个填充),
                // 不能再按 Variant 数组提取 (提取函数形参是 vb6_VARIANT) → C2440:
                //   cPLI 56: Cmd = Join(Array(a, b, c), " ")
                //     → vb6_Join(vb6_VariantToSafeArray1D(_arr_0), ...).
                // <vbeclipse>: 同一条边界扩到"内置函数调用已经返回载体"那一形
                // (Split/Filter/ArrayCreate...), 判定收进 cExprIsSafeArrayCarrier:
                //   Join(Split(s, ","), "|") → C2440 (实测三条全中, 含 UBound(Split(...)))
                bool argIsCarrier092g = cExprIsSafeArrayCarrier(argVal);
                if (paramIsArray && !argIsPA091k && !argIsCarrier092g
                    && (paramBase == Vb6Type::Variant || paramBase == Vb6Type::Byte
                    || paramBase == Vb6Type::String || paramBase == Vb6Type::Long)) {
                    // 数组参数: 从 Variant 提取 SafeArray1D*
                    argVal = "vb6_VariantToSafeArray1D(" + argVal + ")";
                } else if (paramBase == Vb6Type::Long || paramBase == Vb6Type::Integer
                           || paramBase == Vb6Type::Byte || paramBase == Vb6Type::Boolean) {
                    // Fix 092q: Long/Integer/Byte/Boolean ByVal 形参接收 Variant →
                    // vb6_VariantToLong (int32 提取).
                    argVal = "vb6_VariantToLong(" + argVal + ")";
                } else if (paramBase == Vb6Type::LongPtr || paramBase == Vb6Type::ULong) {
                    // Fix 092r (Charts 2020 x64): LongPtr/ULong ByVal 形参接收 Variant
                    // 实参 (如 GetTextSize 里 Dim hFont 无类型 = Variant, 存 GDI+ Font
                    // 对象指针), 必须用 vb6_VariantToLongPtr (intptr_t, 完整64位指针) —
                    // 原生成 vb6_VariantToLong 把 GDI+ 对象指针截断成 32 位 → 绘图时
                    // GdipMeasureString 内解引用坏指针 → 0xC0000005 (x86 下 LongPtr=32
                    // 位无截断, 故仅 x64 崩溃).
                    argVal = "vb6_VariantToLongPtr(" + argVal + ")";
                } else if (paramBase == Vb6Type::Double || paramBase == Vb6Type::Single
                           || paramBase == Vb6Type::Currency || paramBase == Vb6Type::Date) {
                    // Fix 090r: Date 形参 (OLE date = double) 与 Currency 同取
                    // vb6_VariantToDouble — 缺 Date 导致实参是 Variant 数组元素
                    // (vb6_VariantArrayGet) 时裸传 → C2440 VARIANT→double
                    // (cZipArchive pvVfsOpen: pvToFileTime(me, arr(i)) 形参 ByVal Date)
                    argVal = "vb6_VariantToDouble(" + argVal + ")";
                } else if (paramBase == Vb6Type::String) {
                    // Fix 049b/092a: 仅当 C 表达式**顶层**已是 BSTR 时跳过提取.
                    // 旧实现用子串 find("vb6_BSTR") — 对"内部实参含
                    // vb6_BSTR_FromStr(...)"的 Variant 表达式天然命中 → 整体跳过
                    // 提取 → C2440 (vb6_VARIANT→BSTR):
                    //   cAesCBC 25 StrConv(LoadResData("AES.CBC","JSCRIPT"), 64, 0)
                    //   → argVal = vb6_LoadResData(vb6_BSTR_FromStr(..), ..) 含子串.
                    // 改为与 5080 同源的**顶层前缀**判断 (只看表达式开头).
                    // 注: 进入本分支的前提是实参已被判为 Variant
                    // (isDefinitelyVariantExpr / cExprIsVariant), 顶层判断足以避免
                    // 对 MsgBox 专用逻辑/4622 分支已转换结果的二次包装.
                    static const char* bstrTopPrefixes092a[] = {
                        "vb6_BSTR_",
                        "VB6_SA_AT(BSTR,",
                        "vb6_VariantToString(",
                        "vb6_BSTR_FromStr("};
                    bool topIsBstr092a = false;
                    for (auto* bp092a : bstrTopPrefixes092a) {
                        if (argVal.compare(0, strlen(bp092a), bp092a) == 0) {
                            topIsBstr092a = true;
                            break;
                        }
                    }
                    if (!topIsBstr092a) {
                        argVal = "vb6_VariantToString(" + argVal + ")";
                    }
                } else if (paramBase == Vb6Type::Object) {
                    // 右值兼容: 避免对函数返回值取址
                    argVal = "vb6_VariantToObjectVal(" + argVal + ")";
                }
            }
        }
        // Fix 086: ByVal Variant形参兜底 — 实参为具体标量/BSTR/对象指针时,
        // 用 _Generic vb6_VariantFromValue 包装. _Generic 按实参C类型自动选择
        // 构造函数 (int→VariantLong, BSTR→VariantString, SafeArray*→VariantArray,
        // 类指针→VariantObject, 已是VARIANT→恒等), 从根源消除 int/BSTR/double→
        // vb6_VARIANT 方向的 C2440.
        if (!isByRef && i < calleeParams.size() && calleeParams[i].isByVal
            && (calleeParams[i].type == Vb6Type::Variant
                || calleeParams[i].type == Vb6Type::Empty)
            // <vbeclipse> <VBFlexGridDemo> 同族 guard (arg_emit / 091s): 判据
            // 看 mapTypeRef 那份 C 类型是不是 vb6_VARIANT, 不是才不装。
            && (calleeParams[i].typeRefName.empty()
                || cTypeForDeclaredTypeName(calleeParams[i].typeRefName) == "vb6_VARIANT")) {
            bool alreadyVariant = cExprIsVariant(argVal);
            if (!alreadyVariant && node.positional[i]->kind == ASTNodeKind::IdentifierExpr) {
                auto& idArg86 = static_cast<IdentifierExpr&>(*node.positional[i]);
                if (knownVariantVars_.count(Symbol::toLower(idArg86.name))) alreadyVariant = true;
            }
            // &(x) 形态是 ByRef 风格临时/取址, 不适合按值包装
            if (!alreadyVariant && argVal.compare(0, 2, "&(") != 0) {
                Vb6Type argT86 = inferExprType(*node.positional[i]);
                bool argIsArr86 = (static_cast<uint16_t>(argT86) & static_cast<uint16_t>(Vb6Type::Array)) != 0;
                Vb6Type argBase86 = static_cast<Vb6Type>(
                    static_cast<uint16_t>(argT86) & ~static_cast<uint16_t>(Vb6Type::Array));
                bool scalarLike =
                    argT86 == Vb6Type::Long || argT86 == Vb6Type::Integer
                    || argT86 == Vb6Type::Boolean || argT86 == Vb6Type::Byte
                    || argT86 == Vb6Type::Double || argT86 == Vb6Type::Single
                    || argT86 == Vb6Type::Currency || argT86 == Vb6Type::Date
                    || argT86 == Vb6Type::LongPtr || argT86 == Vb6Type::String
                    || argT86 == Vb6Type::Object || argBase86 == Vb6Type::Object
                    || argIsArr86;
                if (scalarLike) {
                    // Fix 198: 布尔实参 → VT_BOOL (_Generic 会按 int16_t 选 VT_I2).
                    argVal = boxToVariant(node.positional[i].get(), argVal);
                }
            }
        }
        // Fix 038b-2: calleeParams 为空 (运行时/内置函数) 时的参数类型转换.
        // 通过 getRuntimeParamCType 查找期望的 C 类型, 当实参为 Variant 时
        // 自动插入 VARIANT→具体类型提取函数. 解决 vb6_ErrRaise, vb6_BSTR_Assign,
        // vb6_BSTR_Concat, vb6_StrCmp 等运行时函数的 C2440 错误.
        // 注意: 仅使用 cExprIsVariant (C 字符串级) 和 knownVariantVars_ 检测,
        // 不使用 isDefinitelyVariantExpr (AST 级), 因为符号表中的 Variant 返回类型
        // 可能与实际 C 函数返回类型不一致 (如 prop_get 返回 void* 而非 vb6_VARIANT).
        // Fix 091e: 形参符号类型不可信 (无非空参数表, 或符号表把 builtin 形参登记为
        // Variant/Empty — 如 LenB/Len 被推断为 Variant) 时同样查运行时表. 此前仅
        // calleeParams 为空才查表, 导致 Variant 实参原样传给 BSTR 形参:
        // cTlsSocket 200 vb6_LenB_BSTR(vb6_VariantFromComResult(ComGetProp(...))) C2440.
        bool rtParamTypeUsable = (i >= calleeParams.size())
            || calleeParams[i].type == Vb6Type::Variant
            || calleeParams[i].type == Vb6Type::Empty;
        if (!isByRef && rtParamTypeUsable) {
            std::string rtParamType = getRuntimeParamCType(callee, i);
            if (!rtParamType.empty() && rtParamType != "vb6_VARIANT"
                 && rtParamType != "vb6_VARIANT*") {
                bool argIsVariant = cExprIsVariant(argVal);
                // Fix 113c: 实参是裸 COM 结果 (void* 承载 COM VARIANT*) 时, 语义即
                // Variant — 必须判为 Variant 才能触发下方 Variant→具体类型提取
                // (vb6_VariantFromComResult + 按 rtParamType 提取). 否则 void* 裸传
                // 给 BSTR/String 形参 (如 UserControl.TextWidth(cAxisItem.Item(i)))
                // 在 C 层 void*→BSTR 隐式转换, 运行时把垃圾指针喂给
                // GetTextExtentPoint32W 崩溃 (Charts 2020 ucChartBar_DrawVertical).
                // 注意: 仅裸 vb6_ComCall/ComGetProp/ComGetObjectProp (返回 void*);
                // vb6_ComCallInt/BSTR/Double/Object 返回具体类型, 前缀不会匹配.
                // 此检测刻意不进 cExprIsVariant (Fix 113 已回退): 那样会让
                // vb6_ComCall 在 ByVal Variant 打包分支 (Fix 086) 被误判为"已装箱
                // vb6_VARIANT", 跳过 vb6_VariantFromValue 导致 void*→vb6_VARIANT C2440.
                if (!argIsVariant) {
                    bool argIsComRes113 =
                        (argVal.compare(0, 12, "vb6_ComCall(") == 0)
                        || (argVal.compare(0, 15, "vb6_ComGetProp(") == 0)
                        || (argVal.compare(0, 21, "vb6_ComGetObjectProp(") == 0);
                    if (argIsComRes113) argIsVariant = true;
                }
                // 也检查已知 Variant 变量
                if (!argIsVariant && node.positional[i]->kind == ASTNodeKind::IdentifierExpr) {
                    auto& idArg = static_cast<IdentifierExpr&>(*node.positional[i]);
                    std::string argLower = idArg.name;
                    std::transform(argLower.begin(), argLower.end(), argLower.begin(), ::tolower);
                    if (knownVariantVars_.count(argLower)) argIsVariant = true;
                }
                // Fix 089i: 函数调用返回 Variant 的实参 (VB 未声明返回类型 =
                // Variant, 如 json_ParseErrorMessage(...) 作为 Err.Raise 的
                // Description) — cExprIsVariant 只查字符串前缀, 对
                // "vb6_<mod>_<fn>(...)" 用户函数调用文本不生效 → C2440.
                // 查被调函数符号的返回类型: kind Function/PropertyGet 且
                // returnType Variant → C 层返回 vb6_VARIANT, 需按参数类型提取.
                // 不匹配对象属性 (As Object → void*, returnType 非 Variant).
                if (!argIsVariant && node.positional[i]->kind == ASTNodeKind::IndexOrCallExpr) {
                    auto& callNode89i = static_cast<IndexOrCallExpr&>(*node.positional[i]);
                    if (callNode89i.callee
                        && callNode89i.callee->kind == ASTNodeKind::IdentifierExpr) {
                        std::string calleeLower89i =
                            static_cast<IdentifierExpr&>(*callNode89i.callee).name;
                        std::transform(calleeLower89i.begin(), calleeLower89i.end(),
                                       calleeLower89i.begin(), ::tolower);
                        // Fix 113d: 多态内置 IIf 的符号类型登记为 Variant
                        // (semantic_analyzer builtin_funcs.inc: addBuiltinFunc("IIf",
                        // Vb6Type::Variant)), 但 codegen 已按上下文发出**具体类型**
                        // 版本 vb6_IIfLong/vb6_IIfDouble/vb6_IIfBSTR — 三者 C 返回
                        // 类型分别是 int32_t/double/BSTR, 均**不是** vb6_VARIANT.
                        // 此处若仍按符号类型判为 Variant, 会对具体类型结果二次套
                        // vb6_VariantToLong/String 提取 → C2440
                        // (int32_t/double/BSTR → vb6_VARIANT). Charts 2020
                        // LabelPlus.PictureFromURL:
                        //   UserControl.AsyncRead sUrl, vbAsyncTypeByteArray, "URL",
                        //       IIf(UseCache, 0, vbAsyncReadForceUpdate)
                        //   → 原生成 vb6_VariantToLong(vb6_IIfLong(...)) C2440.
                        // 注: vb6_IIfVariant( 才是真正的 vb6_VARIANT, 仍需按符号
                        // 判为 Variant (用于 Variant→BSTR/数值 提取).
                        // 语义与 isDefinitelyVariantExpr 的 isPolymorphicBuiltin
                        // 一致 (iif/choose/switch), 此处仅针对已发射的 IIf 具体形态.
                        bool emittedConcretePoly89i =
                            argVal.rfind("vb6_IIfLong(", 0) == 0
                            || argVal.rfind("vb6_IIfDouble(", 0) == 0
                            || argVal.rfind("vb6_IIfBSTR(", 0) == 0;
                        if (!emittedConcretePoly89i)
                        for (const auto& [k89i, s89i] : symTab_.moduleScope()->symbols()) {
                            if (s89i->kind != SymbolKind::Function
                                && s89i->kind != SymbolKind::PropertyGet) continue;
                            if (s89i->lowerName != calleeLower89i) continue;
                            // 返回类型语义 Variant (VB 未声明返回 = Variant).
                            // 对象属性 As Object → void* (returnType Object),
                            // 不会误匹配; PropertyGet 的 Unknown 属性(可能 void*)
                            // 保守不包, 仅模块 Function 的 Unknown(未推断出类型但
                            // 模块函数 C 返回 vb6_VARIANT)才包.
                            bool retIsVar89i =
                                (s89i->type == Vb6Type::Variant
                                 || s89i->type == Vb6Type::Empty);
                            if (s89i->kind == SymbolKind::PropertyGet) {
                                if (retIsVar89i) argIsVariant = true;
                            } else if (retIsVar89i
                                       || s89i->type == Vb6Type::Unknown) {
                                argIsVariant = true;
                            }
                            break;
                        }
                    }
                }
                if (argIsVariant) {
                    // Fix 090c0: 实参是 COM 结果且被判定为 Variant —
                    // vb6_ComCall/vb6_ComGetProp 返回 void* (COM Variant 承载),
                    // 提取 (VariantToSafeArray1D 等期望 vb6_VARIANT 按值) 前需先
                    // vb6_VariantFromComResult 转为 vb6_VARIANT. 如 cLogs
                    // MakeLogContent: Join(CacheDatas(i), Span) — Collection 元素
                    // 是 Variant 数组 → 原生成 VariantToSafeArray1D(vb6_ComCall(...))
                    // → C2440 (无法从 void* 转换为 vb6_VARIANT).
                    // 精确前缀避免误伤 vb6_ComCallInt( (返回 int32) /
                    // vb6_ComGetStringProp( (返回 BSTR) / vb6_ComCallObject( (对象).
                    bool isComResultVal =
                        (argVal.find("vb6_ComCall(") == 0)
                        || (argVal.find("vb6_ComGetProp(") == 0)
                        || (argVal.find("vb6_ComGetObjectProp(") == 0);
                    if (isComResultVal) {
                        argVal = "vb6_VariantFromComResult(" + argVal + ")";
                    }
                    if (rtParamType == "int32_t" || rtParamType == "int16_t"
                        || rtParamType == "uint8_t" || rtParamType == "LONG") {
                        argVal = "vb6_VariantToLong(" + argVal + ")";
                    } else if (rtParamType == "double" || rtParamType == "float") {
                        argVal = "vb6_VariantToDouble(" + argVal + ")";
                    } else if (rtParamType == "BSTR") {
                        // Fix 089k: 顶层表达式是否已是 BSTR — 原 049b 用
                        // argVal.find("vb6_BSTR") 子串包含判断, 对内部实参含
                        // vb6_BSTR_FromStr(...) 的用户函数调用 (如
                        // json_ParseErrorMessage(...)) 误判"已 BSTR" → 跳过
                        // vb6_VariantToString 包装 → C2440 (vb6_VARIANT→BSTR).
                        // 收紧为剥外层括号后的顶层前缀判断.
                        auto exprTopIsBstr = [](const std::string& s) -> bool {
                            std::string t = s;
                            while (t.size() >= 2 && t.front() == '(' && t.back() == ')') {
                                int depth = 0;
                                bool whole = true;
                                for (size_t k = 0; k < t.size(); k++) {
                                    if (t[k] == '(') depth++;
                                    else if (t[k] == ')') {
                                        depth--;
                                        if (depth == 0 && k < t.size() - 1) {
                                            whole = false;
                                            break;
                                        }
                                    }
                                }
                                if (!whole) break;
                                t = t.substr(1, t.size() - 2);
                            }
                            static const char* bstrTopPrefixes[] = {
                                "vb6_BSTR_",
                                "VB6_SA_AT(BSTR,",
                                // 已由 vb6_VariantToString 转换的结果必为 BSTR,
                                // 顶层识别避免二次包装 → C2440 (BSTR→vb6_VARIANT).
                                "vb6_VariantToString(",
                                // Fix 091h: 移除 "vb6_VariantFromComResult(" —
                                // 该函数返回 vb6_VARIANT (vb6rtl.h:918), 不是 BSTR.
                                // 旧前缀把 COM 属性结果误判为"已 BSTR" → 跳过提取 →
                                // C2440 "vb6_VARIANT → BSTR" (cTlsSocket 200 LenB,
                                // cToolsList 25 prop_let_Filter, cAesCBC 25 StrConv,
                                // Demo_Database 310/506/678 BSTR_Concat).
                                "vb6_BSTR_FromStr("};
                            for (auto* bp : bstrTopPrefixes) {
                                size_t bpl = strlen(bp);
                                if (t.compare(0, bpl, bp) == 0) return true;
                            }
                            return false;
                        };
                        if (!exprTopIsBstr(argVal)) {
                            argVal = "vb6_VariantToString(" + argVal + ")";
                        }
                    } else if (rtParamType == "void*") {
                        argVal = "vb6_VariantToObjectVal(" + argVal + ")";
                    } else if (rtParamType == "vb6_SafeArray1D*") {
                        argVal = "vb6_VariantToSafeArray1D(" + argVal + ")";
                    }
                }
            }
        }
        // Fix 113j: ByRef **Variant** 形参 + 具体类型左值实参归一.
        // 很多 .ctl 的辅助函数写成 `text As Variant` (无 ByVal → ByRef), C 形参
        // 是 vb6_VARIANT*. 若实参是 BSTR/Long/Double 之类**具体的左值**
        // (UDT 字段 / me->字段 / VB6_SA_AT(...).field), 前面 ByRef 链只发出
        // &(左值) → 类型是 BSTR*/int32_t*, 被调方按 VARIANT 解析 → vt 与数据字段
        // 全部错位 → 文本渲染成一排"豆腐块"空框.
        //   Charts 2020 ucTreeMaps.DrawText(text As Variant):
        //     DrawText hGraphics, m_Serie(i).SerieName, ...
        //   原生成 vb6_ucTreeMaps_DrawText(..., &(VB6_SA_AT(...).SerieName), ...)
        //   = BSTR* 传进 vb6_VARIANT* 形参.
        // 这里按实参 VB 类型改发字段式 VARIANT 复合字面量.
        // 实参本身就是 Variant 左值时 inferExprType 返回 Variant → 不匹配任何
        // 分支, 保持原 &(x) 不变 (这是正确的, 不能破坏).
        if (isByRef && i < calleeParams.size()
            && calleeParams[i].type == Vb6Type::Variant
            && argVal.size() > 3 && argVal.compare(0, 2, "&(") == 0
            && argVal.back() == ')') {
            Vb6Type argT113j = inferExprType(*node.positional[i]);
            std::string inner113j = argVal.substr(2, argVal.size() - 3);
            if (argT113j == Vb6Type::String) {
                argVal = "(&(vb6_VARIANT){.vt=VT_BSTR, .bstrVal=" + inner113j + "})";
            } else if (argT113j == Vb6Type::Long || argT113j == Vb6Type::Integer
                       || argT113j == Vb6Type::Byte) {
                argVal = "(&(vb6_VARIANT){.vt=VT_I4, .lVal=(int32_t)(" + inner113j + ")})";
            } else if (argT113j == Vb6Type::Double || argT113j == Vb6Type::Single
                       || argT113j == Vb6Type::Currency || argT113j == Vb6Type::Date) {
                argVal = "(&(vb6_VARIANT){.vt=VT_R8, .dblVal=(double)(" + inner113j + ")})";
            } else if (argT113j == Vb6Type::Boolean) {
                argVal = "(&(vb6_VARIANT){.vt=VT_BOOL, .boolVal=(int16_t)(" + inner113j + ")})";
            }
        }
        // Fix <vbeclipse> <VBFlexGridDemo>: 形参 C 类型是 `vb6_ComIface_X*` (VB
        // `As SomeInterface`, 从 TLB 或工程接口来)，而实参是 void* (VB `As Object`
        // 后期绑定值) 时，显式补一层 (T*) —— cl 19.29 静默隐式转 void*→typed*,
        // cl 19.51 直接 C2172 "actual parameter is not a pointer" (VTableHandle.bas
        // `Call ReplaceIOleControl(This)` 六处一次爆红)。VB6 语义本就是 Object →
        // Interface 的 QI 转换，这里做的是**类型层的显式化**，不改运行期行为。
        // 判据走 cParamTypedComIfaceCType (全局 symTab_.lookup)，与 wrap guard
        // 那三条同源 —— cTypeForDeclaredTypeName 走 mapTypeRef 只看当前模块，
        // CI 上 TLB 接口挂全局那一份查不到, 会兜底成 void* (dbg210 实测)。
        if (i < calleeParams.size() && !isByRef) {
            const std::string ifaceCT = cParamTypedComIfaceCType(calleeParams[i]);
            if (!ifaceCT.empty()) {
                if (argVal.rfind("(" + ifaceCT + ")", 0) != 0) {
                    argVal = "(" + ifaceCT + ")" + argVal;
                }
            }
        }
        args.push_back(std::move(argVal));
    }

    // Fix 081c: For named-arg path, track which Optional params were actually passed
    // (before gap-filling sets filled[i]=true for padding params too)