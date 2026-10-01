// cgen_expr_member_obj_dispatch.inc - 函数体片段：COM·接口·后期绑定·Variant 分派与 UDT 字段/函数名引用解析
// 由 src/backend/expr/cgen_expr_member.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 363~552 行，逐行未改。

        // 优先级0: COM前期绑定成员访问 (P6.3, Dim x As FileSystemObject)
        // 有具体类型信息的COM变量, 通过vtable直接调用而非IDispatch::Invoke
        if (knownTypedComVars_.count(objLower)) {
            emitExpr(*node.object);
            comObjExpr_ = lastExpr_;       // 保存对象表达式
            comMemberName_ = node.memberName;  // 保存成员名
            isComMarker_ = true;           // 标记为COM调用
            isEarlyBoundCom_ = true;       // P6.3: 标记为前期绑定
            earlyBoundSym_ = knownTypedComVars_[objLower];  // ComClass符号
            lastExpr_ = lastExpr_;         // 保持不变
            return;
        }

        // tB Interface 契约 (B04): 接口变量的成员访问 → 走 vb6_ivtbl_<I> 槽表
        auto itIvrefVar = knownIvrefVars_.find(objLower);
        if (itIvrefVar != knownIvrefVars_.end()) {
            emitExpr(*node.object);
            comObjExpr_ = lastExpr_;
            comMemberName_ = node.memberName;
            ivrefIfaceName_ = itIvrefVar->second;  // 交给 com_bind 的 ivref 分支消费
            isComMarker_ = true;
            isEarlyBoundCom_ = false;
            earlyBoundSym_ = nullptr;
            lastExpr_ = lastExpr_;
            return;
        }

        // P6.4: 接口引用成员访问 (Dim x As IFoo) → 通过vtable调用
        // 设置接口标记, 由IndexOrCallExpr/AssignmentStmt识别
        auto itIfaceVar = knownIfaceVars_.find(objLower);
        if (itIfaceVar != knownIfaceVars_.end()) {
            emitExpr(*node.object);
            comObjExpr_ = lastExpr_;       // 接口引用变量C表达式
            comMemberName_ = node.memberName;  // 接口方法名
            isComMarker_ = true;           // 复用COM标记机制
            isEarlyBoundCom_ = false;      // 不是COM前期绑定
            // 设置接口标记
            earlyBoundSym_ = nullptr;       // P6.4接口不是ComClass
            lastExpr_ = lastExpr_;         // 保持接口引用变量名
            return;
        }

        // 优先级1: COM对象成员访问 (Object类型变量, 后期绑定)
        // COM对象的成员名不在符号表中, 需通过IDispatch::Invoke调用
        // 设置COM标记, 由IndexOrCallExpr/AssignmentStmt/SetStmt识别并处理
        // 注意: 项目类实例 (Dim b As Button / Dim WithEvents b As Button) 的 C 类型也是
        // void*, 可能被误注册进 knownObjectVars_, 但它不是 IDispatch — 类实例是纯 C
        // 结构体 (vb6_cls_Button*), 用 vb6_ComCall 会解引用 vtable 崩溃 (0xC0000005).
        // 已知是类实例时跳过 COM 晚绑定, 交给下面的"优先级2"做直接分发.
        // Fix <vbeclipse>: knownObjectVars_ 按**过程**不清空 (模块内不同过程同名
        // `Dim l_Tab As Object` 与 `Dim l_Tab As Variant` 并存), 先判 Object 会把
        // Variant 槽位的接收者裸发出 → vb6_ComGetIntProp(l_Tab, ...) C2172
        // (ucTabStrip.c 312/314/387). 同过程内 Variant 声明优先走 P24-04 分支
        // vb6_VariantToObject(&l_Tab) 提取对象指针.
        if (knownObjectVars_.count(objLower) && !knownClassVars_.count(objLower)
            && !knownUdtVars_.count(objLower) && !knownVariantVars_.count(objLower)) {
            emitExpr(*node.object);
            comObjExpr_ = lastExpr_;       // 保存对象表达式
            comMemberName_ = node.memberName;  // 保存成员名
            isComMarker_ = true;           // 标记为COM调用
            // lastExpr_设为对象表达式(可用作值), 具体调度由上层决定
            lastExpr_ = lastExpr_;         // 保持不变 (对象C表达式)
            return;
        }

        // P24-04: Variant变量成员访问 (Variant持有COM对象, 后期绑定)
        // VB6: For Each b In col: b.Index -> vb6_ComCall(vb6_VariantToObject(&b), L"Index", ...)
        // Variant变量可能持有IDispatch指针, 成员访问需要通过COM晚绑定
        if (knownVariantVars_.count(objLower)) {
            emitExpr(*node.object);
            std::string varExpr = std::move(lastExpr_);
            comObjExpr_ = "vb6_VariantToObject(&" + varExpr + ")";
            comMemberName_ = node.memberName;
            isComMarker_ = true;
            lastExpr_ = "vb6_VariantFromComResult(vb6_ComGetProp(vb6_VariantToObject(&" + varExpr + "), L\"" + node.memberName + "\"))";  // P25: Variant COM prop Get → VARIANT
            return;
        }

        // P24-04: VB_GlobalNameSpace promoted函数的成员访问 (如 VBMAN.Version)
        // VBMAN是提升到全局的函数(sGlobal._sGlobal.VBMAN()), 返回cVBMAN COM对象
        // VBMAN.Version = VBMAN().Version = 先创建sGlobal单例, 调用VBMAN()获取cVBMAN, 再访问.Version
        if (auto* gnsSym = symTab_.lookup(objIdent.name)) {
            if (gnsSym->kind == SymbolKind::ComGlobalNs) {
                // 1. 生成sGlobal单例创建 + promoted方法调用表达式
                std::string progIdWide = "L\"" + gnsSym->comProgId + "\"";
                std::string methodName = gnsSym->comGlobalNsMethodName;
                std::string gnsCallExpr = "vb6_ComCallObject(vb6_CreateObject(" + progIdWide + "), L\"" + methodName + "\", NULL, 0)";
                // 2. 设置COM marker, .memberName将在IndexOrCallExpr中消费
                comObjExpr_ = gnsCallExpr;
                comMemberName_ = node.memberName;
                isComMarker_ = true;
                isEarlyBoundCom_ = false;
                earlyBoundSym_ = nullptr;
                lastExpr_ = gnsCallExpr + "  /* GlobalNs." + methodName + " */";
                return;
            }
        }

        // 查找成员名称的符号
        // Fix 031: UDT 变量字段访问 — 必须在 memSym 查找之前拦截.
        // 当 obj 是已知 UDT 变量 (knownUdtVars_) 时, obj.member 总是结构体字段访问
        // (obj.member 或 (*obj).member 形式, 取决于 emitExpr 对该 UDT 变量的求值),
        // 不应进入 memSym 路径把 obj 当作模块名 (或 uOutput 误当类实例).
        // 现象: Private Sub pvTlsBuildClientHello(uCtx As UcsTlsContext, uOutput As UcsBuffer)
        //       其中 UcsBuffer UDT 含 Size As Long 字段, 同时项目里又有 cByteBuffer.cls
        //       声明了 Property Get Size, memSym 查找命中该跨模块属性 → 误生成
        //       vb6_cByteBuffer_Size 函数引用 (而非 (*uOutput).Size 字段访问), 触发 C2065.
        // 已知 UDT 变量与类实例 (knownClassVars_) / Object (knownObjectVars_) 在注册阶段
        // 互斥, 此处不冲突.
        if (knownUdtVars_.count(objLower)) {
            emitExpr(*node.object);
            std::string obj = std::move(lastExpr_);
            // Fix 085: UDT 字段是对象 (项目类/Collection/COM) 时, 在生成的访问文本
            // 上追加 "/* udt objfield <CType> */" 注释标记, 由外层 MemberAccessExpr
            // 消费 → 转类方法调用 / COM dispatch (否则 obj.Field.Method 触发 C2039/
            // C2224). 例: uFile.SourceArchive (As cZipArchive) 后接 .frCopyCompressed.
            lastExpr_ = appendUdtObjFieldMarker(obj, knownUdtVars_[objLower], node.memberName);
            return;
        }

        // Fix 084z-4: 函数名引用返回值的成员访问 (RHS 读取侧, 与 LHS 的
        // Fix 084o-5 对称). VB6 函数体内可用函数名访问返回对象的成员
        // (如 cJson: RootItems.Add NewItem.RootItem, NewItem As cJson) →
        // 展开 vb6_ret_NewItem->RootItem; 而非误当模块名落入 M22 生成
        // vb6_NewItem_RootItem (C2065 未声明).
        if (node.object && node.object->kind == ASTNodeKind::IdentifierExpr) {
            auto& objIdent084z = static_cast<IdentifierExpr&>(*node.object);
            std::string objLower084z = Symbol::toLower(objIdent084z.name);
            if (currentProc_ && objLower084z == Symbol::toLower(currentProc_->name)
                && (currentProc_->kind == SymbolKind::Function || currentProc_->kind == SymbolKind::PropertyGet)
                && !currentReturnVar_.empty()) {
                bool isCls084z = currentReturnCType_.find("vb6_cls_") != std::string::npos;
                if (isCls084z) {
                    // Fix 088d: 函数返回类实例时, 函数名引用的成员可能是方法/属性
                    // (ReturnJson.Decode s) 而非仅公开数据字段 (NewItem.RootItem).
                    // 先尝试类方法/属性分发 (canonical 调用), 找不到才视为字段:
                    //   vb6_cJson_Decode((void*)vb6_ret_ReturnJson, s)
                    // 此前无条件生成 vb6_ret_X->member → C2039 (Decode 不是
                    // vb6_cls_cJson 的成员).
                    std::string retCls088d = currentReturnCType_.substr(8);
                    if (!retCls088d.empty() && retCls088d.back() == '*')
                        retCls088d.pop_back();
                    std::string fn088d =
                        resolveClassMemberCall(retCls088d, node.memberName);
                    if (!fn088d.empty()) {
                        // 类虚表 (ai/022 B08e-6 站点⑨): 接收者是本函数返回值的裸变量
                        // (`vb6_ret_X`, 纯读) → 与 ⑤ 同法接派发。`Function Maker() As KBase`
                        // 体内 `Maker.Speak()` 此前静默绑基类实现, 而 Maker 完全可以返回一个
                        // 更深类的实例 (继承来的方法体在基类里, me 是派生实例)。
                        // 派发串与 this 实参共用同一个串; Let/Set 方向 3.4b 无槽, 不改写。
                        std::string thisArg088d = "(void*)" + currentReturnVar_;
                        if (!suppressVirtDispatch_
                            && fn088d.find("_prop_let_") == std::string::npos
                            && fn088d.find("_prop_set_") == std::string::npos) {
                            const std::string dispatchFn088d = virtDispatchCallee(
                                retCls088d, node.memberName, thisArg088d,
                                /*mustDispatch=*/true, node.loc);
                            if (!dispatchFn088d.empty()) fn088d = dispatchFn088d;
                        }
                        if (asCallCallee_) {
                            pendingChainObj_ = thisArg088d;
                            lastExpr_ = fn088d;  // 外层补参数
                        } else {
                            lastExpr_ = fn088d + "(" + thisArg088d + ")";
                        }
                        return;
                    }
                    // 类无此成员 → 公开数据字段 (RootItem 等)
                    lastExpr_ = currentReturnVar_ + "->" + cIdent(node.memberName);
                } else if (currentReturnCType_ == "void*"
                           || currentReturnCType_.find("vb6_ComIface_") != std::string::npos) {
                    // Fix 089c2: 当前函数返回内置 COM 对象 (As Collection / As Object →
                    // C void*) 时, 函数体内 FuncName.Add(...)/FuncName.Item(...) 走
                    // COM dispatch (vb6_ComCall) — 如 cZipArchive.pvEnumFiles As
                    // Collection → pvEnumFiles.Add path, key. 此前 isCls084z false 走
                    // appendUdtObjFieldMarker(void* 无字段) → 生成
                    // vb6_ret_X.Add(...) 结构成员调用 → C2224 (void* 上 .Add).
                    // 与 Fix 088d (类返回) / Fix 085 (UDT 返回) 对齐.
                    // Fix 090ag: As Dictionary (项目外 COM 类, C 类型 vb6_ComIface_
                    // IDictionary* 而非 void*) 的函数返回同样走 COM dispatch —
                    // ToolsJsonVba json_ParseObject As Dictionary 体内
                    // json_ParseObject.Item(Key) = v 此前生成
                    // vb6_ret_X.Item(...) → C2037 (未定义结构 vb6_ComIface_IDictionary).
                    lastExpr_ = currentReturnVar_;
                    comObjExpr_ = "(void*)" + currentReturnVar_;  // 统一 void* 对象表达式
                    comMemberName_ = node.memberName;
                    isComMarker_ = true;
                    isEarlyBoundCom_ = false;
                    earlyBoundSym_ = nullptr;
                } else {
                    // Fix 085: 当前函数返回 UDT 时, 字段为对象(Collection/COM)需标记,
                    // 使外层链式成员访问走 COM/类方法路径 (例: retUdt.Stream.VfsSetFilePointer)
                    lastExpr_ = appendUdtObjFieldMarker(currentReturnVar_, currentReturnCType_,
                                                        node.memberName);
                }
                return;
            }
        }

        // Fix 086: UDT返回属性的字段读取 (CurLang.Index — Property Get CurLang()
        // As TypeLang). 属性返回结构体右值, C 允许对函数返回值取成员 f().Field;
        // 此前误判为 Module.X → 生成未声明的 vb6_CurLang_Index (C2065).
        // 注意: 仅当对象是当前类的 Property Get 且返回UDT时走此路径.
        if (node.object && node.object->kind == ASTNodeKind::IdentifierExpr && isClassModule_) {
            Symbol* propSym = symTab_.lookupModule(objIdent.name);
            if (propSym && propSym->kind == SymbolKind::PropertyGet
                && !propSym->variableTypeName.empty()) {
                Symbol* udtSym86 = lookupDotted(propSym->variableTypeName);
                if (udtSym86 && udtSym86->kind == SymbolKind::UserDefinedType) {
                    emitExpr(*node.object);
                    std::string objExpr86 = std::move(lastExpr_);
                    lastExpr_ = objExpr86 + "." + cIdent(node.memberName)
                              + "  /* UdtProp." + node.memberName + " */";
                    return;
                }
            }
        }

        auto* memSym = symTab_.lookupModule(node.memberName);
        // Fix 110e: VB6 宿主伪对象 UserControl / Ambient / Extender / PropertyPage 之
        // 后的成员不是"别的模块的同名成员". 它们是宿主状态符号, 由
        // rtl/core/vb6rtl/vb6rtl_userctl.h 提供 (vb6_UserControl_* / vb6_Ambient_* /
        // vb6_Extender_* / vb6_PropertyPage_*), 生成名 = vb6_<伪对象名>_<成员>
        // (即下方 M22 分支的形式).
        // 若在此按"模块名.成员"解析, 成员会命中别处注入的同名外部符号 →
        //   UserControl.Enabled    → 命中 LabelPlus 的 Property Get Enabled
        //                            → vb6_LabelPlus_Enabled   (C2065)   ← 实测
        //   Ambient.ForeColor      → vb6_LabelPlus_ForeColor      (C2065)   ← 实测
        // (在 LabelPlus 模块内恰好回退到 sourceMod = "UserControl" 才碰巧正确,
        //  这也是同一表达式在不同模块生成不同符号名、仅部分模块报错的原因.)
        // 跳过本块, 交由 M22 分支以对象名作前缀生成正确的宿主标识符.
        const bool hostPseudoObj110e =
            objLower == "usercontrol" || objLower == "ambient" ||
            objLower == "extender" || objLower == "propertypage";
        if (!hostPseudoObj110e && memSym
            && (memSym->kind == SymbolKind::Sub || memSym->kind == SymbolKind::Function
                    || memSym->kind == SymbolKind::PropertyGet
                    || memSym->kind == SymbolKind::PropertyLet
                    || memSym->kind == SymbolKind::PropertySet)) {
