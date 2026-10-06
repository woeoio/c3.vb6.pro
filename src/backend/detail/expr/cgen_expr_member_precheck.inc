// cgen_expr_member_precheck.inc - 函数体片段：前置分派 —— AddressOf 模块限定函数 / Err·内建对象 / 控件数组 / 内置对象方法
// 由 src/backend/expr/cgen_expr_member.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 13~131 行，逐行未改。

    // Fix <vbeclipse>: UserControl.Controls 集合 (裸 `Controls.Item(k)` / `.Add(...)`
    // / For Each 形态的 `.Count`) —— 必须在一切类感知解析之前拦截, 否则工程里
    // 恰好有同名方法的类 (List.Item / List.Remove / List.Add) 会把成员抢挂过去,
    // 接收者被首参顶掉 (ucPerspective 实测: Controls.Item(p) → vb6_List_Item((*p)))。
    // rev14: RTL 侧 vb6_UC_Controls() 已返回真实的控件集合单例 (按当前 UC 实例回落),
    // 不再用恒 NULL 的 vb6_UserControl_Controls —— 后者使 `Set x = Controls.Item(...)`
    // 恒为 NULL, ucFolder_AddView(NULL) 解引用即 0xC0000005。
    if (isDesignerModule_ && node.object && node.object->kind == ASTNodeKind::IdentifierExpr
        && Symbol::toLower(static_cast<IdentifierExpr&>(*node.object).name) == "controls") {
        comObjExpr_ = "(void*)vb6_UC_Controls()";
        comMemberName_ = node.memberName;
        isComMarker_ = true;
        isEarlyBoundCom_ = false;
        earlyBoundSym_ = nullptr;
        lastExpr_ = "(void*)vb6_UC_Controls()";
        return;
    }
    // Fix 086: AddressOf 模块限定函数 (AddressOf ToolsTlsThunks.pvCallCollectionItem)
    // — 解析器将 `ToolsTlsThunks.pvCallCollectionItem` 交给 AddressOfExpr(仅取
    // object名) + 外层MemberAccessExpr(成员名). 此前生成非法的
    // (void*)vb6_<mod>_<mod>.<member> (C2065). 在此合成为真正的函数指针.
    if (node.object && node.object->kind == ASTNodeKind::AddressOfExpr) {
        auto& ao86 = static_cast<AddressOfExpr&>(*node.object);
        (void)ao86;
        std::string fnName86 = node.memberName;
        Symbol* fnSym86 = symTab_.lookupModule(fnName86);
        if (!fnSym86) fnSym86 = symTab_.lookup(fnName86);
        if (fnSym86 && (fnSym86->kind == SymbolKind::Sub || fnSym86->kind == SymbolKind::Function)) {
            // 账 #184: 被取址过的过程交**桩**地址 (与 visit(AddressOfExpr) 同一处口径)
            lastExpr_ = "(void*)" + addressOfTargetCName(fnSym86,
                cProcName(fnName86, fnSym86->access,
                          fnSym86->isExternal ? fnSym86->sourceModule : ""));
        } else {
            lastExpr_ = "(void*)" + addressOfTargetCName(fnSym86,
                                                         cProcName(fnName86, AccessLevel::Private));
        }
        return;
    }
    // ============================================================
    // P20-44: OLE 拖放的 **DataObject 形参**成员 —— `Data.GetText` / `Data.GetFileCount`。
    // 处理器形参 `void** Data` 不是 IDispatch, 掉 COM 派发运行期必崩 (实测 0xC0000005)。
    // **必须放在 visit(MemberAccessExpr&) 的最前面**: DataObject 是 COM 类, 走到
    // form_builtin / generic_access / resolveComValue / resolveComMarkerForPack
    // 任何一条都会发 ComGetStringProp (实测三条路都拦过、全被更早的分支抢走)。
    // 判据是 dataObjectParams_ (cgen_decl_func.cpp 注册参数时收的), **不能查 symTab_**
    // —— codegen 生成处理器体时形参不在符号表 (实测 lookup found=0)。
    // DataObject 指针 = *Data, 所以 RTL 调用首参是 `(void*)(*Data)`。
    // ============================================================
    if (node.object && node.object->kind == ASTNodeKind::IdentifierExpr) {
        auto& ddObj44 = static_cast<IdentifierExpr&>(*node.object);
        std::string ddLower44 = Symbol::toLower(ddObj44.name);
        auto* ddSym44 = symTab_.lookup(ddLower44);
        if (dataObjectParams_.count(ddLower44)) {
            std::string memDD44 = Symbol::toLower(node.memberName);
            std::string ddArg44 = "(void*)(*" + ddObj44.name + ")";
            if (memDD44 == "gettext")      { lastExpr_ = "vb6_oleDD_GetText(" + ddArg44 + ")"; return; }
            if (memDD44 == "getfilecount") { lastExpr_ = "vb6_oleDD_GetFileCount(" + ddArg44 + ")"; return; }
            // Data.Files (非调用形态, 后面跟 .Count): 留尾部注释标记, 外层 MemberAccess
            // 的 generic path 消费它 (搜 oleDD-files)。
            if (memDD44 == "files")        { lastExpr_ = ddArg44 + " /* oleDD-files */"; return; }
        }
        // ---- P20-44: 控件的 OLEDrag 方法 (无实参形态) ----
        // 语句 `X.OLEDrag` 的 callee 求值会走到这里 (visit 最前), 比 com_bind 的
        // "无参数 → ComCall" 分支早。实参形态 (allowedEffects) 走 call_callee_member。
        auto itCtlDDp = knownFormControls_.find(ddLower44);
        if (itCtlDDp != knownFormControls_.end()
            && Symbol::toLower(node.memberName) == "oledrag") {
            std::string hwndDDp = makeCtrlHwndArg(ddLower44, itCtlDDp->second);
            static const char* ddSrcEvP[4] = { "_OLEStartDrag", "_OLESetData",
                                               "_OLEGiveFeedback", "_OLECompleteDrag" };
            std::string regDDp;
            for (int ddK = 0; ddK < 4; ddK++) {
                auto* ddSym = symTab_.lookup(ddObj44.name + ddSrcEvP[ddK]);
                if (!ddSym) continue;
                std::string ddFn = cProcName(ddSym->name, AccessLevel::Private);
                regDDp += "  { extern void " + ddFn + "(); vb6_OLEDrop_SetHandler("
                        + hwndDDp + ", " + std::to_string(ddK + 2) + ", (void*)" + ddFn + "); }\n";
            }
            lastExpr_ = (std::string)"{ void* ddObj44 = (void*)vb6_oleDD_MakeDataObject("
                      "vb6_GetControlText(" + hwndDDp + "), NULL, 0);\n"
                      + regDDp
                      + "  vb6_OLEDrag_Start((void*)" + hwndDDp + ", ddObj44, 3); }";
            return;
        }
    }

    // Fix 010k: Err/builtin object member access — check FIRST before anything else
    // This must be at the very top to avoid any other code path consuming the node
    if (node.object && node.object->kind == ASTNodeKind::IdentifierExpr) {
        auto& _objId = static_cast<IdentifierExpr&>(*node.object);
        std::string _objLower = _objId.name;
        std::transform(_objLower.begin(), _objLower.end(), _objLower.begin(), ::tolower);
        std::string _memLower = node.memberName;
        std::transform(_memLower.begin(), _memLower.end(), _memLower.begin(), ::tolower);

        if (_objLower == "err" || _objLower == "lasterror") {
            // Fix 056b: LastError 属性返回 VBA.ErrObject (全局单例),
            // .Number/.Description/.Source 等价于 Err.Number 等.
            if (_memLower == "number")      { lastExpr_ = "vb6_ErrNumber()";      return; }
            if (_memLower == "description") { lastExpr_ = "vb6_ErrDescription()"; return; }
            if (_memLower == "source")      { lastExpr_ = "vb6_ErrSource()";      return; }
            if (_memLower == "lastdllerror") { lastExpr_ = "vb6_ErrLastDllError()"; return; }
            if (_memLower == "helpfile")    { lastExpr_ = "(BSTR)0";              return; }
            if (_memLower == "helpcontext") { lastExpr_ = "0";                    return; }
            if (_memLower == "clear")       { lastExpr_ = "vb6_ErrClear()";       return; }
            if (_memLower == "raise")       { lastExpr_ = "vb6_ErrRaise";         return; }
        }

        // ai/022 B09: `MyBase.<成员>` = 显式基调用 (去虚化, 直调基类会跑的那一份实现)。
        // 按**接收者名字**在这里接管、不占 token: 与上面的 Err / VBA 同一条前置通路 (最早、
        // 小写名匹配), 于是 lexer/parser 一行不动 —— 零新语法工程的逐字节护栏自动成立。
        if (_objLower == "mybase") {
            if (tryEmitMyBaseMember(node)) return;
        }

        // Fix 084y-6: VBA 模块成员 (VBA.vbCr / VBA.Replace 等). 编译器不注册
        // "VBA" 模块符号 → M22 fallback 生成 vb6_VBA_<member> (rtl 无此名 →
        // C2065). 常量内联字面量; 函数生成裸 rtl 名 vb6_<member> (去 $ 后缀,
        // 如 VBA.Mid$ → vb6_Mid), 由外层 IndexOrCallExpr 的 Fix 033 回退
        // (lookupModule + lookup 全表) 解析参数表并填充 Optional 默认值.
        if (_objLower == "vba" || _objLower == "vba5") {
            static const std::unordered_map<std::string, std::string> vbaConsts = {
                {"vbcr", "13"}, {"vblf", "10"}, {"vbtab", "9"},
                {"vbnullchar", "0"}, {"vbnullstring", "vb6_BSTR_FromStr(L\"\")"},
                {"vbcrif", "vb6_BSTR_FromStr(L\"\\r\\n\")"},
                {"vbnewline", "vb6_BSTR_FromStr(L\"\\r\\n\")"},
                {"vbempty", "0"}, {"vbnull", "1"}, {"vbinteger", "2"}, {"vblong", "3"},
                {"vbsingle", "4"}, {"vbdouble", "5"}, {"vbcurrency", "6"}, {"vbdate", "7"},
                {"vbstring", "8"}, {"vbobject", "9"}, {"vberror", "10"}, {"vbboolean", "11"},
                {"vbvariant", "12"}, {"vbdecimal", "14"}, {"vbbyte", "17"}, {"vblonglong", "20"},
                {"vbuserdefinedtype", "36"}, {"vbarray", "8192"},
                {"vbtrue", "-1"}, {"vbfalse", "0"},
            };
            auto itC = vbaConsts.find(_memLower);
            if (itC != vbaConsts.end()) { lastExpr_ = itC->second; return; }
            std::string fnName = node.memberName;
            if (!fnName.empty() && fnName.back() == '$') fnName.pop_back();
            Symbol* vbaFn = symTab_.lookup(fnName);
            if (vbaFn && (vbaFn->kind == SymbolKind::Function || vbaFn->kind == SymbolKind::Sub
                          || vbaFn->kind == SymbolKind::PropertyGet
                          || vbaFn->kind == SymbolKind::DeclareSub
                          || vbaFn->kind == SymbolKind::DeclareFunc)) {
                // Fix 086: VBA.Now() 等无参内置函数作为裸值引用 (非callee上下文,
                // 如 pvToFileTime(VBA.Now)) — 需带调用括号, 否则函数设计符裸名 → C2440
                static const std::unordered_set<std::string> vbaZeroArgFns = {
                    "now", "date", "time", "timer", "freefile", "command",
                    "curdir", "erl", "doevents"
                };
                // Fix 093a: 用符号表里的声明拼写 (VBA.varType → vb6_VarType). 此前用
                // 源码拼写 → vb6_varType, 与 RTL/intrinsic 符号名不一致 → LNK2019.
                std::string fnCanon = vbaFn->name.empty() ? fnName : vbaFn->name;
                std::string fnLower86 = Symbol::toLower(fnCanon);
                if (fnLower86 == "rnd") {
                    lastExpr_ = "vb6_Rnd(1)";
                } else if (vbaZeroArgFns.count(fnLower86)) {
                    lastExpr_ = "vb6_" + cIdent(fnCanon) + "()";
                } else {
                    lastExpr_ = "vb6_" + cIdent(fnCanon);
                }
                return;
            }
        }
    }

    // Fix 110n: CallByName(...).成员 — RTL 的 vb6_CallByName 返回 vb6_VARIANT, 其值
    // 可能是对象 (VbGet 取到 Font 等对象属性). 对 VARIANT 结构体取成员在 C 里非法:
    //   CallByName(oCtrl, PropFont, VbGet).Size
    //     → vb6_CallByName(...).Size  C2039 ("Size" 不是 "vb6_VARIANT" 的成员)
    // 赋值时更会退化成 "<struct> = <double>" C2088. 必须走 COM 晚绑定: 先从 VARIANT
    // 取出对象 (vb6_VariantToObjectVal) 再 vb6_ComGetXxxProp / ComSetProp.
    // 读上下文由外层 (BinaryExpr/赋值) 的 resolveComValue 消费 marker;
    // 写上下文由 AssignmentStmt 的 COM SetProp 分支消费.
    // 实测: Charts 2020 ClsResizer.cls 107/147.
    if (node.object && node.object->kind == ASTNodeKind::IndexOrCallExpr) {
        auto& cbCall = static_cast<IndexOrCallExpr&>(*node.object);
        if (cbCall.callee && cbCall.callee->kind == ASTNodeKind::IdentifierExpr) {
            auto& cbId = static_cast<IdentifierExpr&>(*cbCall.callee);
            if (Symbol::toLower(cbId.name) == "callbyname") {
                emitExpr(*node.object);
                std::string cbObjExpr = std::move(lastExpr_);
                comObjExpr_ = "vb6_VariantToObjectVal(" + cbObjExpr + ")";
                comMemberName_ = node.memberName;
                isComMarker_ = true;
                isEarlyBoundCom_ = false;
                earlyBoundSym_ = nullptr;
                lastExpr_ = comObjExpr_;
                return;
            }
        }
    }

    // P7.6: 控件数组属性读取 ctrlArr(idx).Property
    // 必须在IdentifierExpr分支前检查, 因为cmdBtn(0)的object是IndexOrCallExpr
    if (node.object && node.object->kind == ASTNodeKind::IndexOrCallExpr) {
        auto& idxExpr = static_cast<IndexOrCallExpr&>(*node.object);
        if (idxExpr.callee && idxExpr.callee->kind == ASTNodeKind::IdentifierExpr) {
            auto& arrIdent = static_cast<IdentifierExpr&>(*idxExpr.callee);
            std::string arrNameLower = arrIdent.name;
            std::transform(arrNameLower.begin(), arrNameLower.end(), arrNameLower.begin(), ::tolower);
            if (knownControlArrays_.count(arrNameLower)) {
                auto itCtrl = knownFormControls_.find(arrNameLower);
                if (itCtrl != knownFormControls_.end()) {
                    std::string idxArg;
                    if (!idxExpr.positional.empty()) {
                        emitExpr(*idxExpr.positional[0]);
                        idxArg = std::move(lastExpr_);
                    } else {
                        idxArg = "0";
                    }
                    std::string ctrlHwndArr = "vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(arrIdent.name) + ", " + idxArg + ")";
                    // 账 #191: 工程内 UserControl 的**数组元素** —— 方法/属性必须与单枚控件走同一个
                    // 出口 (`emitUcInstanceMemberExpr`, Fix 112 抽出来的那一个)。以前这条路上没有它,
                    // `ucChartBar1(i).AddSerie ...` 就整条落进 COM 兜底 ⇒ 交回空值 (真工程实测)。
                    auto itUcArr = knownUserControlCtrlVars_.find(arrNameLower);
                    if (itUcArr != knownUserControlCtrlVars_.end()
                        && emitUcInstanceMemberExpr(itUcArr->second, node.memberName, ctrlHwndArr)) {
                        return;
                    }
                    std::string readFn = getControlPropReadFn(itCtrl->second, node.memberName);
                    if (!readFn.empty()) {
                        lastExpr_ = readFn + "(" + ctrlHwndArr + ")";
                        return;
                    }
                    // Fix 0xx: OCX 控件数组元素属性 (如 NewTab1(c).Theme) — 内置 readFn
                    // 匹配不到 (第三方 ActiveX 特有属性), 走 COM 后期绑定读属性. 此前未
                    // 拦截, `NewTab1(c).Theme` 原样输出到 C → C2224. 数组元素取
                    // vb6_CtrlArr_GetAt 的 HWND 后按属性读取 (Theme 等返回 BSTR).
                    comObjExpr_ = ctrlHwndArr;
                    comMemberName_ = node.memberName;
                    isComMarker_ = true;
                    isEarlyBoundCom_ = false;
                    earlyBoundSym_ = nullptr;
                    lastExpr_ = ctrlHwndArr;
                    return;
                }
            }
        }
    }
    // 内置对象方法: Debug.Print → vb6_DebugPrint
    // 检查 object 是否是 IdentifierExpr
