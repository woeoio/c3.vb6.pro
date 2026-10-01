// cgen_expr_call_callee_params.inc - visit(IndexOrCallExpr&) 函数体片段：callee 形参解析（calleeParams / builtin / Declare / ANSI）
// 由 src/backend/expr/cgen_expr_call.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 1554~1770 行，逐行未改。

    std::vector<ParameterInfo> calleeParams;
    // Fix 030b: 跟踪被调用者是否为 builtin — builtin 用 calleeParams 仅为触发
    // Fix 024 P2/Fix 029 的参数包装, 但 RTL C 签名不接受 Optional padding 和
    // IsMissing _has_ flag 尾叜 (那些只适用于用户定义函数). 见 line 3207/3237.
    bool calleeIsBuiltin = false;
    // Fix 042a: Declare 函数 (DeclareSub/DeclareFunc) 的 C 签名不接受 _has_ 尾叜,
    // 与 builtin 类似 — Optional padding 仍需要 (C 函数期望所有参数),
    // 但 IsMissing _has_ flags 不应追加.
    bool calleeIsDeclare = false;
    // Fix 041b: Track whether calleeParams was successfully resolved (even if 0 params).
    // Used to distinguish "params not looked up" from "looked up with 0 params" (e.g., Property Get
    // with no params) — needed for arg truncation when args > params.
    bool calleeParamsFound = false;
    // 从IdentifierExpr或MemberAccessExpr获取被调用函数名, 在符号表中查找
    // Fix 027: 加入 DeclareSub / DeclareFunc — 否则 WinAPI Declare 的 ByRef 参数无法 emit &,
    //          导致 ByRef UDT/SafeArray/标量 全部按值传递 (例如 SOCKADDR_IN → SOCKADDR_IN* 错误).
    if (node.callee && node.callee->kind == ASTNodeKind::IdentifierExpr) {
        auto& idExpr = static_cast<IdentifierExpr&>(*node.callee);
        // Fix <VBFlexGridDemo/GA 36766856082>: 内置函数优先于同名类成员属性 (剥 $ 后撞名)。
        // Task #40 (40960eb) 给使用点剥掉类型后缀 ⇒ `Left$(Value, 1)` 到 cgen 时名字是
        // `Left`, 而控件自己声明了 `Public Property Get Left() As Single` ⇒ 下面
        // lookupModule 命中这个**形参表为空**的属性符号 ⇒ calleeParams 空 ⇒ 实参被清空,
        // 生成 `vb6_Left()` (C2198 参数太少 / C2440 BSTR→double)。
        // VB6 语义: 属性 Left 无形参, `Left(x, y)` 带实参 ⇒ 只能是内置函数。
        // 剥后缀前 (BASE) 名字是 `Left$`, 查不到该属性符号, 天然走内置; 这里把那条
        // 隐式约定写成显式判据。用户 Function 同名不受影响 (只让位给 Property*)。
        std::string idLower40 = Symbol::toLower(idExpr.name);
        if (!idLower40.empty() && idLower40.back() == '$') idLower40.pop_back();
        const bool calleeIsBuiltinFuncName40 = (kBuiltinFuncNames.count(idLower40) != 0);
        // 尝试多种查找方式: 先lookupModule(过程符号), 再lookup(嵌套作用域)
        Symbol* funcSym = symTab_.lookupModule(idExpr.name);
        if (!funcSym || (funcSym->kind != SymbolKind::Sub && funcSym->kind != SymbolKind::Function
            && funcSym->kind != SymbolKind::PropertyGet && funcSym->kind != SymbolKind::PropertyLet
            && funcSym->kind != SymbolKind::PropertySet
            && funcSym->kind != SymbolKind::DeclareSub && funcSym->kind != SymbolKind::DeclareFunc)) {
            funcSym = symTab_.lookup(idExpr.name);
        }
        if (funcSym && (funcSym->kind == SymbolKind::Sub || funcSym->kind == SymbolKind::Function
            || funcSym->kind == SymbolKind::PropertyGet || funcSym->kind == SymbolKind::PropertyLet
            || funcSym->kind == SymbolKind::PropertySet
            || funcSym->kind == SymbolKind::DeclareSub || funcSym->kind == SymbolKind::DeclareFunc)
            // 见上 Fix <VBFlexGridDemo>: 内置函数名 + 符号是属性 ⇒ 让位给内置函数,
            // 不采用属性的 (空) 形参表, 免得实参被清空。
            && !(calleeIsBuiltinFuncName40 && (funcSym->kind == SymbolKind::PropertyGet
                 || funcSym->kind == SymbolKind::PropertyLet
                 || funcSym->kind == SymbolKind::PropertySet))) {
            // 重载变体 (O2): 语义层选定变体时, 形参表/返回类型以变体符号为准
            if (!node.calleeOvlSuffix.empty()) {
                if (auto* ovSym = symTab_.lookupModuleOverloadBySuffix(
                        idExpr.name, node.calleeOvlSuffix)) {
                    funcSym = ovSym;
                }
            }
            calleeParams = funcSym->params;
            calleeIsBuiltin = funcSym->isBuiltin;
            calleeParamsFound = true;
            // Fix 042a: Track Declare functions for _has_ flag suppression
            if (funcSym->kind == SymbolKind::DeclareSub || funcSym->kind == SymbolKind::DeclareFunc) {
                calleeIsDeclare = true;
            }
            // Bug #3 fix: 检测函数返回类型为StdPicture/IPictureDisp (COM Picture)
            // 用于Picture属性赋值时选择vb6_SetControlPictureFromCom
            if (!funcSym->variableTypeName.empty()) {
                std::string vtLower = funcSym->variableTypeName;
                std::transform(vtLower.begin(), vtLower.end(), vtLower.begin(), ::tolower);
                if (vtLower == "stdpicture" || vtLower == "ipicturedisp" || vtLower == "ipicture") {
                    lastExprIsComPicture_ = true;
                }
            }
        }
    } else if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr) {
        // Module.Method 或 obj.Method 调用: 查找方法名的参数签名
        // Fix 027: 同步加入 DeclareSub/DeclareFunc (模块方法形式的 declare 调用).
        auto& maExpr = static_cast<MemberAccessExpr&>(*node.callee);

        // Fix 033: 优先用类感知查找. 原 symTab_.lookupModule(maExpr.memberName) 在
        // 跨模块同名方法冲突下 (例如 cAsyncSocket.Create / cTlsSocket.Create / cPassword.Create
        // 6+ 个类共享 storageKey="create"), 命中首个注册者而非对象真实类的方法,
        // 导致 Optional 参数 _has_ 标志个数填错 → C2197 ("too many arguments").
        // 限制条件: 仅当 maExpr.object 是 IdentifierExpr 时启用 — 此时无副作用,
        // 不需要重复 emit obj 表达式即可推断 className. 复杂链式 obj 留给原回退路径.
        bool classAwareResolved = false;
        if (maExpr.object && maExpr.object->kind == ASTNodeKind::IdentifierExpr) {
            std::string className;
            auto& idObj = static_cast<IdentifierExpr&>(*maExpr.object);
            std::string nameLower = Symbol::toLower(idObj.name);
            if (nameLower == "me") {
                // 当前类模块实例 — className = 本模块名
                if (isClassModule_) className = moduleName_;
            } else {
                auto it = knownClassVars_.find(nameLower);
                if (it != knownClassVars_.end()) {
                    className = it->second;
                } else {
                    // 不是已知类变量 → 可能是模块名或外部类名自身
                    // (例: ToolsStr.HasStr — ToolsStr 是 .bas 模块; cWinsock.SomeStaticMethod — cWinsock 类)
                    // findClassMemberCallParams 内部会按 sourceModule/moduleName_ 匹配
                    // Fix 084z-3: 对象是当前模块的属性 (Property Get 返回类实例, 如
                    // cTlsRemaster.pvSocket → cTlsSocket) 时, inferClassTypeOfExpr
                    // 可推断出类名 → 优先按类解析方法形参; 否则 findClassMemberCallParams
                    // 按属性名查找失败, 回退 lookupModule 命中 storageKey 同名冲突的
                    // 错误类 (SyncReceiveArray 按 cWinsock 12 参展开 → C2197).
                    className = idObj.name;
                    // Fix 090g: 对象是当前函数名引用 (函数体内 FuncName.Add(...) —
                    // 即本函数返回对象, Fix 088c 已把 vb6_ret_FuncName 注册进
                    // knownClassVars_). 例 cToolsStr.SplitLinesToCollection As cCollection
                    // → SplitLinesToCollection.Add Mid(...). 此前 fallthrough 到模块名
                    // 路径, findClassMemberCallParams 失败后又回退 lookupModule(storageKey),
                    // 命中跨类同名方法抢占的错误类 (Add → 其它类的 Add), 形参表错 →
                    // C2197/C2440 (Item 按 ByRef String 打包, 多出 Optional int32).
                    if (currentProc_ && nameLower == Symbol::toLower(currentProc_->name)) {
                        std::string retVar090g = "vb6_ret_" + nameLower;
                        auto itRet090g = knownClassVars_.find(retVar090g);
                        if (itRet090g != knownClassVars_.end()) {
                            className = itRet090g->second;
                        }
                    }
                    if (className == idObj.name) {
                        std::string inferredClass = inferClassTypeOfExpr(*maExpr.object);
                        if (!inferredClass.empty()) {
                            className = inferredClass;
                        }
                    }
                }
            }
            if (!className.empty()) {
                std::vector<ParameterInfo> params;
                bool isBuiltin = false;
                if (findClassMemberCallParams(className, maExpr.memberName, params, isBuiltin)) {
                    calleeParams = std::move(params);
                    calleeIsBuiltin = isBuiltin;
                    classAwareResolved = true;
                    calleeParamsFound = true;
                }
            }
        }
        // Fix 042b: For method chains and complex object expressions (e.g.,
        // db.Table("t").OrderByDesc("id").Limit(10).Offset(20)), the object is
        // not a simple IdentifierExpr but an IndexOrCallExpr or MemberAccessExpr.
        // Use inferClassTypeOfExpr to determine the class type, then look up
        // member params. This enables Optional param padding for cross-module
        // method chain calls.
        if (!classAwareResolved && maExpr.object
            && maExpr.object->kind != ASTNodeKind::IdentifierExpr) {
            std::string className = inferClassTypeOfExpr(*maExpr.object);
            if (!className.empty()) {
                std::vector<ParameterInfo> params;
                bool isBuiltin = false;
                if (findClassMemberCallParams(className, maExpr.memberName, params, isBuiltin)) {
                    calleeParams = std::move(params);
                    calleeIsBuiltin = isBuiltin;
                    classAwareResolved = true;
                    calleeParamsFound = true;
                }
            }
        }

        // Fix 092z: 内置全局对象 (Clipboard/Screen/Printer/Forms/Debug/Err/App) 的成员
        // 由 visit(MemberAccessExpr) 的硬编码分支解析为 RTL 函数 (如 Clipboard.SetText
        // → vb6_Clipboard_SetText), VB 侧**不存在**形参表. 若此处仍用
        // lookupModule(memberName) 兜底, 会命中项目中同名的类成员:
        //   cQRcode.cls "Public Function SetText(ByVal Content As Variant) As cQRcode"
        // 使 FLogs.frm "Clipboard.SetText Text1.Text" 的形参表被误当成 ByVal Variant
        // → 实参被 Fix 086 / Fix 024P2 包装成 vb6_VariantFromValue(...) →
        // vb6_Clipboard_SetText(vb6_VARIANT) C2440 (vb6_VARIANT→BSTR).
        // 这几个名字是 VB6 保留的全局对象, 不可能被用户标识符占用, 直接跳过兜底.
        bool builtinGlobalObj092z = false;
        if (maExpr.object && maExpr.object->kind == ASTNodeKind::IdentifierExpr) {
            static const std::unordered_set<std::string> builtinGlobalObjs092z = {
                "clipboard", "screen", "printer", "forms", "debug", "err", "app",
                // Fix <vbeclipse>: UserControl.Controls 集合 — 禁止 lookupModule 兜底
                // 把 .Item/.Add/.Remove 抢挂到工程里恰好有同名方法的类 (List) 上
                // (实测: Controls.Item(p) → vb6_List_Item((*p)), 接收者被首参顶掉)。
                "controls",
                "console"};  // Fix 161: Console 对象 (对齐 twinBASIC)
            builtinGlobalObj092z = builtinGlobalObjs092z.count(
                Symbol::toLower(static_cast<IdentifierExpr&>(*maExpr.object).name)) > 0;
        }

        // Fix 033 回退: 类感知未命中 (对象为链式表达式 / className 找不到方法符号 /
        // 需要匹配 DeclareSub/DeclareFunc 等) → 用原 class-unaware lookupModule 兜底,
        // 保留旧行为兼容性.
        if (!classAwareResolved && !builtinGlobalObj092z) {
            Symbol* funcSym = symTab_.lookupModule(maExpr.memberName);
            // Fix 084y-7: VBA 内置函数 (VBA.Replace / VBA.Mid$ / VBA.Val 等) 注册在
            // 全局内置符号表, 不在模块作用域 — lookupModule 必然失败. 回退全表
            // lookup (去 $ 后缀: Mid$ → Mid), 拿到含 Optional 的完整形参表,
            // 才能填充默认参数 (vb6_Replace 声明6参, VB6 调用只传3参).
            if (!funcSym) {
                std::string fnName = maExpr.memberName;
                if (!fnName.empty() && fnName.back() == '$') fnName.pop_back();
                funcSym = symTab_.lookup(fnName);
            }
            if (funcSym && (funcSym->kind == SymbolKind::Sub || funcSym->kind == SymbolKind::Function
                || funcSym->kind == SymbolKind::PropertyGet || funcSym->kind == SymbolKind::PropertyLet
                || funcSym->kind == SymbolKind::PropertySet
                || funcSym->kind == SymbolKind::DeclareSub || funcSym->kind == SymbolKind::DeclareFunc)) {
                calleeParams = funcSym->params;
                calleeIsBuiltin = funcSym->isBuiltin;
                calleeParamsFound = true;
                // Fix 042a: Track Declare functions for _has_ flag suppression
                if (funcSym->kind == SymbolKind::DeclareSub || funcSym->kind == SymbolKind::DeclareFunc) {
                    calleeIsDeclare = true;
                }
                // Bug #3 fix: 检测跨模块函数返回类型为StdPicture/IPictureDisp
                if (!funcSym->variableTypeName.empty()) {
                    std::string vtLower = funcSym->variableTypeName;
                    std::transform(vtLower.begin(), vtLower.end(), vtLower.begin(), ::tolower);
                    if (vtLower == "stdpicture" || vtLower == "ipicturedisp" || vtLower == "ipicture") {
                        lastExprIsComPicture_ = true;
                    }
                }
            }
        }
    }

    // Fix 041b: WithMemberExpr callee — With-block member call (e.g., .Add(...) inside With)
    // needs calleeParams for Optional param padding and IsMissing _has_ flags. Without this,
    // cross-module With-block calls with Optional params generate C2198 (too few arguments).
    if (node.callee && node.callee->kind == ASTNodeKind::WithMemberExpr) {
        auto& wmExpr = static_cast<WithMemberExpr&>(*node.callee);
        if (!withObjectInfoStack_.empty()) {
            const auto& info = withObjectInfoStack_.back();
            if (info.kind == WithObjKind::ClassInstance && !info.className.empty()) {
                std::vector<ParameterInfo> params;
                bool isBuiltin = false;
                if (findClassMemberCallParams(info.className, wmExpr.memberName, params, isBuiltin)) {
                    calleeParams = std::move(params);
                    calleeIsBuiltin = isBuiltin;
                    calleeParamsFound = true;
                }
            }
            // Fix 210b: With 目标是 void* / className 未知的 ClassInstance (如
            // `With l_Perspective.Folders` / `With Folder.Views`, With 临时变量落在
            // void*) 时, WithMemberExpr 实际由 cgen_expr_with.cpp:191-209 按名字
            // lookupModule 解析成熟知成员函数 (实测 .Item → vb6_List_Item, 形参
            // ByVal Variant Index). calleeParams 不填 → Fix 210 装箱不生效 →
            // 4x int32_t/int → vb6_VARIANT C2440 (ucPerspective.c 1807 等).
            // 与发射侧同源取符号 (name-match), 让装箱/缺省参数补齐对齐.
            if (!calleeParamsFound && info.kind == WithObjKind::ClassInstance) {
                Symbol* wmSym210b = symTab_.lookupModule(wmExpr.memberName);
                if (!wmSym210b) wmSym210b = symTab_.lookup(wmExpr.memberName);
                if (wmSym210b && (wmSym210b->kind == SymbolKind::Sub
                    || wmSym210b->kind == SymbolKind::Function
                    || wmSym210b->kind == SymbolKind::PropertyGet
                    || wmSym210b->kind == SymbolKind::PropertyLet
                    || wmSym210b->kind == SymbolKind::PropertySet)) {
                    calleeParams = wmSym210b->params;
                    calleeIsBuiltin = wmSym210b->isBuiltin;
                    calleeParamsFound = true;
                }
            }
        }
    }

    // M22: 检测Declare ANSI函数调用 - 需要BSTR->ANSI转换
    bool isDeclareAnsiCall = false;
    if (node.callee && node.callee->kind == ASTNodeKind::IdentifierExpr) {
        auto& idExpr = static_cast<IdentifierExpr&>(*node.callee);
        std::string funcLower = idExpr.name;
        std::transform(funcLower.begin(), funcLower.end(), funcLower.begin(), ::tolower);
        if (knownDeclareAnsi_.count(funcLower)) isDeclareAnsiCall = true;
    }

    std::vector<std::string> args;