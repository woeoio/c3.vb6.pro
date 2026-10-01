// cgen_base_generate_state_scan.inc - generate 函数体片段：P6.11 类成员 BSTR 扫描 + Fix 010n 模块级 UDT 持久化
// 由 src/backend/cgen_base.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 132~255 行，逐行未改。


    // P6.11: 类模块成员变量类型扫描 (用于方法体内的BSTR安全赋值)
    // 生成代码时类成员的C名格式为 m_Xxx (首字母大写), lowercase 后为 m_xxx
    classBstrMembers_.clear();
    classLongMembers_.clear();
    classDoubleMembers_.clear();
    classVariantMembers_.clear();  // Fix 037
    classByteArrayMembers_.clear();  // Fix 140
    classUdtMembers_.clear();  // Fix 010n
    moduleUdtMembers_.clear();  // Fix 010n (扩展): 普通模块模块级UDT变量
    classMemberVars_.clear();  // Fix 010r
    if (isClassModule_) {
        for (auto& decl : module.declarations) {
            if (decl->kind == ASTNodeKind::VariableDecl) {
                auto& var = static_cast<VariableDecl&>(*decl);
                // Fix 010r: Register ALL class member variable names (lowercase)
                {
                    std::string mLower = "m_" + var.name;
                    std::transform(mLower.begin(), mLower.end(), mLower.begin(), ::tolower);
                    std::string oLower = var.name;
                    std::transform(oLower.begin(), oLower.end(), oLower.begin(), ::tolower);
                    classMemberVars_.insert(mLower);
                    classMemberVars_.insert(oLower);
                }
                if (var.asType) {
                    Vb6Type vtype = resolveArrayElemType(var.asType.get());
                    // 注册 C 字段名 m_Xxx 的 lowercase: m_xxx
                    std::string mLower = "m_" + var.name;
                    std::transform(mLower.begin(), mLower.end(), mLower.begin(), ::tolower);
                    // 同时注册不带 m_ 前缀的原始名 lowercase
                    std::string oLower = var.name;
                    std::transform(oLower.begin(), oLower.end(), oLower.begin(), ::tolower);
                    if (vtype == Vb6Type::String) {
                        classBstrMembers_.insert(mLower);
                        classBstrMembers_.insert(oLower);
                    } else if (vtype == Vb6Type::Long || vtype == Vb6Type::Integer || vtype == Vb6Type::Boolean) {
                        classLongMembers_.insert(mLower);
                        classLongMembers_.insert(oLower);
                    } else if (vtype == Vb6Type::Double || vtype == Vb6Type::Single || vtype == Vb6Type::Date) {
                        classDoubleMembers_.insert(mLower);
                        classDoubleMembers_.insert(oLower);
                    }
                    // Fix 140: Byte() 类成员数组 (仅数组, 标量 Byte 不走此路).
                    // 与 knownByteArrayVars_ 对称 — 过程入口 copy 回该集合, 使
                    // `m_Caption = New_Caption` (BSTR) 改写为 vb6_StringToByteArray
                    // 而非把 BSTR 当 SafeArray1D* 直接赋给字段 (LabelPlus 右侧卡片空白).
                    else if (vtype == Vb6Type::Byte
                             && var.asType->kind == ASTNodeKind::ArrayTypeRef) {
                        classByteArrayMembers_.insert(mLower);
                        classByteArrayMembers_.insert(oLower);
                    }
                    // Fix 037: Variant 类成员注册到 classVariantMembers_.
                    // 用于 IndexOrCallExpr 处理 me->VarField(idx) / obj.VarField(args),
                    // 应当 emit vb6_VariantArrayGet(&me->VarField, idx) 而非把 Variant
                    // 字段当函数调用 (C2064 term does not evaluate to a function).
                    else if (vtype == Vb6Type::Variant) {
                        // Fix 090bx: 排除外部对象类型名 (As Collection / 未知
                        // COM 类型等不在 typeSys_ 内置表且符号表无条目的名字).
                        // resolveArrayElemType 对它们回落 Variant, 但 C 层字段是
                        // void* COM 指针. 若误注册为 classVariantMembers_, 类内
                        // m_Col(i) 会生成 vb6_VariantArrayGet(&me->m_Col, i)
                        // (C2172 实参不是指针 / C2440), 应走 void* COM 字段路径
                        // → vb6_ComCall(me->m_Col, L"Item", {ComPackInt(i)}, 1).
                        // 仅内置 Variant/无符号简单名中的真正 Variant 才注册.
                        // vbeclipse: Fix 090bx 的漏网之鱼 — `As <项目类>`
                        // (Private WithEvents m_PopupMenu As PopupMenu /
                        // WithEvents SplitBar As SplitBar) 的 resolveArrayElemType
                        // 同样回落 Variant, 但 C 层字段是 `void*` 类指针, 误入
                        // classVariantMembers_ 后 With/Set 会按 Variant 语义解包 →
                        // C2440 (ucCaption.c 269 / ucSplitBar.c 172:
                        // void* _vb6_with_N = vb6_VariantToObjectVal(me->m_PopupMenu)).
                        // 名字能解析到**工程类** (工程级登记表, 不受同名成员遮蔽) 或
                        // Com* 对象类型时一律不注册.
                        bool realVariantField = true;
                        if (var.asType->kind == ASTNodeKind::SimpleTypeRef) {
                            auto& simpleV = static_cast<SimpleTypeRef&>(*var.asType);
                            std::string nm = simpleV.name;
                            std::transform(nm.begin(), nm.end(), nm.begin(), ::tolower);
                            if (nm != "variant" && nm != "var") {
                                // 工程类判定走**工程级**类名表 (driver 注入, 不受遮蔽).
                                // 模块内符号表在这里不可用: vbeclipse 实测同名成员遮蔽 ——
                                //   ucCaption `Private WithEvents m_PopupMenu As PopupMenu`
                                //     类型名 "PopupMenu" 被同名 Sub 占位 (SymbolKind::Sub);
                                //   ucSplitBar `Private WithEvents SplitBar As SplitBar`
                                //     类型名 "SplitBar" 被字段自身占位 (SymbolKind::Variable);
                                // 四种模块内查找 (lookup/lookupModule/lookupTypeSymbol/
                                // lookupDotted) 都拿不到 Class → 旧守卫放行 → void* 字段
                                // 误按 Variant 解包 (C2440 ucCaption.c 269 / ucSplitBar.c 172).
                                // classViewByName()/clsreg_ 同样不可用: 工程无任何 Inherits
                                // 时 driver_classchain.cpp:355 的 D24 护栏让整表早退为空.
                                if (!projectClassNameOf(simpleV.name).empty()) {
                                    realVariantField = false;
                                } else {
                                    const Symbol* s1 = symTab_.lookup(simpleV.name);
                                    const Symbol* s2 = symTab_.lookupModule(simpleV.name);
                                    const Symbol* s3 = lookupTypeSymbol(simpleV.name);
                                    const Symbol* s4 = lookupDotted(simpleV.name);
                                    auto isObjKind = [](SymbolKind k) {
                                        return k == SymbolKind::Class ||
                                               k == SymbolKind::ComClass ||
                                               k == SymbolKind::ComInterface ||
                                               k == SymbolKind::ComModule ||
                                               k == SymbolKind::ComGlobalNs;
                                    };
                                    // 四种查找**逐一**判, 不取首个非空: 同名 Sub/Variable
                                    // 会遮蔽后面的 Class 命中 (Fix 103/107 同款陷阱).
                                    if (isObjKind(s1 ? s1->kind : SymbolKind::Variable) ||
                                        isObjKind(s2 ? s2->kind : SymbolKind::Variable) ||
                                        isObjKind(s3 ? s3->kind : SymbolKind::Variable) ||
                                        isObjKind(s4 ? s4->kind : SymbolKind::Variable)) {
                                        realVariantField = false;
                                    } else if (!s1 && !s2 && !s3 && !s4) {
                                        // 090bx 原口径: 四种查找都无条目 → 未知对象类型名
                                        realVariantField = false;
                                    }
                                }
                            }
                        }
                        if (realVariantField) {
                            classVariantMembers_.insert(mLower);
                            classVariantMembers_.insert(oLower);
                        }
                    }
                        // Fix 010n: 记录UDT类型成员变量 (用于With块类型检测)
                        if (var.asType->kind == ASTNodeKind::SimpleTypeRef) {
                            auto& simple = static_cast<SimpleTypeRef&>(*var.asType);
                            auto* udtSym = lookupDotted(simple.name);
                            if (udtSym && udtSym->kind == SymbolKind::UserDefinedType) {
                                std::string udtCType = "vb6_type_" + cIdent(simple.name);
                                classUdtMembers_[oLower] = udtCType;
                                classUdtMembers_[mLower] = udtCType;
                            }
                            // NOTE: 类类型成员变量 (Private WithEvents m_oSocket As cTlsSocket 等)
                            // 由 visit(VariableDecl) 在 trackOnly_=true 路径 (cgen_decl.cpp:605-628)
                            // 统一注册到 knownClassVars_, 此处无需重复注册.
                            // Fix 014 改在 cgen_expr.cpp visit(MemberAccessExpr) 通用 fallback 中
                            // 提取 C 表达式尾部标识符以支持链式访问 (me.m_oSocket.Create).
                        }
                } else {
                    // Fix 089j: VB6 无 As 类型声明的成员变量默认 Variant
                    // (如 cTimer.cls 的 `Private mParentsColKey`), 必须注册到
                    // classVariantMembers_, 否则 prop_let/prop_get 内的
                    // `me->mParentsColKey = value` / `ret = me->mParentsColKey`
                    // 不做 Variant 打包/解包 → C2440 (int→vb6_VARIANT /
                    // vb6_VARIANT→int32_t). C 层 struct 字段即 vb6_VARIANT
                    // (mapTypeRef(nullptr) 返回 vb6_VARIANT).
                    std::string mLower89j = "m_" + var.name;
                    std::transform(mLower89j.begin(), mLower89j.end(), mLower89j.begin(), ::tolower);
                    std::string oLower89j = var.name;
                    std::transform(oLower89j.begin(), oLower89j.end(), oLower89j.begin(), ::tolower);
                    classVariantMembers_.insert(mLower89j);
                    classVariantMembers_.insert(oLower89j);
                }
            }
        }
    }
    // Fix 010n (扩展): 普通模块(.bas/.frm)模块级UDT变量持久化.
    // 与类模块 classUdtMembers_ 对称 — 过程开头 knownUdtVars_.clear() 后,
    // 模块级 UDT 变量 (如 Private m_uData As UcsCryptoData) 需要跨过程恢复,
    // 否则过程内 With m_uData 被误分类为 COM 对象 → 生成 vb6_ComGetIntProp 而非
    // 结构体成员访问 (C2224/C2198 错误).
    if (!isClassModule_) {
        for (auto& decl : module.declarations) {
            if (decl->kind == ASTNodeKind::VariableDecl) {
                auto& var = static_cast<VariableDecl&>(*decl);
                if (var.asType && var.asType->kind == ASTNodeKind::SimpleTypeRef) {
                    auto& simple = static_cast<SimpleTypeRef&>(*var.asType);
                    auto* udtSym = lookupDotted(simple.name);
                    if (udtSym && udtSym->kind == SymbolKind::UserDefinedType) {
                        std::string oLower = var.name;
                        std::transform(oLower.begin(), oLower.end(), oLower.begin(), ::tolower);
                        moduleUdtMembers_[oLower] = "vb6_type_" + cIdent(simple.name);
                    }
                }
            }
        }
    }
