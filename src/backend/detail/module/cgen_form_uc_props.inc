// cgen_form_uc_props.inc - emitDesignerControlDecls 函数体片段：紧跟
// `static const vb6_UserControlDesc vb6_<ctl>_ucHostDesc = {` 之前发射
// UserControl **自有 Property 的按名桥**（表 + thunk）。
//
// Fix <vbeclipse> rev18. 为什么需要这一档:
//   工程里 `Dim l_ucFolder As Variant: Set l_ucFolder = Controls.Item(id)` 之后的
//   `l_ucFolder.LeftPos` / `l_ucFolder.Move .LeftPos, .TopPos, …` 是**晚绑定**
//   (工程自己把变量声明成 Variant), cgen 发的是
//     vb6_ComGetIntProp(实例, L"LeftPos") / vb6_ComSetProp(实例, L"LeftPos", …)
//   而"实例"是原生 `vb6_cls_ucFolder*` 结构体 (不是 IDispatch)。宿主模型
//   (uc_hostmodel_{getprop,setprop}.inc) 只认内建成员 (Caption/Visible/Left/Top/
//   Width/Height/hWnd/Refresh/Move/…), 没有名字表就只剩"返回 Empty"。
//   实证 play78 (--arch x86): ucPerspective.Refresh 里四个位置全读到 0,
//   7 个 ucFolder 宿主停在设计期 569x441、重叠在左上角 → 停靠区完全不成形。
//
// 这里为**能安全装箱的标量属性**各发一对 thunk, 再发一张 {名字, get, set} 表交给
// vb6_UserControlDesc.props; 宿主模型在内建成员未命中时查它 (见 vb6_UcPropDesc 注释)。
//
// ⚠ 为什么落在 emitDesignerControlDecls 而不是 emitFormFramework:
//   cgen_base_generate_decl_pass.inc:39 是 `if (module.isFormModule && frmDesc)
//   emitFormFramework(...); else if (frmDesc) emitDesignerControlDecls(*frmDesc);`
//   —— **.ctl 根本不走 emitFormFramework**, 只有这里的 desc 是 .ctl 也跑的。
//
// 判据只用符号表 (moduleScope 的 PropertyGet/PropertyLet/PropertySet):
//   - 类型只认 Vb6Type 的 Boolean/Integer/Long/LongPtr/Byte/Single/Double/Date/String。
//     Object/工程类指针/UDT/数组/Variant 一律**不发** —— 裸结构体指针当 VT_DISPATCH 会被
//     VariantClear 把首字段 (`__comObj`) 当 vtable 解引用 (见 cgen_setlet Fix 179a 一族),
//     UDT/数组/Variant 各有所有权与落临时语义。缺席即保持改动前行为 (返回 Empty), 不会更糟。
//   - 只发 0 参 Property Get 与 1 参 ByVal Property Let/Set; 带参默认属性与 ByRef 形参不发。
//
// rev35: 对象型 setter 改走 vb6_VariantToObjectVal 拆箱(详见下方 Fix <vbeclipse> rev35)。
//       ⚠ 它修的是「桥不该裸读 pdispVal」, 已实测让 Add 的六个属性全部写通, 但 Icon
//         崩溃的最终根因在**调用点侧**的实参打包, 尚未定位 —— 见下方说明。

{
    struct UcPropPick { std::string name; std::string vbType; std::string getFn; std::string setFn; std::string objCType; };
    std::vector<UcPropPick> ucProps;
    ucHostPropCount_ = 0;

    // Fix <vbeclipse>: 对象属性桥 (`.Folder = Folder` 这类 Property Set/Let 收
    // 工程类或 Object)。标量档 (vbScalarOf*) 认不出的, 若声明类型解析成指针类
    // C 类型 (vb6_cls_X* 工程类 / void* 通用 Object), 就按对象属性发**只写**桥:
    // setter 从 VARIANT 的 pdispVal 取回裸实例指针交给原生 prop_set_X(me, <cls>**),
    // 不 AddRef/不 VariantClear (借用, 所有权归调用方), 避开 rev18 注释警告的
    // "裸结构体指针当 VT_DISPATCH → VariantClear 把 __comObj 当 vtable" 崩溃族。
    auto objCTypeOf = [this](const std::string& tn, Vb6Type t) -> std::string {
        std::string ct = tn.empty() ? std::string() : cTypeForDeclaredTypeName(tn);
        if (!ct.empty() && ct.rfind("vb6_cls_", 0) == 0) return ct;   // 工程类实例指针
        if (t == Vb6Type::Object) return "void*";                     // 通用 Object 槽
        return std::string();
    };

    auto vbScalarOfName = [](const std::string& n) -> std::string {
        const std::string l = Symbol::toLower(n);
        if (l == "boolean" || l == "integer" || l == "long" || l == "longptr"
            || l == "byte" || l == "single" || l == "double" || l == "date"
            || l == "string") return l;
        return std::string();
    };
    auto vbScalarOfType = [](Vb6Type t) -> std::string {
        switch (t) {
            case Vb6Type::Boolean: return "boolean";
            case Vb6Type::Integer: return "integer";
            case Vb6Type::Long:    return "long";
            case Vb6Type::LongPtr: return "longptr";
            case Vb6Type::Byte:    return "byte";
            case Vb6Type::Single:  return "single";
            case Vb6Type::Double:  return "double";
            case Vb6Type::Date:    return "date";
            case Vb6Type::String:  return "string";
            default: return std::string();
        }
    };
    auto boxUcGetter = [](const std::string& vb, const std::string& call) -> std::string {
        if (vb == "boolean") return "vb6_VariantBool((int16_t)(" + call + "))";
        if (vb == "integer") return "vb6_VariantLong((int32_t)(" + call + "))";
        if (vb == "long" || vb == "longptr") return "vb6_VariantLong((int32_t)(" + call + "))";
        if (vb == "byte")    return "vb6_VariantByte((uint8_t)(" + call + "))";
        if (vb == "single")  return "vb6_VariantSingle((float)(" + call + "))";
        if (vb == "double" || vb == "date") return "vb6_VariantDouble((double)(" + call + "))";
        if (vb == "string")  return "vb6_VariantString(" + call + ")";
        return std::string();
    };
    auto unboxUcSetter = [](const std::string& vb, const std::string& V) -> std::string {
        if (vb == "boolean") return "(int16_t)vb6_VariantToBool(" + V + ")";
        if (vb == "integer") return "(int16_t)vb6_VariantToLong(" + V + ")";
        if (vb == "long" || vb == "longptr") return "(int32_t)vb6_VariantToLong(" + V + ")";
        if (vb == "byte")    return "(uint8_t)vb6_VariantToLong(" + V + ")";
        if (vb == "single")  return "(float)vb6_VariantToDouble(" + V + ")";
        if (vb == "double" || vb == "date") return "vb6_VariantToDouble(" + V + ")";
        if (vb == "string")  return "vb6_VariantToString(" + V + ")";
        return std::string();
    };

    if (symTab_.moduleScope()) {
        for (auto& kv : symTab_.moduleScope()->symbols()) {
            Symbol* ps = kv.second.get();
            if (!ps || ps->access != AccessLevel::Public) continue;
            // ⚠ 只收**本模块自有**的属性。moduleScope 里同时躺着 driver_crossmod 注入的
            // 外部副本 (isExternal), 不收它们会为别的模块的属性发 thunk → LNK2019
            // "无法解析的外部符号 _vb6_<本模块>_prop_get_<别人的属性>" (实测 ucFolder 里
            // 冒出 ActiveCaptionForeColor/Angle/ActiveViewId 等 ucSplitBar/IScheme 的属性)。
            if (ps->isExternal) continue;
            UcPropPick pick;
            pick.name = ps->name;
            if (ps->kind == SymbolKind::PropertyGet) {
                if (!ps->params.empty()) continue;              // 带参默认属性不走桥
                pick.vbType = vbScalarOfName(ps->variableTypeName);
                if (pick.vbType.empty()) pick.vbType = vbScalarOfType(ps->type);
                if (pick.vbType.empty()) {
                    // Fix <vbeclipse> rev25: **对象返回型 Property Get 也要发桥**。
                    // __comObj 当 vtable"正是 `vb6_ComObject_FromInstance` 要解决的问题:
                    // 它包成真 IDispatch 再转 VARIANT, 引用计数随 VARIANT 走。
                    //
                    // 实证 play78 (--arch x86): `ucTabStrip.ctl:232`
                    //     Public Property Get Tabs() As List
                    //        Set Tabs = m_Tabs
                    // 是对象返回属性, 当年没桥 ⇒ `vb6_ComGetObjectProp(ViewTabs, "Tabs")`
                    // 返 NULL ⇒ `ucFolder.ContainsView` 里 `vb6_List_Count(NULL)` 得到 -98
                    // ⇒ **对任何 ViewId 都返 False**。
                    // 后果: `ucPerspective.ShowView` 的"这个 folder 含该视图吗"判定全不中
                    // ⇒ 落到 `ucPerspective.ctl:2273` 的 not-visible 分支
                    // ⇒ 把视图塞进 `Folders.Item(Folders.Count)`(= 最后一个 folder)
                    // ⇒ `frmViewViews`/`frmViewEvents` 抢了 `Left_Folder` 的位置。
                    pick.objCType = objCTypeOf(ps->variableTypeName, ps->type);
                    // 对象 **getter** 只认工程类 (`vb6_cls_*`): 通用 `void*` 槽拿不到
                    // coclass 名, 后面 FindCoClassDesc 会拿空串去查 ⇒ NULL。
                    if (pick.objCType.rfind("vb6_cls_", 0) != 0) pick.objCType.clear();
                    if (pick.objCType.empty()) continue;
                    pick.getFn = "vb6_" + ctl + "_ucPropGetObj_" + cIdent(ps->name);
                } else {
                    pick.getFn = "vb6_" + ctl + "_ucPropGet_" + cIdent(ps->name);
                }
            } else if (ps->kind == SymbolKind::PropertyLet || ps->kind == SymbolKind::PropertySet) {
                if (ps->params.size() != 1) continue;
                const ParameterInfo& pi = ps->params[0];
                pick.vbType = vbScalarOfName(pi.typeRefName);
                if (pick.vbType.empty()) pick.vbType = vbScalarOfType(pi.type);
                if (!pick.vbType.empty()) {
                    if (!pi.isByVal) continue;                  // 标量 ByRef 形参要落临时, 不作为
                } else {
                    // 标量不中 → 试对象档 (工程类 / Object)。对象 Property Set 形参在 VB6
                    // 天然 ByRef (C 侧 <cls>**), 但桥 setter 用本地 _o 取址传下去即"落临时",
                    // 不受 isByVal 限制。只发 setter, getter 留 NULL (对象 getter 装箱有
                    // VariantClear 把 __comObj 当 vtable 的崩溃风险, 见 objCTypeOf 注释)。
                    pick.objCType = objCTypeOf(pi.typeRefName, pi.type);
                    if (pick.objCType.empty()) continue;
                }
                pick.setFn = "vb6_" + ctl + "_ucPropSet_" + cIdent(ps->name);
            } else {
                continue;
            }
            ucProps.push_back(std::move(pick));
        }
    }
    // Get 与 Let 是两条符号 (键分别为 name$pg / name$pl) → 合并成表里一行
    {
        std::vector<UcPropPick> merged;
        for (auto& p : ucProps) {
            bool hit = false;
            for (auto& m : merged) {
                if (Symbol::toLower(m.name) == Symbol::toLower(p.name)) {
                    if (!p.getFn.empty()) m.getFn = p.getFn;
                    if (!p.setFn.empty()) m.setFn = p.setFn;
                    if (!p.vbType.empty()) m.vbType = p.vbType;
                    if (!p.objCType.empty()) m.objCType = p.objCType;
                    hit = true;
                    break;
                }
            }
            if (!hit) merged.push_back(p);
        }
        ucProps.swap(merged);
    }
    // 按名字排序: 符号表是 unordered_map, 遍历顺序不保证——排序让产物稳定可比
    std::sort(ucProps.begin(), ucProps.end(),
              [](const UcPropPick& a, const UcPropPick& b) {
                  return Symbol::toLower(a.name) < Symbol::toLower(b.name);
              });

    for (auto& p : ucProps) {
        if (!p.getFn.empty()) {
            // ⚠ 只在 getter **自己**是对象返回时才走对象 thunk。若 getter 是标量而
            // setter 是对象 (合并后 objCType 非空), 这里必须走标量分支 ——
            // 否则会用 `vb6_ComObject_FromInstance` 去包一个 long。
            bool objGet = !p.objCType.empty() && p.vbType.empty();
            if (objGet) {
                // Fix <vbeclipse> rev25: **对象返回型 getter**。原生 prop_get_X 返回
                // `vb6_cls_List*` 之类的裸实例指针, 直接塞 VARIANT 会让 VariantClear
                // 把首字段 (`__comObj`) 当 vtable 解引用 —— 必须经 vb6_ComObject_FromInstance
                // 包成真 IDispatch, 引用计数随 VARIANT 走 (那个函数的头文件就是为这个存在的)。
                c_.emitLine("static void* " + p.getFn + "(void* inst, void* outV) {");
                c_.emitLine("    " + p.objCType + " _o = vb6_" + ctl + "_prop_get_"
                            + cIdent(p.name) + "((vb6_cls_" + ctl + "*)inst);");
                c_.emitLine("    if (!_o) { *(vb6_VARIANT*)outV = vb6_VariantEmpty(); return outV; }");
                c_.emitLine("    vb6_VARIANT* _pv = (vb6_VARIANT*)outV;");
                c_.emitLine("    _pv->vt = VT_DISPATCH;");
                // FromInstance 返回的引用直接交给这个 VARIANT (引用计数转移, 不再 AddRef),
                // 调用方 VariantClear 时正好归零 —— 与 vb6_ComPackVB6InstanceRaw 同款取舍。
                // 类名从 "vb6_cls_List*" 里切出 "List": 去掉前缀 "vb6_cls_" 与尾 '*'。
                // ⚠ 别用 substr(6) —— 短于 6 的 C 类型会抛 std::out_of_range,
                //   实测直接把 C3 打进 ICE (invalid string position + 崩溃)。
                c_.emitLine("    _pv->pdispVal = (IDispatch*)vb6_ComObject_FromInstance(");
                c_.emitLine("        vb6_FindCoClassDesc(\""
                            + p.objCType.substr(std::string("vb6_cls_").size(),
                                                p.objCType.size()
                                                - std::string("vb6_cls_").size() - 1)
                            + "\"), (void*)_o);");
                c_.emitLine("    return outV;");
                c_.emitLine("}");
            } else {
            c_.emitLine("static void* " + p.getFn + "(void* inst, void* outV) {");
            c_.emitLine("    *(vb6_VARIANT*)outV = "
                        + boxUcGetter(p.vbType, "vb6_" + ctl + "_prop_get_" + cIdent(p.name)
                                                + "((vb6_cls_" + ctl + "*)inst)") + ";");
            c_.emitLine("    return outV;");
            c_.emitLine("}");
            }
        }
        if (!p.setFn.empty()) {
            // 同上: 只有 setter **自己**是对象形参时才走对象分支 (vbType 空 = 标量档不中)。
            if (!p.objCType.empty() && p.vbType.empty()) {
                // Fix <vbeclipse> rev35: 借 vb6_VariantToObjectVal 拆箱, **不裸读 pdispVal**。
                //   原症状: `ucTabStrip.Add(…, Optional Icon)` 的 `.Icon = Icon` 崩在
                //   vb6_SetControlPictureFromCom 内的 QueryInterface (0xC0000005)。
                //   旧代码无条件 `(inV)->pdispVal`, **不判 vt** —— vt 不是 Dispatch 时读到的
                //   是 union 里另一成员(栈上残留), 当 IDispatch 传下去必崩。
                //   与 cgen_form_uc_methods.inc 的 rev34 同款修法; 原生 prop_set_X 侧的
                //   `!pPictureDisp` 分支负责"清旧图", 语义与 VB6 一致。
                //
                // ⚠ 实测边界(**别扩大结论**): rev35 让 Add 的六个属性**全部写通**(探针实测
                //   Orientation→ToolTip→ViewId→Icon, Caption/Visible 走宿主内建成员表),
                //   但进程**仍** 0xC0000005, 崩点后移到 SetControlPictureFromCom 内部 ——
                //   且此时 `pdispVal` 是**非 NULL 的野值**(实测 010FEE6C), 不是 NULL。
                //   ⇒ 本条修的是"桥不该裸读 union", **不是**"Icon 崩溃的最终根因";
                //   野值来自**调用点侧**对 `Optional Icon As Picture`(省略实参)的打包,
                //   那才是下一步要查的地方。别因为 rev35 在链上就认为 Icon 已经修好。
                c_.emitLine("static int32_t " + p.setFn + "(void* inst, const void* inV) {");
                c_.emitLine("    " + p.objCType + " _o = (" + p.objCType
                            + ")vb6_VariantToObjectVal(*(const vb6_VARIANT*)inV);");
                c_.emitLine("    vb6_" + ctl + "_prop_set_" + cIdent(p.name) + "((vb6_cls_" + ctl
                            + "*)inst, &_o);");
                c_.emitLine("    return 1;");
                c_.emitLine("}");
            } else {
            c_.emitLine("static int32_t " + p.setFn + "(void* inst, const void* inV) {");
            c_.emitLine("    vb6_" + ctl + "_prop_let_" + cIdent(p.name) + "((vb6_cls_" + ctl
                        + "*)inst, "
                        + unboxUcSetter(p.vbType, "(*(const vb6_VARIANT*)inV)") + ");");
            c_.emitLine("    return 1;");
            c_.emitLine("}");
            }
        }
    }
    if (!ucProps.empty()) {
        c_.emitLine("static const vb6_UcPropDesc vb6_" + ctl + "_ucProps[] = {");
        for (auto& p : ucProps) {
            c_.emitLine("    { L\"" + p.name + "\", "
                        + (p.getFn.empty() ? std::string("NULL") : p.getFn) + ", "
                        + (p.setFn.empty() ? std::string("NULL") : p.setFn) + " },");
        }
        c_.emitLine("};");
    }
    ucHostPropCount_ = static_cast<int32_t>(ucProps.size());
}
