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

{
    struct UcPropPick { std::string name; std::string vbType; std::string getFn; std::string setFn; };
    std::vector<UcPropPick> ucProps;
    ucHostPropCount_ = 0;

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
                if (pick.vbType.empty()) continue;
                pick.getFn = "vb6_" + ctl + "_ucPropGet_" + cIdent(ps->name);
            } else if (ps->kind == SymbolKind::PropertyLet || ps->kind == SymbolKind::PropertySet) {
                if (ps->params.size() != 1) continue;
                const ParameterInfo& pi = ps->params[0];
                if (!pi.isByVal) continue;                      // ByRef 形参要落临时, 不作为
                pick.vbType = vbScalarOfName(pi.typeRefName);
                if (pick.vbType.empty()) pick.vbType = vbScalarOfType(pi.type);
                if (pick.vbType.empty()) continue;
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
            c_.emitLine("static void* " + p.getFn + "(void* inst, void* outV) {");
            c_.emitLine("    *(vb6_VARIANT*)outV = "
                        + boxUcGetter(p.vbType, "vb6_" + ctl + "_prop_get_" + cIdent(p.name)
                                                + "((vb6_cls_" + ctl + "*)inst)") + ";");
            c_.emitLine("    return outV;");
            c_.emitLine("}");
        }
        if (!p.setFn.empty()) {
            c_.emitLine("static int32_t " + p.setFn + "(void* inst, const void* inV) {");
            c_.emitLine("    vb6_" + ctl + "_prop_let_" + cIdent(p.name) + "((vb6_cls_" + ctl
                        + "*)inst, "
                        + unboxUcSetter(p.vbType, "(*(const vb6_VARIANT*)inV)") + ");");
            c_.emitLine("    return 1;");
            c_.emitLine("}");
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
