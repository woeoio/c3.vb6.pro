// VB6 TypeLib解析器实现 - P6.3 前期绑定支持
// 编译期使用Windows LoadTypeLib/ITypeInfo API

#include "com/typelib_parser.hpp"
#include <algorithm>
#include <windows.h>
#include <oleauto.h>

namespace vb6c3 {

// --- typelib_parser_parse.cpp: ITypeInfo 解析（TypeLib / Interface / CoClass / Module / Enum） ---

// ============================================================
// 内部: 解析TypeLib
// ============================================================

bool TypeLibParser::parseTypeLib(void* pTypeLib, TypeLibResult& result) {
    ITypeLib* pTL = static_cast<ITypeLib*>(pTypeLib);

    // 获取TypeLib属性
    TLIBATTR* pTlbAttr = nullptr;
    HRESULT hr = pTL->GetLibAttr(&pTlbAttr);
    if (SUCCEEDED(hr) && pTlbAttr) {
        // 版本
        result.version = std::to_string(pTlbAttr->wMajorVerNum) + "." +
                         std::to_string(pTlbAttr->wMinorVerNum);
        pTL->ReleaseTLibAttr(pTlbAttr);
    }

    // 获取名称
    BSTR tlbName = nullptr;
    pTL->GetDocumentation(-1, &tlbName, nullptr, nullptr, nullptr);
    if (tlbName) {
        // 宽字符→窄字符
        result.name.clear();
        for (UINT i = 0; i < SysStringLen(tlbName); i++) {
            result.name += (char)tlbName[i];
        }
        SysFreeString(tlbName);
    }

    // P24-04: 捕获TypeLib项目名 (VB6 ActiveX DLL的Name=属性, 如"VBMANLIB")
    result.typeLibProjectName = result.name;
    // Fix 098: 供 parseCoClass 在 ProgIDFromCLSID 反查失败时拼 ProgID 前缀
    currentProjectName_ = result.name;

    // 枚举所有类型
    UINT count = pTL->GetTypeInfoCount();
    for (UINT i = 0; i < count; i++) {
        TYPEKIND typeKind;
        hr = pTL->GetTypeInfoType(i, &typeKind);
        if (FAILED(hr)) continue;

        ITypeInfo* pTI = nullptr;
        hr = pTL->GetTypeInfo(i, &pTI);
        if (FAILED(hr) || !pTI) continue;

        // 获取类型名
        BSTR typeBStr = nullptr;
        pTI->GetDocumentation(MEMBERID_NIL, &typeBStr, nullptr, nullptr, nullptr);
        std::string typeName;
        if (typeBStr) {
            for (UINT j = 0; j < SysStringLen(typeBStr); j++) {
                typeName += (char)typeBStr[j];
            }
            SysFreeString(typeBStr);
        }

        switch (typeKind) {
            case TKIND_INTERFACE:
            case TKIND_DISPATCH: {
                auto iface = parseInterface(pTI, typeName);
                if (iface) {
                    result.interfaces.push_back(std::move(iface));
                }
                break;
            }
            case TKIND_COCLASS: {
                auto coclass = parseCoClass(pTI, typeName);
                if (coclass) {
                    result.coclasses.push_back(std::move(coclass));
                }
                break;
            }
            case TKIND_MODULE: {
                // P24-04: 解析ActiveX DLL全局模块函数
                auto mod = parseModule(pTI, typeName, result.tlbPath);
                if (mod) {
                    result.modules.push_back(std::move(mod));
                }
                break;
            }
            case TKIND_ENUM: {
                // Fix 018: 解析COM枚举类型 (TextCompare/adStateClosed 等命名常量)
                auto en = parseEnum(pTI, typeName);
                if (en) {
                    result.enums.push_back(std::move(en));
                }
                break;
            }
            case TKIND_ALIAS:
            case TKIND_RECORD:
            default:
                break;
        }

        pTI->Release();
    }

    // P24-04: 链接coclass → default interface (从 loadByPath 移至此处)
    // 必须在GlobalNameSpace检测之前完成, 因为检测依赖 defaultIface 指针
    for (auto& cc : result.coclasses) {
        if (!cc->defaultIfaceName.empty()) {
            cc->defaultIface = result.findInterface(cc->defaultIfaceName);
            // P13.20: Link default source (event) interface
            if (!cc->defaultSourceIfaceName.empty()) {
                cc->defaultSourceIface = result.findInterface(cc->defaultSourceIfaceName);
            }
        }
    }

    // P24-04: GlobalNameSpace检测后处理
    // VB_GlobalNameSpace=True的coclass, 其默认接口的Public方法提升为全局符号
    // 检测启发式: TYPEFLAG_FPREDECLID(0x0008) + 默认接口有方法, 且满足以下之一:
    //   (a) coclass名包含"Global" (如sGlobal)
    //   (b) 默认接口有方法名与TypeLib项目名匹配 (如VBMAN库的VBMAN()方法)
    // 对于每个检测到的GlobalNameSpace coclass, 设置isGlobalNamespace=true
    // 并将默认接口的custom方法(排除IUnknown/IDispatch标准方法)记录为promoted methods
    if (!result.typeLibProjectName.empty()) {
        std::string tlbNameLower = result.typeLibProjectName;
        std::transform(tlbNameLower.begin(), tlbNameLower.end(), tlbNameLower.begin(), ::tolower);
        for (auto& cc : result.coclasses) {
            // 初步筛选: coclass名含"Global" (不区分大小写)
            std::string ccNameLower = cc->name;
            std::transform(ccNameLower.begin(), ccNameLower.end(), ccNameLower.begin(), ::tolower);
            bool nameHasGlobal = ccNameLower.find("global") != std::string::npos;
            // 初步筛选: 默认接口有与TypeLib名同名的方法
            bool hasMatchingMethod = false;
            if (cc->defaultIface) {
                for (auto& m : cc->defaultIface->members) {
                    std::string mLower = m.realName;
                    std::transform(mLower.begin(), mLower.end(), mLower.begin(), ::tolower);
                    if (mLower == tlbNameLower) {
                        hasMatchingMethod = true;
                        break;
                    }
                }
            }
            if ((nameHasGlobal || hasMatchingMethod) && cc->defaultIface) {
                cc->isGlobalNamespace = true;
            }
        }
    }

    return !result.interfaces.empty() || !result.coclasses.empty()
        || !result.modules.empty() || !result.enums.empty();
}

// ============================================================
// 内部: 解析接口
// ============================================================

std::unique_ptr<ComInterfaceInfo> TypeLibParser::parseInterface(void* pTypeInfo,
                                                                 const std::string& name) {
    ITypeInfo* pTI = static_cast<ITypeInfo*>(pTypeInfo);
    auto iface = std::make_unique<ComInterfaceInfo>();
    iface->name = name;

    // 获取TYPEATTR
    TYPEATTR* pTypeAttr = nullptr;
    HRESULT hr = pTI->GetTypeAttr(&pTypeAttr);
    if (FAILED(hr) || !pTypeAttr) return nullptr;

    // IID
    iface->iidStr = iidToString((const uint8_t*)&pTypeAttr->guid);

    // 接口类型
    if (pTypeAttr->typekind == TKIND_DISPATCH) {
        iface->isDispatch = true;
    }
    if (pTypeAttr->wTypeFlags & TYPEFLAG_FDUAL) {
        iface->isDual = true;
    }

    // vtable起始偏移 (IUnknown=3, IDispatch=4+3=7)
    int vtableBase = 7;  // 对于dispinterface, vtable从第7个槽开始
    if (pTypeAttr->typekind == TKIND_INTERFACE) {
        // 纯IUnknown派生, vtable从第3个槽开始
        // 但如果有基接口, 需要调整
        if (pTypeAttr->cImplTypes > 0) {
            // 假设基接口是IUnknown, 3个方法
            vtableBase = 3;
        }
    }

    // 保存需要的值后释放
    UINT cFuncs = pTypeAttr->cFuncs;
    UINT cVars = pTypeAttr->cVars;
    TYPEKIND typeKindSaved = pTypeAttr->typekind;
    pTI->ReleaseTypeAttr(pTypeAttr);

    // 枚举方法/属性
    for (UINT i = 0; i < cFuncs; i++) {
        FUNCDESC* pFuncDesc = nullptr;
        hr = pTI->GetFuncDesc(i, &pFuncDesc);
        if (FAILED(hr) || !pFuncDesc) continue;

        ComMemberInfo member = parseFuncDesc(pTI, pFuncDesc, i);

        // vtable索引: 统一使用FUNCDESC.oVft / sizeof(void*)
        // 注意: 对于dispinterface/dual, 枚举序号i不等于vtable位置,
        // oVft才是正确的vtable字节偏移 (包含IUnknown+IDispatch槽位)
        if (pFuncDesc->oVft > 0) {
            member.vtableIndex = pFuncDesc->oVft / (int)sizeof(void*);
        } else {
            // 兜底: 对于没有oVft的接口, 使用计算值
            member.vtableIndex = vtableBase + (int)i;
        }

        iface->members.push_back(std::move(member));
        pTI->ReleaseFuncDesc(pFuncDesc);
    }

    // 枚举变量(属性)
    for (UINT i = 0; i < cVars; i++) {
        VARDESC* pVarDesc = nullptr;
        hr = pTI->GetVarDesc(i, &pVarDesc);
        if (FAILED(hr) || !pVarDesc) continue;

        ComMemberInfo member = parseVarDesc(pTI, pVarDesc);
        iface->members.push_back(std::move(member));
        pTI->ReleaseVarDesc(pVarDesc);
    }

    // P24-10: 检测默认成员 (DISPID_VALUE=0)
    // VB6语义: obj(args) 等价于 obj.DefaultMember(args)
    // COM中DISPID=0标识默认成员, 如Dictionary.Item, Collection._Item
    for (auto& m : iface->members) {
        if (m.memid == 0 /* DISPID_VALUE */) {
            iface->defaultMemberName = m.name;         // 小写
            iface->defaultMemberRealName = m.realName; // 原始大小写
            break;
        }
    }

    return iface;
}

// ============================================================
// 内部: 解析coclass
// ============================================================

std::unique_ptr<ComCoClassInfo> TypeLibParser::parseCoClass(void* pTypeInfo,
                                                             const std::string& name) {
    ITypeInfo* pTI = static_cast<ITypeInfo*>(pTypeInfo);
    auto cc = std::make_unique<ComCoClassInfo>();
    cc->name = name;

    // 获取TYPEATTR
    TYPEATTR* pTypeAttr = nullptr;
    HRESULT hr = pTI->GetTypeAttr(&pTypeAttr);
    if (FAILED(hr) || !pTypeAttr) return nullptr;

    cc->clsidStr = iidToString((const uint8_t*)&pTypeAttr->guid);

    // VB6 全局对象 (appobject): TYPEFLAG_FAPPOBJECT —— 裸类名即进程级单例
    cc->isAppObject = (pTypeAttr->wTypeFlags & TYPEFLAG_FAPPOBJECT) != 0;

    // 查找默认接口 (IMPLTYPEFLAG_FDEFAULT)
    for (UINT i = 0; i < pTypeAttr->cImplTypes; i++) {
        INT implFlags = 0;
        HREFTYPE refType = 0;
        hr = pTI->GetImplTypeFlags(i, &implFlags);
        if (FAILED(hr)) continue;

        // 获取实现的接口类型
        hr = pTI->GetRefTypeOfImplType(i, &refType);
        if (FAILED(hr)) continue;

        ITypeInfo* pImplTI = nullptr;
        hr = pTI->GetRefTypeInfo(refType, &pImplTI);
        if (FAILED(hr) || !pImplTI) continue;

        BSTR implName = nullptr;
        pImplTI->GetDocumentation(MEMBERID_NIL, &implName, nullptr, nullptr, nullptr);
        if (implName) {
            std::string ifaceName;
            for (UINT j = 0; j < SysStringLen(implName); j++) {
                ifaceName += (char)implName[j];
            }
            SysFreeString(implName);

                        // P13.20: Event source interface recognition
            bool isSource = (implFlags & IMPLTYPEFLAG_FSOURCE) != 0;
            bool isDefault = (implFlags & IMPLTYPEFLAG_FDEFAULT) != 0;
            
            if (isSource) {
                // Event source interface
                cc->sourceIfaceNames.push_back(ifaceName);
                if (isDefault || cc->defaultSourceIfaceName.empty()) {
                    cc->defaultSourceIfaceName = ifaceName;
                }
            } else {
                // Regular (outgoing) interface
                if (isDefault) {
                    cc->defaultIfaceName = ifaceName;
                }
                // If no FDEFAULT flag, take first non-source interface
                if (cc->defaultIfaceName.empty()) {
                    cc->defaultIfaceName = ifaceName;
                }
            }
        }  // end if (implName)
        pImplTI->Release();
    }

    // ProgID: 三级解析 (Fix 160)
    //   ① 类型库二进制里的 coclass `progid` 属性 — 免注册, 权威值
    //   ② 注册表反查 ProgIDFromCLSID — 已注册场景
    //   ③ <TypeLib库名>.<coclass名> 合成兜底 (Fix 098)
    //
    // 旧实现只有 ②→③ 两级。② 依赖注册表, 而「组件未注册」和「32 位组件 + 64 位宿主」
    // 下反查必失败 → 落 ③, 而 ③ 对真实组件已知是错的 (cgen_form_create_controls.inc
    // 记录 MSComctlLib.ImageList ≠ ImageListCtrl)。① 在 typelib 二进制里, 与注册表
    // 无关, 补齐后免注册场景也能拿到真 ProgID。
    {
        // ① coclass 类型属性: MIDL 把 coclass 的 `progid` 属性编码为类型库自定义属性,
        //   其 Guid 为 {B54F3741-5B07-11CF-A4B0-00AA004A55E8}。通过 ITypeInfo2::GetCustData
        //   读取 (返回 VT_BSTR)。1.x 版本类型库无自定义属性 → QI 失败或取空, 落到 ②/③。
        std::string progIdAttr;
        ITypeInfo2* pTI2 = nullptr;
        if (SUCCEEDED(pTI->QueryInterface(IID_ITypeInfo2, (void**)&pTI2)) && pTI2) {
            static const GUID kProgIdAttrGuid = { 0xB54F3741, 0x5B07, 0x11CF,
                                                  { 0xA4, 0xB0, 0x00, 0xAA, 0x00, 0x4A, 0x55, 0xE8 } };
            VARIANT val;
            VariantInit(&val);
            if (SUCCEEDED(pTI2->GetCustData(kProgIdAttrGuid, &val)) && val.vt == VT_BSTR && val.bstrVal) {
                for (ULONG j = 0; j < SysStringLen(val.bstrVal); j++) progIdAttr += (char)val.bstrVal[j];
            }
            VariantClear(&val);
            pTI2->Release();
        }

        // ② 注册表反查
        CLSID clsid = pTypeAttr->guid;
        LPOLESTR progIdW = nullptr;
        HRESULT progHr = ProgIDFromCLSID(clsid, &progIdW);
        std::string regProgId;
        if (SUCCEEDED(progHr) && progIdW) {
            for (const wchar_t* p = progIdW; *p; p++) {
                regProgId += (char)*p;
            }
            CoTaskMemFree(progIdW);
        }

        if (!progIdAttr.empty()) {
            cc->progId = progIdAttr;
        } else if (!regProgId.empty()) {
            cc->progId = regProgId;
        } else {
            // Fix 098: 反查失败时不能退化成裸类名 —— VB6 约定 ProgID = <TypeLib库名>.<coclass名>
            // (实测: VBMAN.dll→"VBMANLIB", stdole2.tlb→"stdole", msado28.tlb→"ADODB", 全部符合).
            // 裸类名会让运行期 CLSIDFromProgID 查不到 → CreateObject 报 429 / 0x800401F3.
            // 反查失败在「组件未注册」或「32 位组件 + 64 位宿主进程」下是常态, 属预期路径.
            cc->progId = currentProjectName_.empty()
                ? name
                : (currentProjectName_ + "." + name);
        }
    }

    pTI->ReleaseTypeAttr(pTypeAttr);

    return cc;
}


// ============================================================
// P24-04: 内部: 解析模块 (TKIND_MODULE → ActiveX DLL全局函数)
// ============================================================

std::unique_ptr<ComModuleInfo> TypeLibParser::parseModule(void* pTypeInfo,
                                                           const std::string& name,
                                                           const std::string& dllPath) {
    ITypeInfo* pTI = static_cast<ITypeInfo*>(pTypeInfo);
    auto mod = std::make_unique<ComModuleInfo>();
    mod->name = name;

    // 推断DLL路径: TypeLib通常嵌在DLL中, 用tlbPath的目录+DLL名推断
    // 如果tlbPath本身是DLL文件, 直接使用
    if (!dllPath.empty()) {
        // tlbPath可能是 "C:\path\VBMAN.dll" 或注册表中的tlb路径
        // 对ActiveX DLL, TypeLib就嵌在DLL中
        if (dllPath.size() >= 4 &&
            (dllPath.substr(dllPath.size()-4) == ".dll" ||
             dllPath.substr(dllPath.size()-4) == ".DLL" ||
             dllPath.substr(dllPath.size()-4) == ".ocx" ||
             dllPath.substr(dllPath.size()-4) == ".OCX")) {
            mod->dllPath = dllPath;
        }
        // 否则不设dllPath, 运行时通过注册表查找
    }

    // 获取TYPEATTR
    TYPEATTR* pTypeAttr = nullptr;
    HRESULT hr = pTI->GetTypeAttr(&pTypeAttr);
    if (FAILED(hr) || !pTypeAttr) return nullptr;

    UINT cFuncs = pTypeAttr->cFuncs;
    UINT cVars = pTypeAttr->cVars;

    pTI->ReleaseTypeAttr(pTypeAttr);

    // 枚举函数 (模块级Public函数)
    for (UINT i = 0; i < cFuncs; i++) {
        FUNCDESC* pFuncDesc = nullptr;
        hr = pTI->GetFuncDesc(i, &pFuncDesc);
        if (FAILED(hr) || !pFuncDesc) continue;

        ComMemberInfo member = parseFuncDesc(pTI, pFuncDesc, i);
        // 模块函数无vtable概念
        member.vtableIndex = -1;
        mod->functions.push_back(std::move(member));
        pTI->ReleaseFuncDesc(pFuncDesc);
    }

    // 枚举常量 (模块级Public Const)
    for (UINT i = 0; i < cVars; i++) {
        VARDESC* pVarDesc = nullptr;
        hr = pTI->GetVarDesc(i, &pVarDesc);
        if (FAILED(hr) || !pVarDesc) continue;

        ComMemberInfo member = parseVarDesc(pTI, pVarDesc);
        mod->constants.push_back(std::move(member));
        pTI->ReleaseVarDesc(pVarDesc);
    }

    // 如果没有函数也没有常量, 仍保留模块(可能仅作为命名空间)
    return mod;
}

// ============================================================
// Fix 018: 内部: 解析枚举 (TKIND_ENUM → 枚举成员命名常量)
// ============================================================

std::unique_ptr<ComEnumInfo> TypeLibParser::parseEnum(void* pTypeInfo,
                                                       const std::string& name) {
    ITypeInfo* pTI = static_cast<ITypeInfo*>(pTypeInfo);
    auto en = std::make_unique<ComEnumInfo>();
    en->name = name;
    en->lowerName = name;
    std::transform(en->lowerName.begin(), en->lowerName.end(),
                   en->lowerName.begin(), ::tolower);

    // 获取TYPEATTR 得到 cVars (枚举成员数)
    TYPEATTR* pTypeAttr = nullptr;
    HRESULT hr = pTI->GetTypeAttr(&pTypeAttr);
    if (FAILED(hr) || !pTypeAttr) return nullptr;

    UINT cVars = pTypeAttr->cVars;
    pTI->ReleaseTypeAttr(pTypeAttr);

    // 枚举每个成员 (VARDESC, varkind==VAR_CONST, 值在 lpvarValue)
    for (UINT i = 0; i < cVars; i++) {
        VARDESC* pVarDesc = nullptr;
        hr = pTI->GetVarDesc(i, &pVarDesc);
        if (FAILED(hr) || !pVarDesc) continue;

        ComEnumMemberInfo m;

        // 成员名 (通过 memid 取文档名)
        BSTR varName = nullptr;
        pTI->GetDocumentation(pVarDesc->memid, &varName, nullptr, nullptr, nullptr);
        if (varName) {
            for (UINT j = 0; j < SysStringLen(varName); j++) {
                m.name += (char)varName[j];
            }
            SysFreeString(varName);
        }
        m.lowerName = m.name;
        std::transform(m.lowerName.begin(), m.lowerName.end(),
                       m.lowerName.begin(), ::tolower);

        // 枚举值: varkind==VAR_CONST 时 lpvarValue 指向 VARIANT
        if (pVarDesc->varkind == VAR_CONST && pVarDesc->lpvarValue) {
            VARIANT* pVar = pVarDesc->lpvarValue;
            VARTYPE vt = pVar->vt & VT_TYPEMASK;
            switch (vt) {
                case VT_I1:   m.value = pVar->cVal; break;
                case VT_UI1:  m.value = pVar->bVal; break;
                case VT_I2:   m.value = pVar->iVal; break;
                case VT_UI2:  m.value = pVar->uiVal; break;
                case VT_I4:   m.value = pVar->lVal; break;
                case VT_UI4:  m.value = (int64_t)pVar->ulVal; break;
                case VT_BOOL: m.value = (pVar->boolVal != 0) ? -1 : 0; break;  // VB6 True=-1
                default:      m.value = pVar->lVal; break;  // 回退到 Long
            }
        }

        if (!m.name.empty()) {
            en->members.push_back(std::move(m));
        }
        pTI->ReleaseVarDesc(pVarDesc);
    }

    return en;
}

} // namespace vb6c3
