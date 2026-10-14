// VB6 TypeLib解析器实现 - P6.3 前期绑定支持
// 编译期使用Windows LoadTypeLib/ITypeInfo API

#include "com/typelib_parser.hpp"
#include <algorithm>
#include <windows.h>
#include <oleauto.h>

namespace vb6c3 {

// --- typelib_parser_desc.cpp: 描述符映射（FuncDesc / VarDesc / TypeDesc / ElemDesc） ---


// ============================================================
// 内部: 解析方法
// ============================================================

ComMemberInfo TypeLibParser::parseFuncDesc(void* pTypeInfo, void* pFuncDesc, int index) {
    ITypeInfo* pTI = static_cast<ITypeInfo*>(pTypeInfo);
    FUNCDESC* pFD = static_cast<FUNCDESC*>(pFuncDesc);

    ComMemberInfo member;
    member.memid = pFD->memid;

    // 获取名称
    BSTR funcName = nullptr;
    UINT cNames = 0;
    pTI->GetNames(pFD->memid, &funcName, 1, &cNames);
    if (funcName) {
        member.realName.clear();
        for (UINT i = 0; i < SysStringLen(funcName); i++) {
            member.realName += (char)funcName[i];
        }
        SysFreeString(funcName);
    }
    member.name = member.realName;
    std::transform(member.name.begin(), member.name.end(), member.name.begin(), ::tolower);

    // 成员类别
    switch (pFD->invkind) {
        case INVOKE_FUNC:
            member.kind = ComMemberKind::Method;
            break;
        case INVOKE_PROPERTYGET:
            member.kind = ComMemberKind::PropertyGet;
            break;
        case INVOKE_PROPERTYPUT:
            member.kind = ComMemberKind::PropertyPut;
            break;
        case INVOKE_PROPERTYPUTREF:
            member.kind = ComMemberKind::PropertyPutRef;
            break;
        default:
            member.kind = ComMemberKind::Method;
            break;
    }

    // 调用约定
    switch (pFD->callconv) {
        case CC_CDECL:   member.callConv = CallConv::CDecl;   break;
        case CC_STDCALL: member.callConv = CallConv::StdCall; break;
        default:         member.callConv = CallConv::StdCall; break;
    }

    // 返回类型
    // Fix-074: COM dispinterface 属性方法的返回类型在 pFD->elemdescFunc.tdesc 中
    // 可能是 VT_HRESULT (0x0019) 或 VT_VOID (0x0018)，表示"方法本身无返回值，
    // 实际返回值在最后一个 [out, retval] 参数中"。此时不应使用 tdesc 作为 returnType，
    // 而应等待参数处理阶段从 [out, retval] 参数中提取。
    VARTYPE funcVt = pFD->elemdescFunc.tdesc.vt;
    if (funcVt == VT_HRESULT || funcVt == VT_VOID) {
        // 方法的直接返回类型是 HRESULT/VOID，returnType 延后设置
        member.returnType = Vb6Type::Empty;  // 标记为待覆盖
    } else {
        member.returnType = mapTypeDesc(&pFD->elemdescFunc.tdesc, pTI);
    }

    // 参数
    // 获取参数名
    std::vector<std::string> paramNames;
    if (pFD->cParams > 0) {
        std::vector<BSTR> nameBuf(pFD->cParams + 1);
        UINT cNamesOut = 0;
        pTI->GetNames(pFD->memid, nameBuf.data(), pFD->cParams + 1, &cNamesOut);
        // nameBuf[0]是函数名, [1..cParams]是参数名
        for (UINT i = 1; i < cNamesOut && i <= pFD->cParams; i++) {
            std::string pname;
            if (nameBuf[i]) {
                for (UINT j = 0; j < SysStringLen(nameBuf[i]); j++) {
                    pname += (char)nameBuf[i][j];
                }
                SysFreeString(nameBuf[i]);
            }
            paramNames.push_back(pname);
        }
        // 释放未使用的BSTR
        for (UINT i = cNamesOut; i <= pFD->cParams; i++) {
            if (nameBuf[i]) SysFreeString(nameBuf[i]);
        }
    }

    for (UINT i = 0; i < pFD->cParams; i++) {
        std::string pname = (i < paramNames.size()) ? paramNames[i] : ("p" + std::to_string(i));
        ComParamInfo param = mapElemDesc(&pFD->lprgelemdescParam[i], pname, pTI);

        // [out, retval] 是返回值参数
        if (pFD->lprgelemdescParam[i].paramdesc.wParamFlags & PARAMFLAG_FRETVAL) {
            param.direction = ComParamDir::RetVal;
            member.returnType = param.type;
        }

        member.params.push_back(std::move(param));
    }

    // Fix-074: 对于 VT_HRESULT/VT_VOID 方法（如 dispinterface 属性 Get），
    // 如果参数中没有 PARAMFLAG_FRETVAL 标志，尝试从最后一个 [out] 参数推断返回类型
    if (member.returnType == Vb6Type::Empty && pFD->cParams > 0) {
        // 检查最后一个参数是否是 [out]
        UINT lastIdx = pFD->cParams - 1;
        ELEMDESC* lastParam = &pFD->lprgelemdescParam[lastIdx];
        if (lastParam->paramdesc.wParamFlags & PARAMFLAG_FOUT) {
            ComParamInfo lastP = mapElemDesc(lastParam, "_retval", pTI);
            member.returnType = lastP.type;
        }
    }

    return member;
}

// ============================================================
// 内部: 解析变量属性
// ============================================================

ComMemberInfo TypeLibParser::parseVarDesc(void* pTypeInfo, void* pVarDesc) {
    ITypeInfo* pTI = static_cast<ITypeInfo*>(pTypeInfo);
    VARDESC* pVD = static_cast<VARDESC*>(pVarDesc);

    ComMemberInfo member;
    member.memid = pVD->memid;
    member.kind = ComMemberKind::PropertyGet;
    member.vtableIndex = -1;  // 变量属性无vtable偏移

    // 获取名称
    BSTR varName = nullptr;
    pTI->GetDocumentation(pVD->memid, &varName, nullptr, nullptr, nullptr);
    if (varName) {
        member.realName.clear();
        for (UINT i = 0; i < SysStringLen(varName); i++) {
            member.realName += (char)varName[i];
        }
        SysFreeString(varName);
    }
    member.name = member.realName;
    std::transform(member.name.begin(), member.name.end(), member.name.begin(), ::tolower);

    // 类型 (账 #245): 属性在类型库里是 VARDESC, VAR_PROPERTY / VAR_CONST / VAR_PERINSTANCE
    // 三种 varkind 的真类型都记在 elemdescVar.tdesc 上 —— 以前只在 PERINSTANCE 那支赋值,
    // 另两支把没赋过值的 returnType 交出去, 下游把垃圾当「这属性是什么型」的答案用。
    member.returnType = mapTypeDesc(&pVD->elemdescVar.tdesc, pTI);

    return member;
}

// ============================================================
// 内部: TYPEDESC → Vb6Type
// ============================================================

Vb6Type TypeLibParser::mapTypeDesc(void* pTypeDesc, void* pTypeInfo) {
    if (!pTypeDesc) return Vb6Type::Variant;
    TYPEDESC* pTD = static_cast<TYPEDESC*>(pTypeDesc);
    ITypeInfo* pTI = static_cast<ITypeInfo*>(pTypeInfo);

    VARTYPE vt = pTD->vt;

    // 处理VT_PTR / VT_SAFEARRAY / VT_USERDEFINED 等间接类型
    if (vt == VT_PTR) {
        // 指针类型, 递归解引用
        if (pTD->lptdesc) {
            Vb6Type inner = mapTypeDesc(pTD->lptdesc, pTypeInfo);
            if (inner == Vb6Type::Object) return Vb6Type::Object;
            // 字符串指针 → String
            if (inner == Vb6Type::String) return Vb6Type::String;
            // SAFEARRAY指针 (VT_PTR → VT_SAFEARRAY) 保留数组类型
            bool isArray = (static_cast<uint16_t>(inner) & static_cast<uint16_t>(Vb6Type::Array)) != 0;
            if (isArray) return inner;
            // Fix-074: VT_PTR → 标量类型别名(如 OLE_XSIZE_HIMETRIC→Long, OLE_HANDLE→Long)
            // 在 COM typelib 中，属性返回类型 VT_PTR → TKIND_ALIAS → VT_I4 是合法的
            // 递归解析已正确还原底层类型，不应被降级为 Object
            switch (inner) {
                case Vb6Type::Long:
                case Vb6Type::Integer:
                case Vb6Type::Byte:
                case Vb6Type::Boolean:
                case Vb6Type::Single:
                case Vb6Type::Double:
                case Vb6Type::Currency:
                case Vb6Type::Date:
                case Vb6Type::Error:
                case Vb6Type::Decimal:
                    return inner;
                default:
                    break;
            }
            // 其他指针 → Object (接口指针)
            return Vb6Type::Object;
        }
        return Vb6Type::Object;
    }

    // Fix-074: VT_BYREF 处理 — COM dispatch 接口属性可能使用 VT_BYREF 指向实际类型
    // 例如 IPicture.Width 的 TypeLib 描述为 VT_BYREF → OLE_XSIZE_HIMETRIC → Long
    if (vt == VT_BYREF) {
        if (pTD->lptdesc) {
            Vb6Type inner = mapTypeDesc(pTD->lptdesc, pTypeInfo);
            // 标量类型直接透传（与 VT_PTR 相同逻辑）
            switch (inner) {
                case Vb6Type::Long:
                case Vb6Type::Integer:
                case Vb6Type::Byte:
                case Vb6Type::Boolean:
                case Vb6Type::Single:
                case Vb6Type::Double:
                case Vb6Type::Currency:
                case Vb6Type::Date:
                case Vb6Type::Error:
                case Vb6Type::Decimal:
                case Vb6Type::String:
                    return inner;
                case Vb6Type::Object:
                    return Vb6Type::Object;
                default:
                    break;
            }
            bool isArray = (static_cast<uint16_t>(inner) & static_cast<uint16_t>(Vb6Type::Array)) != 0;
            if (isArray) return inner;
            return Vb6Type::Variant;
        }
        return Vb6Type::Variant;
    }

    if (vt == VT_SAFEARRAY) {
        // SAFEARRAY → 数组
        if (pTD->lptdesc) {
            Vb6Type inner = mapTypeDesc(pTD->lptdesc, pTypeInfo);
            return (Vb6Type)((int)Vb6Type::Array | (int)inner);
        }
        return Vb6Type::Variant;
    }

    if (vt == VT_USERDEFINED) {
        // 用户定义类型, 通过HREFTYPE解析
        if (pTD->hreftype && pTI) {
            ITypeInfo* pRefTI = nullptr;
            HRESULT hr = pTI->GetRefTypeInfo(pTD->hreftype, &pRefTI);
            if (SUCCEEDED(hr) && pRefTI) {
                Vb6Type refResult = Vb6Type::UserDefinedType;
                TYPEATTR* pRefAttr = nullptr;
                hr = pRefTI->GetTypeAttr(&pRefAttr);
                if (SUCCEEDED(hr) && pRefAttr) {
                    if (pRefAttr->typekind == TKIND_INTERFACE ||
                        pRefAttr->typekind == TKIND_DISPATCH) {
                        refResult = Vb6Type::Object;  // 接口 → Object
                    } else if (pRefAttr->typekind == TKIND_ENUM) {
                        refResult = Vb6Type::Long;  // Enum → Long
                    } else if (pRefAttr->typekind == TKIND_RECORD) {
                        refResult = Vb6Type::UserDefinedType;  // UDT
                    } else if (pRefAttr->typekind == TKIND_ALIAS) {
                        // TKIND_ALIAS: 递归解析别名指向的实际类型
                        // tdescAlias字段包含别名的底层TYPEDESC
                        refResult = mapTypeDesc(&pRefAttr->tdescAlias, pRefTI);
                    }
                    pRefTI->ReleaseTypeAttr(pRefAttr);
                }
                pRefTI->Release();
                return refResult;
            }
        }
        return Vb6Type::Variant;
    }

    // 直接VARTYPE映射
    switch (vt & VT_TYPEMASK) {
        case VT_EMPTY:    return Vb6Type::Empty;
        case VT_NULL:     return Vb6Type::Null;
        case VT_I2:       return Vb6Type::Integer;
        case VT_I4:       return Vb6Type::Long;
        case VT_R4:       return Vb6Type::Single;
        case VT_R8:       return Vb6Type::Double;
        case VT_CY:       return Vb6Type::Currency;
        case VT_DATE:     return Vb6Type::Date;
        case VT_BSTR:     return Vb6Type::String;
        case VT_DISPATCH: return Vb6Type::Object;
        case VT_ERROR:    return Vb6Type::Error;
        case VT_BOOL:     return Vb6Type::Boolean;
        case VT_VARIANT:  return Vb6Type::Variant;
        case VT_DECIMAL:  return Vb6Type::Decimal;
        case VT_I1:       return Vb6Type::Byte;
        case VT_UI1:      return Vb6Type::Byte;
        case VT_UI2:      return Vb6Type::Integer;
        case VT_UI4:      return Vb6Type::Long;
        case VT_UNKNOWN:  return Vb6Type::Object;
        default:          return Vb6Type::Variant;
    }
}

// ============================================================
// 内部: ELEMDESC → ComParamInfo
// ============================================================

ComParamInfo TypeLibParser::mapElemDesc(void* pElemDesc, const std::string& name,
                                         void* pTypeInfo) {
    if (!pElemDesc) return {name, Vb6Type::Variant, ComParamDir::In};
    ELEMDESC* pED = static_cast<ELEMDESC*>(pElemDesc);

    ComParamInfo param;
    param.name = name;
    param.type = mapTypeDesc(&pED->tdesc, pTypeInfo);

    // 参数方向
    WORD flags = pED->paramdesc.wParamFlags;
    if (flags & PARAMFLAG_FOUT) {
        if (flags & PARAMFLAG_FRETVAL) {
            param.direction = ComParamDir::RetVal;
        } else if (flags & PARAMFLAG_FIN) {
            param.direction = ComParamDir::InOut;
        } else {
            param.direction = ComParamDir::Out;
        }
    } else {
        param.direction = ComParamDir::In;
    }

    param.isOptional = (flags & PARAMFLAG_FOPT) != 0;
    param.hasDefault = (flags & 0x10) != 0;  // PARAMFLAG_FDEFAULT = 0x10 (not in all SDK versions)

    return param;
}

// ============================================================
// 内部: IID → 字符串
// ============================================================

std::string TypeLibParser::iidToString(const uint8_t* iidBytes) {
    // GUID结构: Data1(4), Data2(2), Data3(2), Data4(8)
    const uint32_t* d1 = reinterpret_cast<const uint32_t*>(iidBytes);
    const uint16_t* d2 = reinterpret_cast<const uint16_t*>(iidBytes + 4);
    const uint16_t* d3 = reinterpret_cast<const uint16_t*>(iidBytes + 6);
    const uint8_t* d4 = iidBytes + 8;

    char buf[64];
    snprintf(buf, sizeof(buf), "{%08lX-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}",
             (unsigned long)*d1, *d2, *d3,
             d4[0], d4[1], d4[2], d4[3], d4[4], d4[5], d4[6], d4[7]);
    return buf;
}

} // namespace vb6c3
