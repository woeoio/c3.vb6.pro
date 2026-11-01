// tlbprobe.cpp - 读一个 .tlb，把「LIBID / 每个类型的 GUID / coclass 引用了哪些接口 / 方法表」
// 打到 stdout。用途：证明 COM 服务器表（dll_entry.c）与类型库里的身份 GUID 是否逐值一致
// （ai/022 D57-5 的未证条目，B13c 读数用）。
// 编译：cl /nologo /EHsc /utf-8 tlbprobe.cpp /Fe:tlbprobe.exe /link ole32.lib oleaut32.lib
#include <windows.h>
#include <ole2.h>
#include <oleauto.h>
#include <oaidl.h>
#include <stdio.h>
#include <string>

static std::string w2u(const OLECHAR* w) {
    if (!w) return std::string();
    int n = WideCharToMultiByte(CP_UTF8, 0, w, -1, nullptr, 0, nullptr, nullptr);
    if (n <= 0) return std::string();
    std::string s((size_t)n - 1, '\0');
    WideCharToMultiByte(CP_UTF8, 0, w, -1, &s[0], n, nullptr, nullptr);
    return s;
}

static std::string a2w(const std::string& s) {
    int n = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, nullptr, 0);
    std::string out((size_t)(n > 0 ? n : 1) * sizeof(wchar_t), '\0');
    if (n > 0) MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, (__wchar_t*)&out[0], n);
    return out;
}

static std::string guidStr(const GUID& g) {
    OLECHAR buf[64] = {0};
    if (FAILED(StringFromGUID2(g, buf, 64))) return std::string("?");
    std::string s = w2u(buf);
        for (auto& c : s) c = (char)toupper((unsigned char)c);
    return s;
}

static const char* kindName(TYPEKIND k) {
    switch (k) {
        case TKIND_ENUM: return "enum";
        case TKIND_RECORD: return "record";
        case TKIND_MODULE: return "module";
        case TKIND_INTERFACE: return "interface";
        case TKIND_DISPATCH: return "dispinterface";
        case TKIND_COCLASS: return "coclass";
        case TKIND_ALIAS: return "alias";
        case TKIND_UNION: return "union";
        default: return "other";
    }
}

static void dumpRef(ITypeInfo* ti, UINT r) {
    HREFTYPE hRef = 0;
    INT rflags = 0;
    if (FAILED(ti->GetImplTypeFlags(r, &rflags))) return;
    if (FAILED(ti->GetRefTypeOfImplType(r, &hRef))) return;
    ITypeInfo* rti = nullptr;
    if (FAILED(ti->GetRefTypeInfo(hRef, &rti)) || !rti) return;
    BSTR rn = nullptr;
    rti->GetDocumentation(-1, &rn, nullptr, nullptr, nullptr);
    std::string refName = w2u(rn);
    SysFreeString(rn);
    TYPEATTR* rta = nullptr;
    std::string refGuid = "?", refKind = "?";
    if (SUCCEEDED(rti->GetTypeAttr(&rta)) && rta) {
        refGuid = guidStr(rta->guid);
        refKind = kindName(rta->typekind);
        rti->ReleaseTypeAttr(rta);
    }
    unsigned flags = (unsigned)rflags;
    std::string ann;
    if (rflags & 1) ann += " DEFAULT";
    if (rflags & 2) ann += " SOURCE";
    if (rflags & 4) ann += " DEFAULTVTABLE";
    printf("  REF %u kind=%s name=%s guid=%s flags=0x%X%s\n", r, refKind.c_str(), refName.c_str(),
           refGuid.c_str(), flags, ann.c_str());
    rti->Release();
}

static void dumpType(ITypeLib* lib, UINT i) {
    ITypeInfo* ti = nullptr;
    if (FAILED(lib->GetTypeInfo(i, &ti)) || !ti) {
        printf("TYPE %u <unreadable>\n", i);
        return;
    }
    BSTR name = nullptr;
    lib->GetDocumentation(i, &name, nullptr, nullptr, nullptr);
    std::string typeName = w2u(name);
    SysFreeString(name);

    TYPEATTR* ta = nullptr;
    if (FAILED(ti->GetTypeAttr(&ta)) || !ta) {
        printf("TYPE %u name=%s <no TYPEATTR>\n", i, typeName.c_str());
        ti->Release();
        return;
    }
    {
        std::string g = guidStr(ta->guid);
        const char* k = kindName(ta->typekind);
        printf("TYPE %u kind=%s name=%s guid=%s typeflags=0x%X\n", i, k, typeName.c_str(), g.c_str(),
               (unsigned)ta->wTypeFlags);
    }
    for (UINT r = 0; r < ta->cImplTypes; r++) dumpRef(ti, r);
    if (ta->typekind == TKIND_ENUM) {
        for (UINT v = 0; v < ta->cVars; v++) {
            VARDESC* vd = nullptr;
            if (FAILED(ti->GetVarDesc(v, &vd)) || !vd) continue;
            BSTR vn = nullptr;
            ti->GetDocumentation(vd->memid, &vn, nullptr, nullptr, nullptr);
            std::string varName = w2u(vn);
            SysFreeString(vn);
            long val = 0;
            if (vd->lpvarValue) {
                VARIANT tmp;
                VariantInit(&tmp);
                if (SUCCEEDED(VariantChangeType(&tmp, (VARIANT*)vd->lpvarValue, 0, VT_I4))) val = tmp.lVal;
                VariantClear(&tmp);
            }
            printf("  VAR %u name=%s value=%ld\n", v, varName.c_str(), val);
            ti->ReleaseVarDesc(vd);
        }
    }
    for (UINT f = 0; f < ta->cFuncs; f++) {
        FUNCDESC* fa = nullptr;
        if (FAILED(ti->GetFuncDesc(f, &fa)) || !fa) continue;
        BSTR fn = nullptr;
        ti->GetDocumentation(fa->memid, &fn, nullptr, nullptr, nullptr);
        std::string funcName = w2u(fn);
        SysFreeString(fn);
        int memid = (int)fa->memid;
        int invkind = (int)fa->invkind;
        unsigned nparams = (unsigned)fa->cParams;
        printf("  FUNC %u memid=%d name=%s invkind=%d nparams=%u\n", f, memid, funcName.c_str(),
               invkind, nparams);
        ti->ReleaseFuncDesc(fa);
    }
    ti->ReleaseTypeAttr(ta);
    ti->Release();
}

int main(int argc, char** argv) {
    if (argc < 2) {
        fprintf(stderr, "usage: tlbprobe <path.tlb>\n");
        return 2;
    }
    CoInitialize(nullptr);
    ITypeLib* lib = nullptr;
    std::string pathW = a2w(argv[1]);
    const wchar_t* pathArg = (const wchar_t*)pathW.data();
    HRESULT hr = LoadTypeLib(pathArg, &lib);
    if (FAILED(hr) || !lib) {
        fprintf(stderr, "LoadTypeLib failed hr=0x%08X\n", (unsigned)hr);
        CoUninitialize();
        return 1;
    }
    UINT count = lib->GetTypeInfoCount();

    TLIBATTR* la = nullptr;
    lib->GetLibAttr(&la);
    std::string libGuid = la ? guidStr(la->guid) : std::string("?");
    if (la) lib->ReleaseTLibAttr(la);

    BSTR libName = nullptr;
    lib->GetDocumentation(-1, &libName, nullptr, nullptr, nullptr);
    std::string libNameU = w2u(libName);
    SysFreeString(libName);

    printf("LIB guid=%s name=%s types=%u\n", libGuid.c_str(), libNameU.c_str(), count);
    for (UINT i = 0; i < count; i++) dumpType(lib, i);
    lib->Release();
    CoUninitialize();
    return 0;
}
