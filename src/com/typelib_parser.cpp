// VB6 TypeLib解析器实现 - P6.3 前期绑定支持
// 编译期使用Windows LoadTypeLib/ITypeInfo API

#include "com/typelib_parser.hpp"
#include <algorithm>
#include <cstdio>
#include <windows.h>
#include <oleauto.h>

namespace vb6c3 {

// --- typelib_parser.cpp: 加载入口（按路径 / CLSID / 名称 / ProgID 定位 TypeLib） ---

// P24-12: 扩展注册表路径中的环境变量 (如 %SystemRoot%)
static std::wstring expandEnvironmentStrings(const std::wstring& src) {
    DWORD len = ExpandEnvironmentStringsW(src.c_str(), nullptr, 0);
    if (len == 0) return src;
    std::wstring result(len - 1, L'\\0');
    ExpandEnvironmentStringsW(src.c_str(), &result[0], len);
    return result;
}

// 诊断消息走窄串 (stderr / c3-error.log 是 ACP 口径) => 宽串按 **CP_ACP** 窄串化。
// 不用 common/encoding.hpp 的 wideToUtf8: 那给 UTF-8, 非 ASCII 路径在 ACP 出口
// 会乱码 —— 比「打注册表原值」更糟 (那一份本来就是 ANSI)。
static std::string wideToAcpNarrow(const std::wstring& w) {
    if (w.empty()) return std::string();
    const int n = WideCharToMultiByte(CP_ACP, 0, w.c_str(), (int)w.size(),
                                      nullptr, 0, nullptr, nullptr);
    if (n <= 0) return std::string();
    std::string out(static_cast<size_t>(n), '\0');
    WideCharToMultiByte(CP_ACP, 0, w.c_str(), (int)w.size(), &out[0], n,
                        nullptr, nullptr);
    return out;
}

// ============================================================
// 构造/析构
// ============================================================

TypeLibParser::TypeLibParser(Diagnostics& diag)
    : diag_(diag) {}

TypeLibParser::~TypeLibParser() = default;

// ============================================================
// 公共接口
// ============================================================

std::unique_ptr<TypeLibResult> TypeLibParser::loadByProgId(const std::string& progId, bool silent) {
    // 1. 查找缓存
    ComCoClassInfo* cached = findCachedCoClass(progId);
    if (cached) {
        // 已缓存, 返回所在TypeLib的拷贝引用 (无需重新解析)
        // 返回nullptr表示已在缓存中, 调用方应使用findCachedCoClass
        return nullptr;
    }

    // 2. 从注册表查找TypeLib路径
    std::string tlbPath = findTypeLibPathForProgId(progId);
    if (tlbPath.empty()) {
        // silent模式: auto-load常用组件未安装是预期情况, 不报warning避免污染c3-error.log
        if (!silent) {
            diag_.warn(DiagnosticID::CodeGenUnsupportedFeature, SourceLocation{}, "TypeLib not found for ProgID: " + progId);
        }
        return nullptr;
    }

    // 3. 按路径加载
    return loadByPath(tlbPath);
}

std::unique_ptr<TypeLibResult> TypeLibParser::loadByPath(const std::string& tlbPath) {
    // P24-05: 规范化路径(短路径→长路径), 避免同DLL不同路径导致缓存未命中
    std::wstring tlbPathW(tlbPath.begin(), tlbPath.end());
    tlbPathW = expandEnvironmentStrings(tlbPathW);  // P24-12: expand %SystemRoot% etc.
    wchar_t canonicalPath[MAX_PATH];
    DWORD len = GetLongPathNameW(tlbPathW.c_str(), canonicalPath, MAX_PATH);
    std::string canonPath;
    if (len > 0 && len < MAX_PATH) {
        // 统一转小写用于缓存比较 (Windows路径不区分大小写)
        canonPath = std::string(canonicalPath, canonicalPath + len);
        for (auto& c : canonPath) c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
    } else {
        canonPath = tlbPath;
        for (auto& c : canonPath) c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
    }

    // 检查缓存 (用规范化路径比较)
    for (auto& cached : cache_) {
        if (cached->canonPath == canonPath) {
            return nullptr;  // 已缓存
        }
    }

    // 加载TypeLib
    ITypeLib* pTypeLib = nullptr;
    HRESULT hr = LoadTypeLibEx(
        tlbPathW.c_str(),
        REGKIND_NONE,
        &pTypeLib
    );

    if (FAILED(hr) || !pTypeLib) {
        // 路径打**展开后**的那一份 (tlbPathW): 加载用的就是它, 消息若打注册表原值
        // (含 %CommonProgramFiles%) 会被读成「路径写错」—— 实际错因常是 32/64 位
        // 不匹配 (0x80029C4A = TYPE_E_CANTLOADLIBRARY: 32 位 DLL 在 64 位进程里必败)。
        // HRESULT 打 8 位十六进制: std::to_string 给的是十进制 (-2147312566),
        // 读不出这是个 HRESULT; 十六进制才是 HRESULT 的样子。
        char hrHex[16];
        snprintf(hrHex, sizeof(hrHex), "0x%08X", static_cast<unsigned int>(hr));
        diag_.warn(DiagnosticID::CodeGenUnsupportedFeature, SourceLocation{},
                   "Failed to load TypeLib: " + wideToAcpNarrow(tlbPathW) +
                   " (hr=" + hrHex + ")");
        return nullptr;
    }

    // 解析
    auto result = std::make_unique<TypeLibResult>();
    result->tlbPath = tlbPath;
    result->canonPath = canonPath;

    bool ok = parseTypeLib(pTypeLib, *result);
    pTypeLib->Release();

    if (!ok) {
        diag_.warn(DiagnosticID::CodeGenUnsupportedFeature, SourceLocation{},
                   "Failed to parse TypeLib: " + tlbPath);
        return nullptr;
    }

    // 链接coclass → default interface
    for (auto& cc : result->coclasses) {
        if (!cc->defaultIfaceName.empty()) {
            cc->defaultIface = result->findInterface(cc->defaultIfaceName);
        // P13.20: Link default source (event) interface
        if (!cc->defaultSourceIfaceName.empty()) {
            cc->defaultSourceIface = result->findInterface(cc->defaultSourceIfaceName);
        }
        }
    }

    // 缓存
    cache_.push_back(std::move(result));
    // 返回nullptr表示已缓存, 调用方应使用findCachedCoClass或cachedResults()
    return nullptr;
}

std::unique_ptr<TypeLibResult> TypeLibParser::loadByClsid(const std::string& clsidStr) {
    // P24-05: Object=行的GUID可能是TypeLib ID或CoClass CLSID
    // 先尝试按TypeLib ID直接查找 HKCR\TypeLib\{guid}, 失败再走CLSID反查
    std::wstring guidW(clsidStr.begin(), clsidStr.end());

    // --- 尝试1: 直接按TypeLib ID查找 ---
    {
        std::wstring tlbidKey = L"TypeLib\\" + guidW;
        HKEY hTlbKey;
        if (RegOpenKeyExW(HKEY_CLASSES_ROOT, tlbidKey.c_str(), 0, KEY_READ, &hTlbKey) == ERROR_SUCCESS) {
            // TypeLib ID有效, 枚举版本找win32路径
            wchar_t version[32];
            DWORD verSz = sizeof(version);
            DWORD idx = 0;
            std::wstring latestPath;
            while (RegEnumKeyExW(hTlbKey, idx, version, &verSz, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS) {
                verSz = sizeof(version);
                std::wstring pathKey = tlbidKey + L"\\" + version + L"\\0\\win32";
                HKEY hPathKey;
                if (RegOpenKeyExW(HKEY_CLASSES_ROOT, pathKey.c_str(), 0, KEY_READ, &hPathKey) == ERROR_SUCCESS) {
                    wchar_t path[MAX_PATH];
                    DWORD pathSz = sizeof(path);
                    if (RegQueryValueExW(hPathKey, nullptr, nullptr, nullptr, (LPBYTE)path, &pathSz) == ERROR_SUCCESS) {
                        latestPath = path;
                    }
                    RegCloseKey(hPathKey);
                }
                idx++;
            }
            RegCloseKey(hTlbKey);
            if (!latestPath.empty()) {
                std::string pathStr(latestPath.begin(), latestPath.end());
                return loadByPath(pathStr);
            }
        }
    }

    // --- 尝试2: 按CoClass CLSID反查 TypeLib ---
    // 格式: HKCR\CLSID\{...}\TypeLib → {typelibid}
    CLSID clsid;
    HRESULT hr = CLSIDFromString(guidW.c_str(), &clsid);
    if (FAILED(hr)) return nullptr;

    // 查TypeLib子键
    std::wstring keyPath = L"CLSID\\" + guidW + L"\\TypeLib";
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, keyPath.c_str(), 0, KEY_READ, &hKey) != ERROR_SUCCESS) {
        // P24-04 续: GUID 既不在 HKCR\TypeLib (尝试1) 也不在 HKCR\CLSID (尝试2),
        // 说明该类型库未注册到本机. 此处原为静默返回, 会让早绑定 / GlobalNameSpace
        // 语法退化成"模块名.成员"并生成未定义符号 (实测 tests/test_vbman 的
        // VBMAN.Version → vb6_VBMAN_Version → LNK2019), 故给出明确诊断.
        diag_.warn(DiagnosticID::CodeGenUnsupportedFeature, SourceLocation{},
                   "TypeLib not registered for CLSID/LibID: " + clsidStr);
        return nullptr;
    }
    wchar_t tlbidStr[64];
    DWORD sz = sizeof(tlbidStr);
    if (RegQueryValueExW(hKey, nullptr, nullptr, nullptr, (LPBYTE)tlbidStr, &sz) != ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return nullptr;
    }
    RegCloseKey(hKey);

    // TypeLib ID → 查TypeLib版本路径
    std::wstring tlbidKey2 = L"TypeLib\\" + std::wstring(tlbidStr);
    HKEY hTlbKey2;
    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, tlbidKey2.c_str(), 0, KEY_READ, &hTlbKey2) != ERROR_SUCCESS) {
        return nullptr;
    }
    wchar_t version2[32];
    DWORD idx2 = 0;
    std::wstring latestPath2;
    while (RegEnumKeyExW(hTlbKey2, idx2, version2, &sz, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS) {
        sz = sizeof(version2);
        std::wstring pathKey = tlbidKey2 + L"\\" + version2 + L"\\0\\win32";
        HKEY hPathKey;
        if (RegOpenKeyExW(HKEY_CLASSES_ROOT, pathKey.c_str(), 0, KEY_READ, &hPathKey) == ERROR_SUCCESS) {
            wchar_t path[MAX_PATH];
            DWORD pathSz = sizeof(path);
            if (RegQueryValueExW(hPathKey, nullptr, nullptr, nullptr, (LPBYTE)path, &pathSz) == ERROR_SUCCESS) {
                latestPath2 = path;
            }
            RegCloseKey(hPathKey);
        }
        idx2++;
    }
    RegCloseKey(hTlbKey2);

    if (latestPath2.empty()) {
        // P24-04 续: 显式引用的 TypeLib GUID/LibID 未在注册表登记时此处静默返回,
        // 会让 GlobalNameSpace / 早绑定语法 (如 VBMAN.Version) 退化成"模块名.成员",
        // 生成未定义符号 vb6_VBMAN_Version, 直到链接期才以 LNK2019 暴露,
        // 把排查引向错误方向 (实测 tests/test_vbman). 这里给出明确诊断,
        // 指向"类型库未注册且未提供有效路径"这一实际原因.
        diag_.warn(DiagnosticID::CodeGenUnsupportedFeature, SourceLocation{},
                   "TypeLib not registered for CLSID/LibID: " + clsidStr);
        return nullptr;
    }

    std::string pathStr2(latestPath2.begin(), latestPath2.end());
    return loadByPath(pathStr2);
}

std::unique_ptr<TypeLibResult> TypeLibParser::loadByName(const std::string& name,
                                                          const std::string& version) {
    // 按TypeLib名称从注册表查找
    // 枚举 HKCR\TypeLib 下的子键, 匹配名称
    std::wstring nameW(name.begin(), name.end());
    std::wstring verW(version.begin(), version.end());

    HKEY hTlbRoot;
    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, L"TypeLib", 0, KEY_READ, &hTlbRoot) != ERROR_SUCCESS) {
        return nullptr;
    }

    wchar_t tlbid[64];
    DWORD tlbidSz = sizeof(tlbid);
    DWORD idx = 0;
    while (RegEnumKeyExW(hTlbRoot, idx, tlbid, &tlbidSz, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS) {
        tlbidSz = sizeof(tlbid);
        // 检查此TLBID下是否有匹配版本
        std::wstring verKey = std::wstring(L"TypeLib\\") + tlbid + L"\\" + verW;
        HKEY hVer;
        if (RegOpenKeyExW(HKEY_CLASSES_ROOT, verKey.c_str(), 0, KEY_READ, &hVer) == ERROR_SUCCESS) {
            // 检查名称
            wchar_t tlbName[256];
            DWORD nameSz = sizeof(tlbName);
            if (RegQueryValueExW(hVer, nullptr, nullptr, nullptr, (LPBYTE)tlbName, &nameSz) == ERROR_SUCCESS) {
                if (_wcsicmp(tlbName, nameW.c_str()) == 0) {
                    RegCloseKey(hVer);
                    RegCloseKey(hTlbRoot);
                    // 找到! 加载win32路径
                    std::wstring pathKey = verKey + L"\\0\\win32";
                    HKEY hPath;
                    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, pathKey.c_str(), 0, KEY_READ, &hPath) == ERROR_SUCCESS) {
                        wchar_t path[MAX_PATH];
                        DWORD pathSz = sizeof(path);
                        if (RegQueryValueExW(hPath, nullptr, nullptr, nullptr, (LPBYTE)path, &pathSz) == ERROR_SUCCESS) {
                            RegCloseKey(hPath);
                            // wchar_t → std::string (窄字符转换)
                            std::string pathStr;
                            DWORD charCount = pathSz / sizeof(wchar_t);
                            for (DWORD k = 0; k < charCount && path[k] != 0; k++) {
                                pathStr += (char)path[k];
                            }
                            return loadByPath(pathStr);
                        }
                        RegCloseKey(hPath);
                    }
                    return nullptr;
                }
            }
            RegCloseKey(hVer);
        }
        idx++;
    }
    RegCloseKey(hTlbRoot);
    return nullptr;
}

ComCoClassInfo* TypeLibParser::findCachedCoClass(const std::string& name) const {
    for (auto& tl : cache_) {
        ComCoClassInfo* cc = tl->findCoClass(name);
        if (cc) return cc;
    }
    return nullptr;
}

// ============================================================
// 内部: ProgID → TypeLib路径
// ============================================================

std::string TypeLibParser::findTypeLibPathForProgId(const std::string& progId) {
    // HKCR\ProgID\CLSID → CLSID值
    // HKCR\CLSID\{clsid}\TypeLib → TypeLib ID
    // HKCR\TypeLib\{tlbid}\{version}\0\win32 → 路径

    std::wstring progIdW(progId.begin(), progId.end());

    // 1. ProgID → CLSID
    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(progIdW.c_str(), &clsid);
    if (FAILED(hr)) return {};

    // 2. CLSID → TypeLib ID
    OLECHAR clsidStr[64];
    StringFromGUID2(clsid, clsidStr, 64);
    std::wstring clsidKey = L"CLSID\\" + std::wstring(clsidStr) + L"\\TypeLib";

    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, clsidKey.c_str(), 0, KEY_READ, &hKey) != ERROR_SUCCESS) {
        return {};
    }

    wchar_t tlbidStr[64];
    DWORD sz = sizeof(tlbidStr);
    if (RegQueryValueExW(hKey, nullptr, nullptr, nullptr, (LPBYTE)tlbidStr, &sz) != ERROR_SUCCESS) {
        RegCloseKey(hKey);
        return {};
    }
    RegCloseKey(hKey);

    // 3. TypeLib ID → 找最新版本 → win32路径
    std::wstring tlbKey = L"TypeLib\\" + std::wstring(tlbidStr);
    HKEY hTlbKey;
    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, tlbKey.c_str(), 0, KEY_READ, &hTlbKey) != ERROR_SUCCESS) {
        return {};
    }

    // 枚举版本子键, 选最大的
    wchar_t bestVer[32] = {};
    DWORD idx = 0;
    wchar_t verBuf[32];
    DWORD verSz = sizeof(verBuf);
    while (RegEnumKeyExW(hTlbKey, idx, verBuf, &verSz, nullptr, nullptr, nullptr, nullptr) == ERROR_SUCCESS) {
        if (wcscmp(verBuf, bestVer) > 0) {
            wcscpy_s(bestVer, verBuf);
        }
        verSz = sizeof(verBuf);
        idx++;
    }
    RegCloseKey(hTlbKey);

    if (bestVer[0] == 0) return {};

    // 4. win32路径
    std::wstring pathKey = tlbKey + L"\\" + bestVer + L"\\0\\win32";
    HKEY hPathKey;
    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, pathKey.c_str(), 0, KEY_READ, &hPathKey) != ERROR_SUCCESS) {
        return {};
    }

    wchar_t path[MAX_PATH];
    DWORD pathSz = sizeof(path);
    if (RegQueryValueExW(hPathKey, nullptr, nullptr, nullptr, (LPBYTE)path, &pathSz) != ERROR_SUCCESS) {
        RegCloseKey(hPathKey);
        return {};
    }
    RegCloseKey(hPathKey);

    // 转换为窄字符串
    std::string result;
    for (DWORD i = 0; i < pathSz / sizeof(wchar_t); i++) {
        if (path[i] == 0) break;
        result += (char)path[i];
    }
    return result;
}

} // namespace vb6c3
