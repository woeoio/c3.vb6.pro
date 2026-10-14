// vb6rtl_registry.c - VB6 运行时库: 注册表家族：SaveSetting/GetSetting/DeleteSetting/GetAllSettings
// 2026-09-17 从 src/rtl/core/vb6rtl/vb6rtl.c 按家族拆出（纯搬移，逐行未改）:
//   原第 4794~4954 行

#include "vb6rtl.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <stdarg.h>
#include <wchar.h>
#include <wctype.h>
#include <time.h>
#ifdef _WIN32
#include <direct.h>
#include <io.h>
#include <oleauto.h>
#include <olectl.h>
#include <windows.h>
#endif

// P24-08: MessageBoxW (user32) + GetConsoleWindow (kernel32)
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "kernel32.lib")

// ============================================================
// P20-37: Registry functions (VB6: SaveSetting/GetSetting/DeleteSetting/GetAllSettings)
// VB6 registry path: HKEY_CURRENT_USER\Software\VB and VBA Program Settings\
// ============================================================

static BSTR vb6_RegBuildKey(BSTR appName, BSTR section) {
    const wchar_t* base = L"Software\\VB and VBA Program Settings\\";
    int baseLen = (int)wcslen(base);
    int appLen = appName ? (int)wcslen(appName) : 0;
    int secLen = section ? (int)wcslen(section) : 0;
    int totalLen = baseLen + appLen + 1 + secLen + 1;
    wchar_t* buf = (wchar_t*)calloc(totalLen + 1, sizeof(wchar_t));
    if (!buf) return vb6_BSTR_Empty();
    memcpy(buf, base, baseLen * sizeof(wchar_t));
    if (appName) memcpy(buf + baseLen, appName, appLen * sizeof(wchar_t));
    buf[baseLen + appLen] = L'\\';
    if (section) memcpy(buf + baseLen + appLen + 1, section, secLen * sizeof(wchar_t));
    buf[baseLen + appLen + 1 + secLen] = L'\0';
    BSTR result = SysAllocString(buf);
    free(buf);
    return result;
}

void vb6_SaveSetting(BSTR appName, BSTR section, BSTR key, BSTR setting) {
    if (!appName || !section || !key) return;
    BSTR keyPath = vb6_RegBuildKey(appName, section);
    HKEY hKey;
    DWORD disp;
    LONG rc = RegCreateKeyExW(HKEY_CURRENT_USER, keyPath, 0, NULL, 0, KEY_SET_VALUE, NULL, &hKey, &disp);
    SysFreeString(keyPath);
    if (rc != ERROR_SUCCESS) return;
    if (setting) {
        RegSetValueExW(hKey, key, 0, REG_SZ, (const BYTE*)setting, (DWORD)(wcslen(setting) + 1) * sizeof(wchar_t));
    } else {
        RegSetValueExW(hKey, key, 0, REG_SZ, (const BYTE*)L"", sizeof(wchar_t));
    }
    RegCloseKey(hKey);
}

BSTR vb6_GetSetting(BSTR appName, BSTR section, BSTR key, BSTR default_) {
    if (!appName || !section || !key) return default_ ? SysAllocString(default_) : vb6_BSTR_Empty();
    BSTR keyPath = vb6_RegBuildKey(appName, section);
    HKEY hKey;
    LONG rc = RegOpenKeyExW(HKEY_CURRENT_USER, keyPath, 0, KEY_QUERY_VALUE, &hKey);
    SysFreeString(keyPath);
    if (rc != ERROR_SUCCESS) return default_ ? SysAllocString(default_) : vb6_BSTR_Empty();
    
    DWORD dataSize = 0;
    rc = RegQueryValueExW(hKey, key, NULL, NULL, NULL, &dataSize);
    if (rc != ERROR_SUCCESS || dataSize == 0) {
        RegCloseKey(hKey);
        return default_ ? SysAllocString(default_) : vb6_BSTR_Empty();
    }
    
    wchar_t* buf = (wchar_t*)calloc(dataSize + sizeof(wchar_t), 1);
    if (!buf) {
        RegCloseKey(hKey);
        return default_ ? SysAllocString(default_) : vb6_BSTR_Empty();
    }
    rc = RegQueryValueExW(hKey, key, NULL, NULL, (LPBYTE)buf, &dataSize);
    RegCloseKey(hKey);
    if (rc != ERROR_SUCCESS) {
        free(buf);
        return default_ ? SysAllocString(default_) : vb6_BSTR_Empty();
    }
    BSTR result = SysAllocString(buf);
    free(buf);
    return result;
}

void vb6_DeleteSetting(BSTR appName, BSTR section, BSTR key) {
    if (!appName || !section) return;
    BSTR keyPath = vb6_RegBuildKey(appName, section);
    if (key && wcslen(key) > 0) {
        HKEY hKey;
        LONG rc = RegOpenKeyExW(HKEY_CURRENT_USER, keyPath, 0, KEY_SET_VALUE, &hKey);
        SysFreeString(keyPath);
        if (rc == ERROR_SUCCESS) {
            RegDeleteValueW(hKey, key);
            RegCloseKey(hKey);
        }
    } else {
        RegDeleteKeyW(HKEY_CURRENT_USER, keyPath);
        SysFreeString(keyPath);
    }
}

vb6_VARIANT vb6_GetAllSettings(BSTR appName, BSTR section) {
    vb6_VARIANT result;
    memset(&result, 0, sizeof(result));
    /* 账 #305 §B149 (2026-10-11 第四十刀): VB6 这一族只有两档 ——
     *   应用/节没有条目  => Null
     *   有条目          => **二维** Variant 数组: 第一维 0..n-1 是键的序号,
     *                      第二维 **1..2** (1 = 键名, 2 = 值)
     * 外部证人: VB6 真语料 mdoceditor/E-Type/Main.bas:269 写的就是
     *   `For i = 0 To UBound(varFiles, 1)` + `varFiles(i, 1)`。
     * 改前这里交回一枚 vt 写着 vb6_vtVariant、parray 挂着 1D 交替表的槽 —— 两档都不是,
     * 于是五读全错而一声不响: VarType=12 / IsNull=False / UBound(_,1)=0 / UBound(_,2)=0 /
     * v(0,2) 读回空串 (读数 .build/b1265_nd/r6.txt, 两台同形)。 */
    result.vt = vb6_vtNull;
    if (!appName || !section) return result;

    BSTR keyPath = vb6_RegBuildKey(appName, section);
    HKEY hKey;
    LONG rc = RegOpenKeyExW(HKEY_CURRENT_USER, keyPath, 0, KEY_QUERY_VALUE, &hKey);
    SysFreeString(keyPath);
    if (rc != ERROR_SUCCESS) return result;

    DWORD numValues = 0;
    DWORD maxNameLen = 0;
    DWORD maxDataLen = 0;
    RegQueryInfoKeyW(hKey, NULL, NULL, NULL, NULL, NULL, NULL, &numValues, &maxNameLen, &maxDataLen, NULL, NULL);
    if (numValues == 0) {
        RegCloseKey(hKey);
        return result;
    }

    vb6_SafeArrayBound bounds[2];
    bounds[0].lBound = 0;
    bounds[0].cElements = (int32_t)numValues;
    bounds[1].lBound = 1;
    bounds[1].cElements = 2;
    vb6_SafeArrayND* nd = vb6_SafeArrayCreateND(vb6_sa_variant, 2, bounds);
    if (!nd) {
        RegCloseKey(hKey);
        return result;
    }

    wchar_t* nameBuf = (wchar_t*)calloc(maxNameLen + 2, sizeof(wchar_t));
    BYTE* dataBuf = (BYTE*)calloc(maxDataLen + sizeof(wchar_t), 1);
    if (!nameBuf || !dataBuf) {
        /* 连名字缓冲都分配不出来 => 按"没有条目"那一档答, 不交半枚数组 */
        free(nameBuf);
        free(dataBuf);
        vb6_SafeArrayDestroyND(nd);
        RegCloseKey(hKey);
        memset(&result, 0, sizeof(result));
        result.vt = vb6_vtNull;
        return result;
    }

    for (DWORD i = 0; i < numValues; i++) {
        DWORD nameSize = maxNameLen + 2;
        DWORD dataSize = maxDataLen + sizeof(wchar_t);
        DWORD type = 0;
        nameBuf[0] = L'\0';
        dataBuf[0] = 0;
        if (RegEnumValueW(hKey, i, nameBuf, &nameSize, NULL, &type, dataBuf, &dataSize) != ERROR_SUCCESS) {
            type = 0;
            dataSize = sizeof(wchar_t);
        }

        vb6_VARIANT varKey;
        memset(&varKey, 0, sizeof(varKey));
        varKey.vt = vb6_vtBSTR;
        varKey.bstrVal = SysAllocString(nameBuf);

        vb6_VARIANT varVal;
        memset(&varVal, 0, sizeof(varVal));
        varVal.vt = vb6_vtBSTR;
        if (type == REG_SZ && dataSize >= sizeof(wchar_t)) {
            varVal.bstrVal = SysAllocString((wchar_t*)dataBuf);
        } else {
            varVal.bstrVal = vb6_BSTR_Empty();
        }

        /* 两个格子都从那一枚带检查的 ND 落点里拿 (账 #209 那一族的口径);
         * Assign 是"接管"语义 (rev16 那一课), 所以借用出来的 BSTR 有唯一的主人。
         * 步长问 vb6_SaElemSize —— 「Variant 槽数组一格多宽」只许那一张表回答。 */
        const int32_t es = vb6_SaElemSize(vb6_sa_variant);
        int32_t row[2];
        row[0] = (int32_t)i;
        row[1] = 1;
        vb6_VariantAssign((vb6_VARIANT*)vb6_SaNdElemPtr(nd, 2, row, es), varKey);
        row[1] = 2;
        vb6_VariantAssign((vb6_VARIANT*)vb6_SaNdElemPtr(nd, 2, row, es), varVal);
    }

    free(nameBuf);
    free(dataBuf);
    RegCloseKey(hKey);

    /* 账 #305 §B155 格①（第四十六刀）: 问描述符而不是写死；这一枚是 variant 档 ⇒ 仍 8204 = vbArray + vbVariant */
    result.vt = (vb6_vartype)vb6_SaVtOfArrayPtr(nd);
    result.parray = (struct vb6_SafeArray1D*)nd;
    return result;
}

