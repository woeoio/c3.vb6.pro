#pragma once
// 编码工具 — Windows路径UTF-8/UTF-16/ACP互转
// VB6编译器内部统一使用UTF-8字符串, 但Windows filesystem::path按ACP解释char*,
// 导致VBP解析器输出的UTF-8字符串被双重编码。本模块提供安全转换函数。

#include <string>
#include <filesystem>
#include <fstream>
#include <cstdio>

#ifdef _WIN32
#include <windows.h>
#endif

namespace vb6c3 {

// filesystem::path → UTF-8 string
// On Windows, path::string() returns ACP-encoded string which is wrong for internal UTF-8 usage.
// Path internally stores UTF-16 on Windows, so use wstring() + conversion for correct results.
inline std::string pathToUtf8(const std::filesystem::path& p) {
#ifdef _WIN32
    std::wstring wide = p.wstring();
    if (wide.empty()) return "";
    int len = WideCharToMultiByte(CP_UTF8, 0, wide.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (len <= 0) return p.string();  // fallback
    std::string utf8(len, '\0');
    WideCharToMultiByte(CP_UTF8, 0, wide.c_str(), -1, &utf8[0], len, nullptr, nullptr);
    while (!utf8.empty() && utf8.back() == '\0') utf8.pop_back();
    return utf8;
#else
    return p.string();
#endif
}

// UTF-8 string → wide string
// When a string is already UTF-8 (e.g. from VBP parser GBK→UTF-8 conversion),
// we must NOT pass it to filesystem::path(const char*) which interprets as ACP on Windows.
inline std::wstring utf8ToWide(const std::string& utf8) {
#ifdef _WIN32
    if (utf8.empty()) return L"";
    int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, nullptr, 0);
    if (wlen <= 0) return L"";
    std::wstring wide(wlen, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, utf8.c_str(), -1, &wide[0], wlen);
    while (!wide.empty() && wide.back() == L'\0') wide.pop_back();
    return wide;
#else
    return std::wstring(utf8.begin(), utf8.end());
#endif
}

// wide string → UTF-8 string
// Windows: 用于把 GetCommandLineW() 得到的 Unicode 命令行参数转成内部统一的 UTF-8.
inline std::string wideToUtf8(const std::wstring& wide) {
#ifdef _WIN32
    if (wide.empty()) return "";
    int len = WideCharToMultiByte(CP_UTF8, 0, wide.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (len <= 0) return "";
    std::string utf8(len, '\0');
    WideCharToMultiByte(CP_UTF8, 0, wide.c_str(), -1, &utf8[0], len, nullptr, nullptr);
    while (!utf8.empty() && utf8.back() == '\0') utf8.pop_back();
    return utf8;
#else
    return std::string(wide.begin(), wide.end());
#endif
}

// UTF-8 string → filesystem::path (safe construction)
// On Windows, uses wstring to avoid ACP reinterpretation of UTF-8 bytes.
// On Linux/macOS, passes through directly (native UTF-8 filesystem).
inline std::filesystem::path utf8ToPath(const std::string& utf8) {
#ifdef _WIN32
    return std::filesystem::path(utf8ToWide(utf8));
#else
    return std::filesystem::path(utf8);
#endif
}

// ============================================================
// Fix 196: 文件流的 UTF-8 安全包装
// ------------------------------------------------------------
// std::ofstream/std::ifstream/std::fopen 的**窄字符串**重载在 Windows 上把 char*
// 当**系统 ANSI 代码页 (936/GBK)** 解释, 而 C3 内部路径一律 UTF-8。中文/韩文/俄文
// 目录下这些调用会打开乱码路径 —— 轻则「无法写入文件」直接失败, 重则写到一个
// **同名的乱码文件**里, 静默丢产物。下面包装统一经 utf8ToPath() 走宽字符重载。
//
// 实测 (非 ASCII %TEMP%):
//   std::ofstream ofs("C:\\Windows\\Temp\\新建temp\\C3C\\..\\FrxData.h")  -> 失败
//   ofstreamUtf8(同一字符串)                                            -> 成功
//
// ASCII 路径下与原行为逐字节一致 (多一次 UTF-8→UTF-16 转换), 无副作用。
// ============================================================

inline std::ofstream ofstreamUtf8(const std::string& utf8Path,
                                  std::ios::openmode mode = std::ios::out) {
    return std::ofstream(utf8ToPath(utf8Path), mode);
}

inline std::ifstream ifstreamUtf8(const std::string& utf8Path,
                                  std::ios::openmode mode = std::ios::in) {
    return std::ifstream(utf8ToPath(utf8Path), mode);
}

inline std::FILE* fopenUtf8(const std::string& utf8Path, const char* mode) {
#ifdef _WIN32
    return _wfopen(utf8ToPath(utf8Path).c_str(), utf8ToWide(mode).c_str());
#else
    return std::fopen(utf8Path.c_str(), mode);
#endif
}

// 窄串版存在性判断 (收 UTF-8 路径; 失败一律当"不存在", 不抛异常)
inline bool existsUtf8(const std::string& utf8Path) {
    std::error_code ec;
    return std::filesystem::exists(utf8ToPath(utf8Path), ec) && !ec;
}

// ============================================================
// 账 #267: 产物字节走 stdout，不过控制台代码页那一层
// ------------------------------------------------------------
// std::cout 在 main() 一进来就被 ConsoleUtf8Buf 接管了：**句柄是控制台**就转宽字符交
// WriteConsoleW（渲染与 chcp 无关），**句柄是管道/文件**就按**控制台代码页**转字节 —— 因为
// 诊断文本是给那个人/那个终端看的，实测 PS 5.1 与 pwsh 7 的 [Console]::OutputEncoding 都跟
// 控制台代码页走。这个口径对"消息"是对的，对"产物"是错的：--emit-c 交出的是一份 C 源码，
// 而同一次发码写进中间目录的那两份 .h/.c 是**内部 UTF-8 原样**（cl 正是拿 /utf-8 编它们的，
// 见 msvc_driver.cpp 的命令行）。同一个问题两个答案，其中一个还是机器的函数：本机控制台
// 代码页 936 ⇒ stdout 落 GBK，CI runner 是 65001 ⇒ 落 UTF-8 —— 实测同一枚 C3.exe、同一次运行，
// Temp\C3C\...\AccMain.c 里"过程实现"是 UTF-8 (E8 BF 87…)、--emit-c 的 stdout 里是 GBK (B9 FD…)。
//
// 所以产物这一路直接 fwrite 到 stdout 的 C 流（文本模式：'\n' 照旧翻成 CRLF，与 ofstream 写的
// 那两份一致），调用方先 cout.flush() 把压着的诊断落干净，顺序不乱。
// ============================================================
inline void writeStdoutRaw(const std::string& bytes) {
    if (bytes.empty()) return;
    std::fwrite(bytes.data(), 1, bytes.size(), stdout);
}

} // namespace vb6c3
