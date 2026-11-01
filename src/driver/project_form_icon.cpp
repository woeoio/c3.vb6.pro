// project_form_icon.cpp - 把 `.vbp` 的 IconForm 那一格读到链接期（账 #305 第328刀）
//
// 只读、不改写任何 VB 源文件。认不出的形状一律交回 false ⇒ 调用方落回第303刀那份默认窗体
// 图标，与「没有 Icon 属性」那一档同一个出口 —— 不猜、不半抄。

#include "driver/project_form_icon.hpp"

#include "common/encoding.hpp"
#include "driver/default_form_icon.hpp"

#include <cstdio>
#include <cstdint>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <ios>

namespace vb6c3 {
namespace {

std::uint32_t rdLE32(const unsigned char* p) {
    return static_cast<std::uint32_t>(p[0]) |
           (static_cast<std::uint32_t>(p[1]) << 8) |
           (static_cast<std::uint32_t>(p[2]) << 16) |
           (static_cast<std::uint32_t>(p[3]) << 24);
}
std::uint16_t rdLE16(const unsigned char* p) {
    return static_cast<std::uint16_t>(p[0] | (p[1] << 8));
}

bool ciEquals(const char* a, size_t n, const char* b) {
    if (n != std::strlen(b)) return false;
    for (size_t i = 0; i < n; i++) {
        char ca = a[i], cb = b[i];
        if (ca >= 'A' && ca <= 'Z') ca = static_cast<char>(ca + 32);
        if (cb >= 'A' && cb <= 'Z') cb = static_cast<char>(cb + 32);
        if (ca != cb) return false;
    }
    return true;
}

int hexVal(char c) {
    if (c >= '0' && c <= '9') return c - '0';
    if (c >= 'a' && c <= 'f') return c - 'a' + 10;
    if (c >= 'A' && c <= 'F') return c - 'A' + 10;
    return -1;
}

// 控件块那一行 = `Begin VB.PictureBox picX`。按整词认：`BeginProperty Font` 是**属性块内**的
// 子块（窗体自己的 `Icon` 常常排在它后面），前缀匹配会提前 break ⇒ 第328刀第一发就栽在这
// （实测BalloonTooltips 报 "no Icon property on that form"，而 .frm 第 18 行明明写着）。
bool isControlBeginLine(const std::string& line, size_t start) {
    if (!ciEquals(line.data() + start, 5, "Begin")) return false;
    size_t after = start + 5;
    return after >= line.size() || line[after] == ' ' || line[after] == '\t' ||
           line[after] == '\r' || line[after] == '\n';
}

// 一行形如 `   Icon            =   "Form1.frx":0000` ⇒ 取出 frx 名与偏移。
// 属性名按整词比（`IconText` / `IconHeight` 那种不算），等号两侧的空白照 VB6 的列对齐吃。
bool parseIconPropLine(const std::string& line, std::string& frxName, std::uint32_t& offset) {
    size_t i = line.find_first_not_of(" \t\r\n");
    if (i == std::string::npos) return false;
    size_t nameEnd = line.find_first_of(" \t=\r\n", i);
    if (nameEnd == std::string::npos) return false;
    if (!ciEquals(line.data() + i, nameEnd - i, "Icon")) return false;
    size_t eq = line.find('=', nameEnd);
    if (eq == std::string::npos) return false;
    size_t q1 = line.find('"', eq);
    if (q1 == std::string::npos) return false;
    size_t q2 = line.find('"', q1 + 1);
    if (q2 == std::string::npos || q2 == q1 + 1) return false;
    frxName = line.substr(q1 + 1, q2 - q1 - 1);
    // 引号后紧跟 `:0000`（四位十六进制，VB6 写死四位；位数多也按实际吃）
    size_t colon = q2 + 1;
    while (colon < line.size() && (line[colon] == ' ' || line[colon] == '\t')) colon++;
    if (colon >= line.size() || line[colon] != ':') return false;
    std::uint32_t v = 0;
    int digits = 0;
    for (size_t k = colon + 1; k < line.size(); k++) {
        int h = hexVal(line[k]);
        if (h < 0) break;
        v = v * 16u + static_cast<std::uint32_t>(h);
        digits++;
    }
    if (digits == 0) return false;
    offset = v;
    return true;
}

}  // namespace

bool loadFormIconFromFrm(const std::string& frmPathUtf8, FormIconData& out, std::string& why) {
    out = FormIconData{};
    if (frmPathUtf8.empty()) {
        why = "no .frm for the vbp's IconForm/Startup";
        return false;
    }
    std::ifstream in(utf8ToPath(frmPathUtf8), std::ios::binary);
    if (!in) {
        why = "cannot open " + frmPathUtf8;
        return false;
    }
    std::string frm((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());

    // 只吃窗体自己的属性块。`.frm` 的形状是 `Begin VB.Form <名>` → 属性若干行（其中孩子用
    // `Begin VB.<类型>` 包着）→ `End` —— 所以**第一条** Begin 就是窗体自己，要跨过它；
    // 第二条 Begin 才轮到孩子，孩子那一格的 Icon（MDIForm / UserControl 都有）不是应用图标。
    // `BeginProperty Font` 也在属性块里（窗体的 Icon 常排在它后面），所以 Begin 按整词认。
    std::string frxName;
    std::uint32_t offset = 0;
    bool inOwnBlock = false;
    size_t pos = 0;
    while (pos < frm.size()) {
        size_t nl = frm.find('\n', pos);
        std::string line = frm.substr(pos, nl == std::string::npos ? std::string::npos : nl - pos);
        pos = (nl == std::string::npos) ? frm.size() : nl + 1;
        size_t s = line.find_first_not_of(" \t\r\n");
        if (s == std::string::npos) continue;
        if (isControlBeginLine(line, s)) {
            if (inOwnBlock) break;
            inOwnBlock = true;
            continue;
        }
        if (parseIconPropLine(line, frxName, offset)) break;
        frxName.clear();
    }
    if (frxName.empty()) {
        why = "no Icon property on that form (" + frmPathUtf8 + ")";
        return false;
    }

    std::error_code ec;
    std::filesystem::path frx = std::filesystem::path(utf8ToPath(frmPathUtf8)).parent_path();
    frx /= utf8ToPath(frxName);
    ec.clear();
    if (!std::filesystem::exists(frx, ec)) {
        why = "frx not found: " + pathToUtf8(frx);
        return false;
    }
    std::ifstream fi(frx, std::ios::binary);
    if (!fi) {
        why = "cannot open frx";
        return false;
    }
    std::vector<unsigned char> blob((std::istreambuf_iterator<char>(fi)),
                                    std::istreambuf_iterator<char>());
    if (offset + 12 > blob.size()) {
        why = "record offset past end of frx";
        return false;
    }
    const unsigned char* rec = blob.data() + offset;
    std::uint32_t inner = rdLE32(rec + 8);
    if (!(rec[4] == 'l' && rec[5] == 't')) {
        why = "frx record is not an icon stream";
        return false;
    }
    if (inner < 22 || offset + 12 + inner > blob.size()) {
        // 空记录（`Icon = (None)` 那一档）与「声明长度越过文件尾」都走这里 —— 后者是第302刀
        // 量到的那条：最后一条记录的 lnext 是垃圾，声明长度不可信，所以只认吃得下的。
        why = (inner == 0) ? "empty icon record" : "icon record truncated";
        return false;
    }
    const unsigned char* ico = rec + 12;
    if (rdLE16(ico) != 0 || rdLE16(ico + 2) != 1) {
        why = "not an ICONDIR";
        return false;
    }
    std::uint32_t count = rdLE16(ico + 4);
    if (count == 0 || count > 256 || 6 + count * 16 > inner) {
        why = "bad icon directory";
        return false;
    }
    out.recordPath = pathToUtf8(frx);
    for (std::uint32_t i = 0; i < count; i++) {
        const unsigned char* e = ico + 6 + i * 16;
        std::uint32_t size = rdLE32(e + 8);
        std::uint32_t ofs = rdLE32(e + 12);
        if (size == 0 || ofs < 6 + count * 16 || ofs + size > inner) {
            out = FormIconData{};
            why = "icon frame " + std::to_string(i) + " points outside the record";
            return false;
        }
        FormIconFrame fr;
        std::memcpy(fr.dir, e, 12);
        fr.image.assign(ico + ofs, ico + ofs + size);
        out.frames.push_back(std::move(fr));
    }
    why = "ok";
    return !out.frames.empty();
}

bool buildFormIconResources(const FormIconData& icon,
                            std::vector<std::vector<unsigned char>>& framesOut,
                            std::vector<unsigned int>& idsOut,
                            std::vector<unsigned char>& groupOut) {
    framesOut.clear();
    idsOut.clear();
    groupOut.clear();
    if (icon.frames.empty() || icon.frames.size() > 256) return false;
    // GRPICONDIR：保留 0 / 类型 1 / 帧数；每帧 = 目录项前 12 字节 + **资源号**（取代 .ICO 里的偏移）
    groupOut.push_back(0x00);
    groupOut.push_back(0x00);
    groupOut.push_back(0x01);
    groupOut.push_back(0x00);
    groupOut.push_back(static_cast<unsigned char>(icon.frames.size() & 0xFF));
    groupOut.push_back(static_cast<unsigned char>((icon.frames.size() >> 8) & 0xFF));
    unsigned int nextId = kFormIconImageIdBase;
    for (const FormIconFrame& fr : icon.frames) {
        groupOut.insert(groupOut.end(), fr.dir, fr.dir + 12);
        groupOut.push_back(static_cast<unsigned char>(nextId & 0xFF));
        groupOut.push_back(static_cast<unsigned char>((nextId >> 8) & 0xFF));
        framesOut.push_back(fr.image);
        idsOut.push_back(nextId);
        nextId++;
    }
    return true;
}

}  // namespace vb6c3
