// vb6forms_uc 拆分片段 uc_debug_dib (函数体)
// 内容 = 拆分前 vb6forms_uc.c 第 308~369 行，逐行未改（结构体除外）
// 在函数体内 #include；路径一律用 basename
// 注: vb6_UCDib 结构体已上移到 vb6forms_uc_internal.h（uc_host_window.c 也要用）
// 注: vb6_ucDumpSeq / vb6_uc_dibCreate / dibSaveBmp / dibDestroy 已去 static 供跨单元使用

int32_t vb6_ucDumpSeq = 0;

void vb6_uc_dibCreate(vb6_UCDib* d, HDC refDC, int32_t w, int32_t h) {
    BITMAPINFO bi;
    memset(d, 0, sizeof(*d));
    d->w = w;
    d->h = h;
    d->stride = ((w * 3 + 3) & ~3);
    d->memDC = CreateCompatibleDC(refDC);
    memset(&bi, 0, sizeof(bi));
    bi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bi.bmiHeader.biWidth = w;
    bi.bmiHeader.biHeight = h;   /* bottom-up: 与 BMP 默认行序一致, 可直接转存 */
    bi.bmiHeader.biPlanes = 1;
    bi.bmiHeader.biBitCount = 24;
    bi.bmiHeader.biCompression = BI_RGB;
    bi.bmiHeader.biSizeImage = (DWORD)(d->stride * h);
    d->bmp = CreateDIBSection(refDC, &bi, DIB_RGB_COLORS, &d->bits, NULL, 0);
    if (d->bmp) d->oldBmp = (HBITMAP)SelectObject(d->memDC, d->bmp);
}

void vb6_uc_dibSaveBmp(const vb6_UCDib* d, const char* path) {
    if (!d->bits) return;
    FILE* f = fopen(path, "wb");
    if (!f) return;
    uint32_t imgSize = (uint32_t)(d->stride * d->h);
    unsigned char hdr[54];
    memset(hdr, 0, sizeof(hdr));
    hdr[0] = 'B'; hdr[1] = 'M';
    uint32_t fileSize = 54u + imgSize;
    memcpy(hdr + 2, &fileSize, 4);
    uint32_t offBits = 54u;
    memcpy(hdr + 10, &offBits, 4);
    uint32_t hdrSize = 40u;
    memcpy(hdr + 14, &hdrSize, 4);
    int32_t W = d->w, H = d->h;
    memcpy(hdr + 18, &W, 4);
    memcpy(hdr + 22, &H, 4);
    uint16_t planes = 1, bpp = 24;
    memcpy(hdr + 26, &planes, 2);
    memcpy(hdr + 28, &bpp, 2);
    memcpy(hdr + 34, &imgSize, 4);
    fwrite(hdr, 1, 54, f);
    fwrite(d->bits, 1, (size_t)imgSize, f);
    fclose(f);
}

void vb6_uc_dibDestroy(vb6_UCDib* d) {
    if (d->memDC) {
        if (d->oldBmp) SelectObject(d->memDC, d->oldBmp);
        DeleteDC(d->memDC);
    }
    if (d->bmp) DeleteObject(d->bmp);
}
