// builtin_consts_ext.inc - registerBuiltins 的扩展常量注册段（函数体片段）
// 由 src/semantics/semantic_analyzer_builtin.cpp 拆出（2026-09-17），在函数体内被
// #include，非独立文件。内容 = 原文件第 298~503 行，逐行未改（P21-25 / P22 / P23-04）。
// 依赖上一片段定义的 addConst / addStrConst，必须排在它之后 include。


    // P22: Shell constants (for Shell function window style)
    addConst("vbHide", Vb6Type::Long, 0);
    addConst("vbNormalFocus", Vb6Type::Long, 1);
    addConst("vbMinimizedFocus", Vb6Type::Long, 2);
    addConst("vbMaximizedFocus", Vb6Type::Long, 3);
    addConst("vbNormalNoFocus", Vb6Type::Long, 4);
    addConst("vbMinimizedNoFocus", Vb6Type::Long, 6);
    // P22: Date format constants (for FormatDateTime)
    addConst("vbGeneralDate", Vb6Type::Long, 0);
    addConst("vbLongDate", Vb6Type::Long, 1);
    addConst("vbShortDate", Vb6Type::Long, 2);
    addConst("vbLongTime", Vb6Type::Long, 3);
    addConst("vbShortTime", Vb6Type::Long, 4);
    // P22: System color constants (OLE system colors)
    addConst("vbScrollBars", Vb6Type::Long, -2147483648LL);
    addConst("vbDesktop", Vb6Type::Long, -2147483647LL);
    addConst("vbActiveTitleBar", Vb6Type::Long, -2147483646LL);
    addConst("vbInactiveTitleBar", Vb6Type::Long, -2147483645LL);
    addConst("vbMenu", Vb6Type::Long, -2147483644LL);
    addConst("vbWindowBackground", Vb6Type::Long, -2147483643LL);
    addConst("vbWindowFrame", Vb6Type::Long, -2147483642LL);
    addConst("vbMenuText", Vb6Type::Long, -2147483641LL);
    addConst("vbTitleBarText", Vb6Type::Long, -2147483639LL);
    // Fix 131: vbWindowTitle (0x80000008) 缺失, vbWindowText 原误为 ...640
    addConst("vbWindowTitle", Vb6Type::Long, -2147483640LL);
    addConst("vbWindowText", Vb6Type::Long, -2147483639LL);
    addConst("vbActiveBorder", Vb6Type::Long, -2147483638LL);
    addConst("vbInactiveBorder", Vb6Type::Long, -2147483637LL);
    addConst("vbApplicationWorkspace", Vb6Type::Long, -2147483636LL);
    addConst("vbHighlight", Vb6Type::Long, -2147483635LL);
    addConst("vbHighlightText", Vb6Type::Long, -2147483634LL);
    addConst("vbButtonFace", Vb6Type::Long, -2147483633LL);
    addConst("vbButtonShadow", Vb6Type::Long, -2147483632LL);
    addConst("vbGrayText", Vb6Type::Long, -2147483631LL);
    addConst("vbButtonText", Vb6Type::Long, -2147483630LL);
    addConst("vbInactiveCaptionText", Vb6Type::Long, -2147483629LL);
    addConst("vb3DHighlight", Vb6Type::Long, -2147483628LL);
    addConst("vb3DDKShadow", Vb6Type::Long, -2147483627LL);
    addConst("vb3DLight", Vb6Type::Long, -2147483626LL);
    // Fix 131: 原表 vb3DFace=-2147483625 / vb3DShadow=-2147483624 整体错位 ——
    // 那两个值其实是 vbInfoText/vbInfoBackground (淡黄), 导致仪表 Caption2
    // "Band/Visitors/Volume" 渲染成淡黄色而非 vb3DShadow 深灰。
    addConst("vbInfoText", Vb6Type::Long, -2147483625LL);       // 0x80000017
    addConst("vbInfoBackground", Vb6Type::Long, -2147483624LL); // 0x80000018
    addConst("vb3DFace", Vb6Type::Long, -2147483633LL);         // 0x8000000F = vbButtonFace
    addConst("vb3DShadow", Vb6Type::Long, -2147483632LL);       // 0x80000010 = vbButtonShadow
    // P22: IME status constants
    addConst("vbIMEModeNoControl", Vb6Type::Long, 0);
    addConst("vbIMEModeOn", Vb6Type::Long, 1);
    addConst("vbIMEModeOff", Vb6Type::Long, 2);
    addConst("vbIMEModeDisable", Vb6Type::Long, 3);
    addConst("vbIMEModeHiragana", Vb6Type::Long, 4);
    addConst("vbIMEModeKatakana", Vb6Type::Long, 5);
    addConst("vbIMEModeKatakanaHalf", Vb6Type::Long, 6);
    addConst("vbIMEModeAlphaFull", Vb6Type::Long, 7);
    addConst("vbIMEModeAlpha", Vb6Type::Long, 8);
    addConst("vbIMEModeHangulFull", Vb6Type::Long, 9);
    addConst("vbIMEModeHangul", Vb6Type::Long, 10);

    // P22-09: Printer constants
    addConst("vbPRORPortrait", Vb6Type::Long, 1);        // Printer orientation: Portrait
    addConst("vbPRORLandscape", Vb6Type::Long, 2);       // Printer orientation: Landscape
    addConst("vbPRPQDraft", Vb6Type::Long, -1);          // Print quality: Draft
    addConst("vbPRPQLow", Vb6Type::Long, -2);            // Print quality: Low
    addConst("vbPRPQMedium", Vb6Type::Long, -3);         // Print quality: Medium
    addConst("vbPRPQHigh", Vb6Type::Long, -4);           // Print quality: High
    addConst("vbPRCMMillimeters", Vb6Type::Long, 1);     // Page scale: Millimeters
    addConst("vbPRCMCentimeters", Vb6Type::Long, 2);     // Page scale: Centimeters
    addConst("vbPRCMInches", Vb6Type::Long, 3);          // Page scale: Inches
    addConst("vbPRCMCharacters", Vb6Type::Long, 4);      // Page scale: Characters
    addConst("vbPRBPSingle", Vb6Type::Long, 1);          // Binary performation: Single
    addConst("vbPRBPSDouble", Vb6Type::Long, 2);         // Binary performation: Double
    addConst("vbPRBPTriple", Vb6Type::Long, 3);          // Binary performation: Triple
    addConst("vbPRDPHorizontal", Vb6Type::Long, 2);      // Duplex: Horizontal (DMDUP_HORIZONTAL) — Fix 214: 原记 1，错
    addConst("vbPRDPVertical", Vb6Type::Long, 3);        // Duplex: Vertical (DMDUP_VERTICAL) — Fix 214: 原记 2，错

    // Fix 214: Printer ColorMode / Duplex 缺族 + 订正
    //
    // 上面 P22-09 的 Duplex 两枚把值写错了 (Horizontal 记 1 / Vertical 记 2) —— 手册是
    // 一套 0/1/2/3 的枚举，缺了 **0 (PrinterDefault)** 与 **1 (Simplex)** 两个位置，
    // 所以整族往前挪了一格。逐值对照 Win32 DEVMODE.dmDuplex 也是同一个序:
    //   DMDUP_SIMPLEX=1, DMDUP_HORIZONTAL=2, DMDUP_VERTICAL=3。
    // 仓库内独立佐证: tests/dbgdlg/cDlg.cls 的 `cdeDuplexConstants` 枚举写的就是
    //   vbPRDPPrinterDefault=0, vbPRDPSimplex=1, vbPRDPHorizontal=2, vbPRDPVertical=3。
    // 另: ColorMode 整族 (vbPRCM*) 之前完全没进表 —— 暴露面是 ComCtlsDemo 的
    //   Builds/CommonDialog.cls: 它拿 vbPRCMMonochrome / vbPRCMColor 当**数组上界**用,
    //   语义层不认识 ⇒ 发码层发裸标识符 ⇒ C2065 每 TU 3 条 × 111 TU = 333 条,
    //   外加同址 C2057「应输入常量表达式」333 条, cl 到 C1003 提前收摊。
    // 值取自 MSDN "Printer Object Constants" (aa243034) ColorMode / Duplex 两张表。
    // ⚠ 注意与上面 P22-09 几条**名字可疑**的条目区分 (vbPRCMMillimeters 等 7 枚,
    //   VB6 的 Printer 常数表里没有这些名字; 真正同名族是 unprefixed 的 vbMillimeters
    //   ~ vbCentimeters, 已在本文件 :262 一带有正确值)。本轮不动它们, 另案清理。
    addConst("vbPRCMMonochrome", Vb6Type::Long, 1);      // ColorMode: Monochrome (DMCOLOR_MONOCHROME)
    addConst("vbPRCMColor", Vb6Type::Long, 2);           // ColorMode: Color (DMCOLOR_COLOR)
    addConst("vbPRDPSimplex", Vb6Type::Long, 1);         // Duplex: Simplex (DMDUP_SIMPLEX)

    // Fix 211: PrinterObjectConstants 的 PaperSize (vbPRPS*) 与 PaperBin (vbPRBN*) 两族
    // 从未进过这张表 —— 上面 P22-09 那一批只铺了 Orientation / PrintQuality / PageScale /
    // BinaryPerforation / Duplex 五族，PaperSize 与 PaperBin 漏了。
    // 暴露面: ComCtlsDemo 的 Builds/CommonDialog.cls 拿 PaperSize 常量当**数组上界/枚举
    // 初值**用，语义层不认识 ⇒ 发码层照名发出裸标识符 ⇒ 生成的 C 里 C2065「未声明的
    // 标识符」+ C2057「应输入常量表达式」，一个工程 5660+5550 条，cl 到 C1003 提前收摊。
    // 这一格是「三处答案」哨兵 (scripts/check_builtin_const_authority.ps1) 的 C2 条:
    // 常量必须住在**表**里，不许只在发码折叠里再写第二份 —— 所以只往这里加，不动折叠。
    // 值取自 MSDN "Printer Object Constants" / PaperSize 属性手册 (docs 同口径)。
    // ⚠ EnvC5=28 而 EnvC3=29 / EnvC4=30 —— **不是笔误**，手册就是这个序 (DMPAPER_ENV_C5=28)，
    //   别按字面把它"理顺"成 27/28/29。
    addConst("vbPRPSLetter", Vb6Type::Long, 1);            // Letter, 8 1/2 x 11 in
    addConst("vbPRPSLetterSmall", Vb6Type::Long, 2);       // Letter Small
    addConst("vbPRPSTabloid", Vb6Type::Long, 3);           // Tabloid, 11 x 17 in
    addConst("vbPRPSLedger", Vb6Type::Long, 4);            // Ledger, 17 x 11 in
    addConst("vbPRPSLegal", Vb6Type::Long, 5);             // Legal, 8 1/2 x 14 in
    addConst("vbPRPSStatement", Vb6Type::Long, 6);         // Statement, 5 1/2 x 8 1/2 in
    addConst("vbPRPSExecutive", Vb6Type::Long, 7);         // Executive, 7 1/2 x 10 1/2 in
    addConst("vbPRPSA3", Vb6Type::Long, 8);                // A3, 297 x 420 mm
    addConst("vbPRPSA4", Vb6Type::Long, 9);                // A4, 210 x 297 mm
    addConst("vbPRPSA4Small", Vb6Type::Long, 10);          // A4 Small
    addConst("vbPRPSA5", Vb6Type::Long, 11);               // A5, 148 x 210 mm
    addConst("vbPRPSB4", Vb6Type::Long, 12);               // B4, 250 x 354 mm
    addConst("vbPRPSB5", Vb6Type::Long, 13);               // B5, 182 x 257 mm
    addConst("vbPRPSFolio", Vb6Type::Long, 14);            // Folio, 8 1/2 x 13 in
    addConst("vbPRPSQuarto", Vb6Type::Long, 15);           // Quarto, 215 x 275 mm
    addConst("vbPRPS10x14", Vb6Type::Long, 16);            // 10 x 14 in
    addConst("vbPRPS11x17", Vb6Type::Long, 17);            // 11 x 17 in
    addConst("vbPRPSNote", Vb6Type::Long, 18);             // Note, 8 1/2 x 11 in
    addConst("vbPRPSEnv9", Vb6Type::Long, 19);             // Envelope #9
    addConst("vbPRPSEnv10", Vb6Type::Long, 20);            // Envelope #10
    addConst("vbPRPSEnv11", Vb6Type::Long, 21);            // Envelope #11
    addConst("vbPRPSEnv12", Vb6Type::Long, 22);            // Envelope #12
    addConst("vbPRPSEnv14", Vb6Type::Long, 23);            // Envelope #14
    addConst("vbPRPSCSheet", Vb6Type::Long, 24);           // C size sheet
    addConst("vbPRPSDSheet", Vb6Type::Long, 25);           // D size sheet
    addConst("vbPRPSESheet", Vb6Type::Long, 26);           // E size sheet
    addConst("vbPRPSEnvDL", Vb6Type::Long, 27);            // Envelope DL, 110 x 220 mm
    addConst("vbPRPSEnvC5", Vb6Type::Long, 28);            // Envelope C5, 162 x 229 mm
    addConst("vbPRPSEnvC3", Vb6Type::Long, 29);            // Envelope C3, 324 x 458 mm
    addConst("vbPRPSEnvC4", Vb6Type::Long, 30);            // Envelope C4, 229 x 324 mm
    addConst("vbPRPSEnvC6", Vb6Type::Long, 31);            // Envelope C6, 114 x 162 mm
    addConst("vbPRPSEnvC65", Vb6Type::Long, 32);           // Envelope C65, 114 x 229 mm
    addConst("vbPRPSEnvB4", Vb6Type::Long, 33);            // Envelope B4, 250 x 353 mm
    addConst("vbPRPSEnvB5", Vb6Type::Long, 34);            // Envelope B5, 176 x 250 mm
    addConst("vbPRPSEnvB6", Vb6Type::Long, 35);            // Envelope B6, 176 x 125 mm
    addConst("vbPRPSEnvItaly", Vb6Type::Long, 36);         // Envelope, 110 x 230 mm
    addConst("vbPRPSEnvMonarch", Vb6Type::Long, 37);       // Envelope Monarch
    addConst("vbPRPSEnvPersonal", Vb6Type::Long, 38);      // Envelope, 3 5/8 x 6 1/2 in
    addConst("vbPRPSFanfoldUS", Vb6Type::Long, 39);        // U.S. Standard Fanfold
    addConst("vbPRPSFanfoldStdGerman", Vb6Type::Long, 40); // German Standard Fanfold
    addConst("vbPRPSFanfoldLglGerman", Vb6Type::Long, 41); // German Legal Fanfold
    addConst("vbPRPSUser", Vb6Type::Long, 256);            // User-defined (DMPAPER_USER)
    addConst("vbPRBNUpper", Vb6Type::Long, 1);             // Upper bin
    addConst("vbPRBNLower", Vb6Type::Long, 2);             // Lower bin
    addConst("vbPRBNMiddle", Vb6Type::Long, 3);            // Middle bin
    addConst("vbPRBNManual", Vb6Type::Long, 4);            // Manual insertion per sheet
    addConst("vbPRBNEnvelope", Vb6Type::Long, 5);          // Envelope feeder
    addConst("vbPRBNEnvManual", Vb6Type::Long, 6);         // Envelope feeder, manual insertion
    addConst("vbPRBNAuto", Vb6Type::Long, 7);              // (Default) current default bin
    addConst("vbPRBNTractor", Vb6Type::Long, 8);           // Tractor feeder
    addConst("vbPRBNSmallFmt", Vb6Type::Long, 9);          // Small paper feeder
    addConst("vbPRBNLargeFmt", Vb6Type::Long, 10);         // Large paper bin
    addConst("vbPRBNLargeCapacity", Vb6Type::Long, 11);    // Large capacity feeder
    addConst("vbPRBNCassette", Vb6Type::Long, 14);         // Attached cassette cartridge

    // P22-09: Other missing constants
    addConst("vbUseSystem", Vb6Type::Long, 0);            // VB6 = 0（用 NLS API 设置）；原记 -1 是把它和 vbUseCompareOption(-1) 混了
    addConst("vbUseCompareOption", Vb6Type::Long, -1);   // Use Option Compare setting
    addConst("Win16", Vb6Type::Long, 0);                  // Obsolete: always False
    addConst("Win32", Vb6Type::Long, -1);                 // True on 32-bit Windows
    addConst("vbDot", Vb6Type::Long, 46);                 // "." character code


    // P22-09: MsgBox additional constants
    addConst("vbMsgBoxHelpButton", Vb6Type::Long, 16384); // &H4000


    // --- P23-04: Low-frequency constants (85 additions) ---

    // VbCallType (CallType)
    addConst("vbMethod", Vb6Type::Long, 2);
    addConst("vbGet", Vb6Type::Long, 3);
    addConst("vbLet", Vb6Type::Long, 1);
    addConst("vbSet", Vb6Type::Long, 4);

    // VbCalendar

    // VbQueryClose

    // Clipboard format constants
    addConst("vbCFText", Vb6Type::Long, 1);
    addConst("vbCFBitmap", Vb6Type::Long, 2);
    addConst("vbCFMetafile", Vb6Type::Long, 3);
    addConst("vbCFDIB", Vb6Type::Long, 8);
    addConst("vbCFPalette", Vb6Type::Long, 9);
    addConst("vbCFRTF", Vb6Type::Long, -16639);       // &HFFFFBF01
    addConst("vbCFEMetafile", Vb6Type::Long, 14);

    // DriveType constants
    addConst("vbDriveTypeRemovable", Vb6Type::Long, 1);
    addConst("vbDriveTypeFixed", Vb6Type::Long, 2);
    addConst("vbDriveTypeNetwork", Vb6Type::Long, 3);
    addConst("vbDriveTypeCDRom", Vb6Type::Long, 4);
    addConst("vbDriveTypeRAMDisk", Vb6Type::Long, 5);

    // VbAppWinStyle
    addConst("vbAppWinStyleNormal", Vb6Type::Long, 1);
    addConst("vbAppWinStyleMinimize", Vb6Type::Long, 2);
    addConst("vbAppWinStyleMaximize", Vb6Type::Long, 3);

    // VbIMEStatus (distinct from IMEMode)
    addConst("vbIMEOn", Vb6Type::Long, 1);
    addConst("vbIMEOff", Vb6Type::Long, 0);
    addConst("vbIMEDisable", Vb6Type::Long, 2);
    addConst("vbIMEHiragana", Vb6Type::Long, 4);
    addConst("vbIMEKatakana", Vb6Type::Long, 5);
    addConst("vbIMEKatakanaHalf", Vb6Type::Long, 6);
    addConst("vbIMEAlphaFull", Vb6Type::Long, 7);
    addConst("vbIMEAlpha", Vb6Type::Long, 8);

    // Shape control constants
    addConst("vbShapeRectangle", Vb6Type::Long, 0);
    addConst("vbShapeSquare", Vb6Type::Long, 1);
    addConst("vbShapeOval", Vb6Type::Long, 2);
    addConst("vbShapeCircle", Vb6Type::Long, 3);
    addConst("vbShapeRoundedRectangle", Vb6Type::Long, 4);
    addConst("vbShapeRoundedSquare", Vb6Type::Long, 5);

    // BorderStyle (Shape/Line)
    addConst("vbTransparent", Vb6Type::Long, 0);
    addConst("vbBSSolid", Vb6Type::Long, 1);
    addConst("vbBSDash", Vb6Type::Long, 2);
    addConst("vbBSDot", Vb6Type::Long, 3);
    addConst("vbBSDashDot", Vb6Type::Long, 4);
    addConst("vbBSDashDotDot", Vb6Type::Long, 5);
    addConst("vbBSInsideSolid", Vb6Type::Long, 6);

    // Fix 214: 控件族常量的三处缺口 —— OLE 拖放 / Form·TextBox BorderStyle / DrawStyle /
    // MultiSelect。都是 ComCtlsDemo 在**赋值与比较**里逐处用到的真 VB6 内在常量, 表里没有
    // ⇒ 发码层不发常量值而发裸标识符 ⇒ C2065 (实测 vbOLEDragManual 41 处 / vbOLEDragAutomatic
    // 39 / vbBSNone 23 / vbMultiSelectNone 22 / vbOLEDropAutomatic 6 / vbInvisible 5 /
    // vbSolid 5 / vbMultiSelectSimple 4 / vbMultiSelectExtended 4)。
    // 值取自 MSDN: OLEDragMode 属性 (aa245845) / OLEDropMode 属性 (aa245847) /
    // BorderStyle Constants (ms855749) / DrawStyle 属性 (aa234894) / ListBox MultiSelect。
    // ⚠ 这几族**在手册里是分开的表**, 别按前缀归并:
    //   - OLEDragMode 与 OLEDropMode 起点不同 (Manual 在拖侧是 0、在放侧是 1);
    //   - `vbBS*` 有两族: Form/TextBox 的 (vbBSNone=0 …) 与 Shape/Line 的
    //     (vbTransparent=0, vbBSSolid=1 …), 上面那一组是后者;
    //   - DrawStyle 的 vbInvisible 是 **5** 不是 1 (1 是 vbDash)。
    //   - 放侧的 vbOLEDropNone=0 / vbOLEDropManual=1 已在 builtin_consts.inc:223 一带。
    addConst("vbOLEDragManual", Vb6Type::Long, 0);       // OLEDragMode: manual (OLEDrag method)
    addConst("vbOLEDragAutomatic", Vb6Type::Long, 1);    // OLEDragMode: automatic
    addConst("vbOLEDropAutomatic", Vb6Type::Long, 2);    // OLEDropMode: automatic accept
    addConst("vbBSNone", Vb6Type::Long, 0);              // BorderStyle (Form/TextBox/Frame): none
    addConst("vbSolid", Vb6Type::Long, 0);               // DrawStyle: solid (default)
    addConst("vbInvisible", Vb6Type::Long, 5);           // DrawStyle: transparent line
    addConst("vbMultiSelectNone", Vb6Type::Long, 0);     // ListBox/ListView MultiSelect: none
    addConst("vbMultiSelectSimple", Vb6Type::Long, 1);   // MultiSelect: simple
    addConst("vbMultiSelectExtended", Vb6Type::Long, 2); // MultiSelect: extended (Ctrl/Shift)

    // FillStyle constants
    addConst("vbFSSolid", Vb6Type::Long, 0);
    addConst("vbFSTransparent", Vb6Type::Long, 1);
    addConst("vbFSHorizontalLine", Vb6Type::Long, 2);
    addConst("vbFSVerticalLine", Vb6Type::Long, 3);
    addConst("vbFSUpwardDiagonal", Vb6Type::Long, 4);
    addConst("vbFSDownwardDiagonal", Vb6Type::Long, 5);
    addConst("vbFSCross", Vb6Type::Long, 6);
    addConst("vbFSDiagonalCross", Vb6Type::Long, 7);

    // MousePointer constants
    addConst("vbDefault", Vb6Type::Long, 0);
    addConst("vbArrow", Vb6Type::Long, 1);
    addConst("vbCrosshair", Vb6Type::Long, 2);
    addConst("vbIbeam", Vb6Type::Long, 3);
    addConst("vbIconPointer", Vb6Type::Long, 4);
    addConst("vbSizePointer", Vb6Type::Long, 5);
    addConst("vbSizeNESW", Vb6Type::Long, 6);
    addConst("vbSizeNS", Vb6Type::Long, 7);
    addConst("vbSizeNWSE", Vb6Type::Long, 8);
    addConst("vbSizeWE", Vb6Type::Long, 9);
    addConst("vbUpArrow", Vb6Type::Long, 10);
    addConst("vbHourglass", Vb6Type::Long, 11);
    addConst("vbNoDrop", Vb6Type::Long, 12);
    addConst("vbArrowHourglass", Vb6Type::Long, 13);
    addConst("vbArrowQuestion", Vb6Type::Long, 14);
    addConst("vbSizeAll", Vb6Type::Long, 15);
    addConst("vbCustom", Vb6Type::Long, 99);

    // Alignment constants
    addConst("vbLeftJustify", Vb6Type::Long, 0);
    addConst("vbRightJustify", Vb6Type::Long, 1);
    addConst("vbCenter", Vb6Type::Long, 2);

    // ScrollBar constants
    addConst("vbSBNone", Vb6Type::Long, 0);
    addConst("vbSBHorizontal", Vb6Type::Long, 1);
    addConst("vbSBVertical", Vb6Type::Long, 2);
    addConst("vbSBBoth", Vb6Type::Long, 3);

    // ScaleMode constants
    addConst("vbTwips", Vb6Type::Long, 1);
    addConst("vbPoints", Vb6Type::Long, 2);
    addConst("vbPixels", Vb6Type::Long, 3);
    addConst("vbCharacters", Vb6Type::Long, 4);
    addConst("vbInches", Vb6Type::Long, 5);
    addConst("vbMillimeters", Vb6Type::Long, 6);
    addConst("vbCentimeters", Vb6Type::Long, 7);
    addConst("vbHimetric", Vb6Type::Long, 8);
    addConst("vbContainerPosition", Vb6Type::Long, 9);
    addConst("vbContainerSize", Vb6Type::Long, 10);
    addConst("vbUser", Vb6Type::Long, 0);

    // WindowState constants
    addConst("vbMinimized", Vb6Type::Long, 1);
    addConst("vbMaximized", Vb6Type::Long, 2);

    // Fix 056: CheckBox constants
    addConst("vbUnchecked", Vb6Type::Long, 0);
    addConst("vbChecked", Vb6Type::Long, 1);
    addConst("vbGrayed", Vb6Type::Long, 2);

    // String constants (additional)
    addStrConst("vbNullString", "");
