// cgen_util_dllentry_prelude.inc - generateDllEntry 函数体片段：局部结构体 CoClassInfo 与哈希/格式化辅助 lambda
// 由 src/backend/cgen_util.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 12~72 行，逐行未改。

    CodeEmitter entry;
    dllProgId_ = progId;

    // 收集所有instancing >= PublicNotCreatable的类模块
    struct CoClassInfo {
        std::string moduleName;
        std::string clsidStr;
        std::string progId;
        std::vector<std::string> publicMethodNames;
        std::vector<Symbol*> publicMethodSyms;
        // Fix 099: 显式 Public 数据字段 (真 VB6 暴露为 Property Get/Let).
        // 与 publicMethodNames 分开存: 字段不是 Symbol, 只有名字 + dispid,
        // 其读写访问器由类模块自身 .c 提供 (见 cgen_com.cpp emitClassFieldAccessors).
        std::vector<std::string> publicFields;
        std::vector<int32_t> publicFieldDispids;  // 与 publicFields 平行; 0=未分配
        std::vector<std::string> implementsNames;  // P12.1: Implements接口列表
        // P6.6.4: 事件源信息
        std::vector<std::string> eventNames;       // 事件名称列表
        std::string defaultIfaceIid;              // 默认dispinterface IID (用于早绑定QI)
        std::string sourceIfaceIid;               // source dispinterface IID
        // ai/022 B13b: 这个类的 CoClass **块名**档 ProgID ("<工程>.<块名>")。
        // 只在块是手写的（块名 != 类模块名）且 [ComCreatable(True)] 时非空 ——
        // 折算记录（C04）的名字恒等于类模块名，天然不会产生第二行。
        std::string groupProgId;
        // ai/022 B16: 这个类有没有"新式接口"那一面（薄指针 + 包装器放手桩）。
        // 判据必须是 cgen_iface_vtbl 的发码判据本身（ivLookupIface 命中），不能拿
        // implementsNames 非空代替 —— legacy `Implements` 的类模块同样有名字，但
        // 它那条路一个 vb6_iv_* 符号都不发，照抄会 LNK2019。
        bool hasNewStyleIface = false;
        /* §B217 (第一百一十八刀): 这一行能不能对外 CoCreate。缺省档（Instancing 缺省 /
           PublicNotCreatable）为 false —— 行照发（装箱要靠它找 desc），但 progId/clsid
           两格发空：DllRegisterServer/DllUnregisterServer 跳过、ClassFactory 匹配不上。 */
        bool registerable = true;
    };
    std::vector<CoClassInfo> coClasses;
    // 组名档多出的行数（g_vb6_coclasses[] 的总行数 = coClasses.size() + extraGroupRows）
    int extraGroupRows = 0;

    // P12.1: Hex格式化辅助 (IID常量生成用)
    auto StringFormatHex2 = [](uint8_t v) -> std::string {
        char buf[4]; snprintf(buf, sizeof(buf), "%02X", v); return buf;
    };
    auto StringFormatHex4 = [](uint16_t v) -> std::string {
        char buf[6]; snprintf(buf, sizeof(buf), "%04X", v); return buf;
    };
    auto StringFormatHex8 = [](uint32_t v) -> std::string {
        char buf[10]; snprintf(buf, sizeof(buf), "%08X", v); return buf;
    };

    auto generateClsid = [](const std::string& name) -> std::string {
        uint32_t h1 = 0x811c9dc5, h2 = 0x01000193, h3 = 0xabcd1234, h4 = 0x5678ef01;
        for (char c : name) {
            h1 ^= (uint32_t)(unsigned char)c; h1 *= 0x01000193;
            h2 ^= (uint32_t)(unsigned char)c; h2 *= 0x01000193;
            h3 ^= (uint32_t)(unsigned char)c; h3 *= 0x01000193;
            h4 ^= (uint32_t)(unsigned char)c; h4 *= 0x01000193;
        }
        char buf[64];
        snprintf(buf, sizeof(buf), "{%08X-%04X-%04X-%04X-%04X%08X}",
                 h1, (h2 >> 16) & 0xFFFF, h2 & 0xFFFF | 0x4000,
                 (h3 >> 16) & 0xFFFF | 0x8000,
                 h3 & 0xFFFF, h4);
        return std::string(buf);
    };

    // IID/CLSID generation helpers
    auto generateIid = [](const std::string& name) -> std::string {
        uint32_t h1 = 0xa1b2c3d4, h2 = 0xe5f60718, h3 = 0x9a0b1c2d, h4 = 0x3e4f5061;
        for (char c : name) {
            h1 ^= (uint32_t)(unsigned char)c; h1 *= 0x01000193;
            h2 ^= (uint32_t)(unsigned char)c; h2 *= 0x01000193;
            h3 ^= (uint32_t)(unsigned char)c; h3 *= 0x01000193;
            h4 ^= (uint32_t)(unsigned char)c; h4 *= 0x01000193;
        }
        char buf[64];
        snprintf(buf, sizeof(buf), "{%08X-%04X-%04X-%04X-%04X%08X}",
                 h1, (h2 >> 16) & 0xFFFF, (h2 & 0xFFFF) | 0x4000,
                 (h3 >> 16) & 0xFFFF | 0x8000,
                 h3 & 0xFFFF, h4);
        return std::string(buf);
    };