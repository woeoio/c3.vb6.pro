// cgen_base_generate_epilogue.inc - generate 函数体片段：保存生成结果与 COM 接口 typedef 声明
// 由 src/backend/cgen_base.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 1102~1152 行，逐行未改。

    // Delegate (tB 扩展): 所有过程体已发射, 桩定义落地 (原型在 decl pass, 定义可后置)
    emitDelegateThunks();

    // AddressOf 回调桩 (账 #184): 与委托桩同一批后置 —— 桩体要转调的过程定义此时已发完。
    emitAddressOfThunkDefs();

    // 类虚表 (tB B08d): 表实例也后置在这里 —— 表项可能是本类的 Private 过程, 那份定义是
    // `static` 且没有前置原型 (只有自家 .h 里 Public/Friend 那份有), 放前面就是未声明标识符。
    emitClassVirtImpl(module);

    // 构造链入口 (tB B09): 同样后置 —— 它桥接的 Class_Initialize 是 `static` 定义, 之前没有原型。
    emitChainInitDef(module);

    // 保存生成结果
    header_ = h_.str();
    source_ = c_.str();

    // Fix 178: 本模块请求到的「含所有权成员 UDT」深拷贝函数落地。必须在 body pass
    // 之后才知道用到了哪些类型, 而 C 又要求先定义后使用 —— 故插到 .c 自身
    // `#include "<Base>.h"` 之后 (先于所有 #define 与函数体)。static: 每个 .c 各一份,
    // 不与其他模块的同名副本冲突。
    if (!udtCopyRequested_.empty()) {
        std::string udtCopyBlock = emitUdtCopyBlock();
        if (!udtCopyBlock.empty()) {
            // 落在文件顶部**连续 include 段**之后。只跟第一行 include 走不够:
            // 标准模块之间的 include 发射在 .c 里 (c_open), 跨模块 UDT 在那一刻
            // 仍是不完整类型 → C2236/C2059。
            std::string incMarker = "#include \"" + baseName_ + ".h\"";
            size_t ip = source_.find(incMarker);
            size_t at = std::string::npos;
            if (ip != std::string::npos) {
                size_t nl = source_.find('\n', ip);
                at = (nl == std::string::npos) ? source_.size() : nl + 1;
                while (at < source_.size()) {
                    size_t eol = source_.find('\n', at);
                    if (eol == std::string::npos) eol = source_.size();
                    size_t b = source_.find_first_not_of(" \t\r", at);
                    if (b == std::string::npos || b > eol) break;
                    if (source_.compare(b, 8, "#include") != 0) break;
                    at = eol + 1;
                }
            }
            if (at != std::string::npos) source_.insert(at, udtCopyBlock);
            else source_.insert(0, udtCopyBlock);
        }
    }

    // P6.3: 如果使用了COM接口类型, 在头文件中插入typedef前向声明
    // vb6_ComIface_<Name> 是不透明结构体, 仅用作类型化指针
    if (!usedComIfaceTypes_.empty() || !usedVb6IfaceTypes_.empty() || !usedClassTypes_.empty() || !usedUdtTypes_.empty()) {
        std::string typedefBlock;
        if (!usedComIfaceTypes_.empty()) {
            typedefBlock += "\n// P6.3: COM接口类型前向声明 (早期绑定)\n";
            for (const auto& ifaceName : usedComIfaceTypes_) {
                typedefBlock += "typedef struct vb6_ComIface_" + ifaceName + " vb6_ComIface_" + ifaceName + ";\n";
            }
        }
        if (!usedVb6IfaceTypes_.empty()) {
            typedefBlock += "\n// P6.4: VB6接口类型前向声明 (Implements)\n";
            for (const auto& ifaceName : usedVb6IfaceTypes_) {
                typedefBlock += "typedef struct vb6_vtbl_" + ifaceName + " vb6_vtbl_" + ifaceName + ";\n";
                typedefBlock += "typedef struct vb6_iface_" + ifaceName + " vb6_iface_" + ifaceName + ";\n";
            }
        }
        // Fix 010: VB6类类型前向声明
        // C11允许typedef重定义, 所以无条件发出所有使用到的类类型前向声明
        // 解决: (1) 当前模块类struct中使用UDT类型 (2) 循环#include导致的类型不可见
        if (!usedClassTypes_.empty()) {
            typedefBlock += "\n// VB6类类型前向声明\n";
            for (const auto& clsName : usedClassTypes_) {
                typedefBlock += "typedef struct vb6_cls_" + clsName + " vb6_cls_" + clsName + ";\n";
            }
        }
        // Fix 010: VB6 UDT类型前向声明
        // C11允许typedef重定义, 无条件发出 — 解决类struct中UDT类型先使用后定义的问题
        if (!usedUdtTypes_.empty()) {
            typedefBlock += "\n// VB6 UDT类型前向声明\n";
            for (const auto& udtName : usedUdtTypes_) {
                typedefBlock += "typedef struct vb6_type_" + udtName + " vb6_type_" + udtName + ";\n";
            }
        }
        // 在 #include "vb6rtl.h" 之后插入
        std::string marker = "#include \"vb6rtl.h\"";
        size_t pos = header_.find(marker);
        if (pos != std::string::npos) {
            pos = header_.find('\n', pos);  // 找到行尾
            if (pos != std::string::npos) {
                header_.insert(pos + 1, typedefBlock);
            }
        }
    }

    currentModule_ = nullptr;
    return diag_.errorCount() == 0;
