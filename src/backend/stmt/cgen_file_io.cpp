#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_file_io.cpp: 文件 I/O 语句生成 (Open/Close/Print/Write/Input/Get/Put...) ---

// Print # / Write # 的单项文本出口。**类型优先**: 静态类型能答准就走实测 VB6 的
// 格式化 (vb6_File*Text); 只有类型信息不可得 (inferExprType 仍答 Unknown/ Object)
// 才回退旧 vb6_Str(int32_t)。String 那一档的分流由 fileOutArgIsBstr 一位主人答:
// ① 语义类型是 String ② 发射形状本身证明是 BSTR 槽 (见下方的注释)。旧的名字兜底表
// (34 名 C 出口名名单与旧 lambda) 已删; 删表前提与剩余边界见 docs/print_write_numeric.md §6.4/§7。
std::string CCodeGen::printWriteTextExpr(Expr& expr, const std::string& val, bool writeMode) {
    const std::string mode = writeMode ? "1" : "0";
    // Null/Empty 字面量与「明确是 Variant」的表达式必须走 Variant 档:
    // inferExprType 对 Null 字面量会答成整数型 (符号表里它没有独立类型), 若按数值发码
    // 就会把 vb6_VariantNull() 结构体强转成 int64 (实测编译错 C2440)。判据不猜:
    // ① 明确 Variant (isDefinitelyVariantExpr) ② 发射形状就是 RTL 的 Null/Empty 构造。
    if (isDefinitelyVariantExpr(expr) ||
        val == "vb6_VariantNull()" || val == "vb6_VariantEmpty()")
        return "vb6_FileVariantText(" + val + ", " + mode + ")";
    switch (inferExprType(expr)) {
        case Vb6Type::Byte:
        case Vb6Type::Integer:
        case Vb6Type::Long:
        case Vb6Type::SByte:
        case Vb6Type::UInteger:
        case Vb6Type::ULong:
        case Vb6Type::LongLong:
        case Vb6Type::ULongLong:
        case Vb6Type::LongPtr:
            return "vb6_FileLongText((int64_t)(" + val + "), " + mode + ")";
        case Vb6Type::Single:
            return "vb6_FileNumText((double)(" + val + "), 1, " + mode + ")";
        case Vb6Type::Double:
        case Vb6Type::Currency:
            return "vb6_FileNumText((double)(" + val + "), 0, " + mode + ")";
        case Vb6Type::Date:
            return "vb6_FileDateText((double)(" + val + "), " + mode + ")";
        case Vb6Type::Boolean:
            return "vb6_FileBoolText((int32_t)(" + val + "), " + mode + ")";
        case Vb6Type::Decimal:
            // Decimal 在 C 侧就是整只 vb6_VARIANT (cgen_base_type.cpp:56)。
            return "vb6_FileVariantText(" + val + ", " + mode + ")";
        case Vb6Type::Empty:
        case Vb6Type::Null:
        case Vb6Type::Variant:
            // "Variant" 是 inferExprType 的**兜底答案**, 不是断言: 走到这里说明
            // 开头那两问 (明确 Variant / RTL 的 Null/Empty 构造) 都没命中。只有
            // 发射形状本身是 RTL 里返回 vb6_VARIANT 的出口 (cExprIsVariant, 静态
            // 证据) 才按整只槽发; 其余 (IIf/Choose/Switch 的兜底答案、UDT 字段
            // 失配、外部类成员 …) 判不准, 保持旧路径, 不猜。
            // 形态清单与在 tests 语料里的频度见 docs/print_write_numeric.md。
            if (cExprIsVariant(val))
                return "vb6_FileVariantText(" + val + ", " + mode + ")";
            return "vb6_Str((int32_t)(" + val + "))";
        default:
            // 判不准的形态一律保持旧路径 (不猜); 清单见 docs/print_write_numeric.md。
            return "vb6_Str((int32_t)(" + val + "))";
    }
}

void CCodeGen::visit(OpenStmt& node) {
    // VB6: Open pathname For Mode [Access access] As #filenumber
    // C:   vb6_Open(pathname, mode, access, filenumber)
    emitExpr(*node.pathName);
    std::string pathName = std::move(lastExpr_);

    // OpenMode → 数值: Input=1, Output=2, Random=4, Append=8, Binary=16
    int32_t modeVal = 1;
    switch (node.mode) {
        case OpenMode::Input:   modeVal = 1; break;
        case OpenMode::Output:  modeVal = 2; break;
        case OpenMode::Random:  modeVal = 4; break;
        case OpenMode::Append:  modeVal = 8; break;
        case OpenMode::Binary:  modeVal = 16; break;
    }

    // OpenAccess → 数值: Read=1, Write=2, ReadWrite=3, Default=0
    int32_t accessVal = 0;
    switch (node.access) {
        case OpenAccess::Read:      accessVal = 1; break;
        case OpenAccess::Write:     accessVal = 2; break;
        case OpenAccess::ReadWrite: accessVal = 3; break;
        case OpenAccess::Default:   accessVal = 0; break;
    }

    emitExpr(*node.fileNumber);
    std::string fnum = std::move(lastExpr_);

    std::string recLen = "0";
    if (node.recordLength) {
        emitExpr(*node.recordLength);
        recLen = std::move(lastExpr_);
    }
    c_.emitLine("vb6_Open(" + pathName + ", " + std::to_string(modeVal) + ", " +
                std::to_string(accessVal) + ", " + fnum + ", " + recLen + ");");
}

void CCodeGen::visit(CloseStmt& node) {
    if (node.fileNumbers.empty()) {
        c_.emitLine("vb6_CloseAll();");
    } else {
        for (auto& fn : node.fileNumbers) {
            emitExpr(*fn);
            c_.emitLine("vb6_Close(" + lastExpr_ + ");");
        }
    }
}

// 账 #608 / 第331刀 的「一位主人」结构 + 第二步（§B146 收口）的两问判据：
//   ① 语义类型是 String（inferExprType；builtin 返回档由
//      cgen_builtin_return_type_table.inc 答出，表本身查 RTL 声明）；
//   ② 发射形状本身证明是 BSTR 槽（不是名字名单）：BSTR 构造/读取族、
//      String 数组元素、IIf 选中串支、静态 String 变量。
// 34 名 C 出口名名录已删：类型答不出时**不许猜**。
bool CCodeGen::fileOutArgIsBstr(const std::string& expr, Expr& argNode) const {
    if (inferExprType(argNode) == Vb6Type::String) return true;
    if (expr.compare(0, 8, "vb6_BSTR") == 0) return true;
    if (expr.compare(0, 11, "vb6_IIfBSTR") == 0) return true;
    if (expr.find("VB6_SA_AT(BSTR,") != std::string::npos) return true;
    std::string lower = expr;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
    return knownBstrVars_.count(lower) != 0;
}

// 「该以什么形状交给写文件的出口」：是 BSTR 就原样交，不是就按档转成 BSTR
// (Date→CStrDate 等; 与窗体面 Print 同一位主人)。
std::string CCodeGen::fileOutArgBstrExpr(const std::string& expr, Expr& argNode) {
    if (fileOutArgIsBstr(expr, argNode)) return expr;
    return wrapToBSTR(expr, argNode);
}

void CCodeGen::visit(PrintStmt& node) {
    if (node.isFormPrint) {
        // M22-Issue6: Form surface Print - "Print expr" in form module
        // Generate: vb6_Form_Print(formHwnd, bstrExpr) for each output item
        std::string formHwnd = isFormModule_ ? ("vb6_hwnd_" + cIdent(formName_)) : "NULL";
        for (auto& expr : node.outputList) {
            emitExpr(*expr);
            std::string val = lastExpr_;
            // Use wrapToBSTR for proper type conversion (Date→CStrDate, etc.)
            c_.emitLine("vb6_Form_Print(" + formHwnd + ", " + fileOutArgBstrExpr(val, *expr) + ");");
        }
        return;
    }

    // File I/O Print: Print #fnum, expr1; expr2; ...
    emitExpr(*node.fileNumber);
    std::string fnum = std::move(lastExpr_);

    if (node.outputList.empty()) {
        c_.emitLine("vb6_PrintText(" + fnum + ", NULL, 0);");
    } else {
        for (size_t i = 0; i < node.outputList.size(); ++i) {
            auto& expr = node.outputList[i];
            emitExpr(*expr);
            std::string val = lastExpr_;
            // 分隔符: 0=行尾, 1=';', 2=',' (AST 里逐项记录; 旧路径没有这个信息时按行尾)
            int sep = 0;
            if (i < node.separators.size()) {
                if (node.separators[i] == ';') sep = 1;
                else if (node.separators[i] == ',') sep = 2;
            }
            if (fileOutArgIsBstr(val, *expr)) {
                c_.emitLine("vb6_PrintText(" + fnum + ", " + val + ", " + std::to_string(sep) + ");");
            } else {
                // 数值/日期/布尔/Empty/Null: 按**实测 VB6 字节**产文本 (见 vb6rtl_file.c)。
                c_.emitLine("vb6_PrintText(" + fnum + ", " +
                            printWriteTextExpr(*expr, val, false) + ", " +
                            std::to_string(sep) + ");");
            }
        }
    }
}

void CCodeGen::visit(WriteStmt& node) {
    emitExpr(*node.fileNumber);
    std::string fnum = std::move(lastExpr_);

    // 第二步：与 PrintStmt 同一判据（两问，无名字名单；见 fileOutArgIsBstr 的注释）
    if (node.outputList.empty()) {
        c_.emitLine("vb6_WriteText(" + fnum + ", NULL, 0, 0);");
    } else {
        for (size_t i = 0; i < node.outputList.size(); ++i) {
            auto& expr = node.outputList[i];
            emitExpr(*expr);
            std::string val = lastExpr_;
            // Write # 项间恒为逗号 (实测: 源码里写 ';' 也输出 ','), 整条语句末尾一个 CRLF。
            int sep = (i + 1 < node.outputList.size()) ? 2 : 0;
            if (fileOutArgIsBstr(val, *expr)) {
                c_.emitLine("vb6_WriteText(" + fnum + ", " + val + ", 1, " +
                            std::to_string(sep) + ");");
            } else {
                c_.emitLine("vb6_WriteText(" + fnum + ", " +
                            printWriteTextExpr(*expr, val, true) + ", 0, " +
                            std::to_string(sep) + ");");
            }
        }
    }
}

void CCodeGen::visit(LineInputStmt& node) {
    // Line Input #fnum, varName
    emitExpr(*node.fileNumber);
    std::string fnum = std::move(lastExpr_);
    emitExpr(*node.varName);
    std::string varExpr = lastExpr_;
    c_.emitLine(varExpr + " = vb6_LineInput(" + fnum + ");");
}

void CCodeGen::visit(InputStmt& node) {
    // Input #fnum, var1, var2, ...
    emitExpr(*node.fileNumber);
    std::string fnum = std::move(lastExpr_);
    for (auto& var : node.varList) {
        emitExpr(*var);
        std::string varExpr = lastExpr_;
        c_.emitLine("vb6_Input(" + fnum + ", &" + varExpr + ");");
    }
}

void CCodeGen::visit(GetStmt& node) {
    // Get #fnum, [recnum], var
    emitExpr(*node.fileNumber);
    std::string fnum = std::move(lastExpr_);
    emitExpr(*node.varName);
    std::string varExpr = std::move(lastExpr_);
    // Whole 1D array target (`Get #f, , data`): every C3 array is a SafeArray1D
    // descriptor, so the payload is arr->data and the byte count is
    // count * elemSize. The old path emitted &var + sizeof(int32_t), which for a
    // descriptor variable wrote 4 bytes over the descriptor pointer slot.
    // (ND arrays are left on the old path; vb6_GetArray only knows 1D.)
    auto wholeArray1D = [&](const Expr* e) -> bool {
        if (!isWholeArrayRef(e)) return false;
        std::string nm;
        if (e->kind == ASTNodeKind::IdentifierExpr) {
            nm = Symbol::toLower(static_cast<const IdentifierExpr&>(*e).name);
        } else if (e->kind == ASTNodeKind::IndexOrCallExpr) {
            auto& ic = static_cast<const IndexOrCallExpr&>(*e);
            if (ic.callee && ic.callee->kind == ASTNodeKind::IdentifierExpr)
                nm = Symbol::toLower(static_cast<const IdentifierExpr&>(*ic.callee).name);
        }
        auto it = arrayDimCounts_.find(nm);
        return it == arrayDimCounts_.end() || it->second <= 1;
    };
    if (wholeArray1D(node.varName.get())) {
        std::string recnum = "0";
        if (node.recordNumber) { emitExpr(*node.recordNumber); recnum = std::move(lastExpr_); }
        c_.emitLine("vb6_GetArray(" + fnum + ", " + recnum + ", " + varExpr + ");");
        return;
    }
    // 确定变量大小: 根据变量名查询已知类型
    std::string varLower = varExpr;
    std::transform(varLower.begin(), varLower.end(), varLower.begin(), ::tolower);
    std::string bareName = varLower;
    std::string sizeExpr;
    if (knownBstrVars_.count(bareName)) {
        sizeExpr = "0";  // BSTR: varSize=0 让RTL层特殊处理
    } else if (knownDoubleVars_.count(bareName)) {
        sizeExpr = "sizeof(double)";
    } else if (knownLongVars_.count(bareName)) {
        sizeExpr = "sizeof(int32_t)";
    } else {
        sizeExpr = "sizeof(int32_t)";  // 默认: Long
    }
    if (node.recordNumber) {
        emitExpr(*node.recordNumber);
        std::string recnum = std::move(lastExpr_);
        c_.emitLine("vb6_Get(" + fnum + ", " + recnum + ", &" + varExpr + ", " + sizeExpr + ");");
    } else {
        // 无 recnum: 顺序读, recnum=0 表示当前位置
        c_.emitLine("vb6_Get(" + fnum + ", 0, &" + varExpr + ", " + sizeExpr + ");");
    }
}

void CCodeGen::visit(PutStmt& node) {
    // Put #fnum, [recnum], var
    emitExpr(*node.fileNumber);
    std::string fnum = std::move(lastExpr_);
    emitExpr(*node.varName);
    std::string varExpr = std::move(lastExpr_);
    // Whole 1D array source (`Put #f, , data`): every C3 array is a SafeArray1D
    // descriptor, so the payload is arr->data and the byte count is
    // count * elemSize. The old path emitted &var + sizeof(int32_t), which for a
    // descriptor variable wrote 4 bytes over the descriptor pointer slot.
    // (ND arrays are left on the old path; vb6_PutArray only knows 1D.)
    auto wholeArray1D = [&](const Expr* e) -> bool {
        if (!isWholeArrayRef(e)) return false;
        std::string nm;
        if (e->kind == ASTNodeKind::IdentifierExpr) {
            nm = Symbol::toLower(static_cast<const IdentifierExpr&>(*e).name);
        } else if (e->kind == ASTNodeKind::IndexOrCallExpr) {
            auto& ic = static_cast<const IndexOrCallExpr&>(*e);
            if (ic.callee && ic.callee->kind == ASTNodeKind::IdentifierExpr)
                nm = Symbol::toLower(static_cast<const IdentifierExpr&>(*ic.callee).name);
        }
        auto it = arrayDimCounts_.find(nm);
        return it == arrayDimCounts_.end() || it->second <= 1;
    };
    if (wholeArray1D(node.varName.get())) {
        std::string recnum = "0";
        if (node.recordNumber) { emitExpr(*node.recordNumber); recnum = std::move(lastExpr_); }
        c_.emitLine("vb6_PutArray(" + fnum + ", " + recnum + ", " + varExpr + ");");
        return;
    }
    // 确定变量大小: 根据变量名查询已知类型
    std::string varLower = varExpr;
    std::transform(varLower.begin(), varLower.end(), varLower.begin(), ::tolower);
    std::string bareName = varLower;
    std::string sizeExpr;
    if (knownBstrVars_.count(bareName)) {
        sizeExpr = "0";  // BSTR: varSize=0 让RTL层特殊处理
    } else if (knownDoubleVars_.count(bareName)) {
        sizeExpr = "sizeof(double)";
    } else if (knownLongVars_.count(bareName)) {
        sizeExpr = "sizeof(int32_t)";
    } else {
        sizeExpr = "sizeof(int32_t)";  // 默认: Long
    }
    if (node.recordNumber) {
        emitExpr(*node.recordNumber);
        std::string recnum = std::move(lastExpr_);
        c_.emitLine("vb6_Put(" + fnum + ", " + recnum + ", &" + varExpr + ", " + sizeExpr + ");");
    } else {
        c_.emitLine("vb6_Put(" + fnum + ", 0, &" + varExpr + ", " + sizeExpr + ");");
    }
}
void CCodeGen::visit(SeekStmt& node) {
    emitExpr(*node.fileNumber);
    std::string fnum = std::move(lastExpr_);
    emitExpr(*node.position);
    c_.emitLine("vb6_SeekStmt(" + fnum + ", " + lastExpr_ + ");");
}

void CCodeGen::visit(LockStmt& node) {
    emitExpr(*node.fileNumber);
    std::string fnum = std::move(lastExpr_);
    std::string startArg = "0", endArg = "0";
    if (node.start) {
        emitExpr(*node.start);
        startArg = std::move(lastExpr_);
        if (node.end) {
            emitExpr(*node.end);
            endArg = std::move(lastExpr_);
        } else {
            endArg = "0";
        }
    }
    c_.emitLine("vb6_Lock((int32_t)(" + fnum + "), (int64_t)(" + startArg + "), (int64_t)(" + endArg + "));");
}

void CCodeGen::visit(UnlockStmt& node) {
    emitExpr(*node.fileNumber);
    std::string fnum = std::move(lastExpr_);
    std::string startArg = "0", endArg = "0";
    if (node.start) {
        emitExpr(*node.start);
        startArg = std::move(lastExpr_);
        if (node.end) {
            emitExpr(*node.end);
            endArg = std::move(lastExpr_);
        } else {
            endArg = "0";
        }
    }
    c_.emitLine("vb6_Unlock((int32_t)(" + fnum + "), (int64_t)(" + startArg + "), (int64_t)(" + endArg + "));");
}

void CCodeGen::visit(ResetStmt& node) {
    (void)node;
    c_.emitLine("vb6_Reset();");
}

void CCodeGen::visit(WidthStmt& node) {
    emitExpr(*node.fileNumber);
    std::string fnum = std::move(lastExpr_);
    emitExpr(*node.width);
    std::string w = std::move(lastExpr_);
    c_.emitLine("vb6_Width((int32_t)(" + fnum + "), (int32_t)(" + w + "));");
}

void CCodeGen::visit(KillStmt& node) {
    emitExpr(*node.pathName);
    c_.emitLine("vb6_Kill(" + lastExpr_ + ");");
}

void CCodeGen::visit(NameStmt& node) {
    emitExpr(*node.oldPath);
    std::string oldP = std::move(lastExpr_);
    emitExpr(*node.newPath);
    c_.emitLine("vb6_Name(" + oldP + ", " + lastExpr_ + ");");
}

void CCodeGen::visit(MkDirStmt& node) {
    emitExpr(*node.pathName);
    c_.emitLine("vb6_MkDir(" + lastExpr_ + ");");
}

void CCodeGen::visit(RmDirStmt& node) {
    emitExpr(*node.pathName);
    c_.emitLine("vb6_RmDir(" + lastExpr_ + ");");
}

void CCodeGen::visit(ChDirStmt& node) {
    emitExpr(*node.pathName);
    c_.emitLine("vb6_ChDir(" + lastExpr_ + ");");
}

void CCodeGen::visit(ChDriveStmt& node) {
    emitExpr(*node.drive);
    c_.emitLine("vb6_ChDrive(" + lastExpr_ + ");");
}

void CCodeGen::visit(FileCopyStmt& node) {
    emitExpr(*node.source);
    std::string src = std::move(lastExpr_);
    emitExpr(*node.destination);
    c_.emitLine("vb6_FileCopy(" + src + ", " + lastExpr_ + ");");
}


} // namespace vb6c3
