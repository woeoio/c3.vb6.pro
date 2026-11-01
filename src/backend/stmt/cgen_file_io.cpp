#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_file_io.cpp: 文件 I/O 语句生成 (Open/Close/Print/Write/Input/Get/Put...) ---

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

void CCodeGen::visit(PrintStmt& node) {
    // M22-Issue6: BSTR expression detection shared by both Print modes
    static const std::vector<std::string> bstrFuncs = {
        "vb6_BSTR_FromStr", "vb6_Left", "vb6_Right", "vb6_Mid",
        "vb6_UCase", "vb6_LCase", "vb6_UCase_str", "vb6_LCase_str",
        "vb6_Trim", "vb6_LTrim", "vb6_RTrim", "vb6_Chr",
        "vb6_Str", "vb6_CStr", "vb6_Format", "vb6_Hex", "vb6_Oct",
        "vb6_Replace", "vb6_Space", "vb6_String", "vb6_StrReverse",
        "vb6_BSTR_Concat", "vb6_BSTR_Empty", "vb6_App_Path", "vb6_App_EXEName", "vb6_App_Title", "vb6_App_HelpFile", "vb6_Command", "vb6_CurDir", "vb6_Environ", "vb6_Dir", "vb6_IIfBSTR", "vb6_GetControlText", "vb6_GetControlCaption"
    };
    auto isBstrExpr = [&](const std::string& expr) -> bool {
        for (auto& prefix : bstrFuncs) {
            if (expr.compare(0, prefix.size(), prefix) == 0) return true;
        }
        if (expr.compare(0, 8, "vb6_BSTR") == 0) return true;
        if (expr.find("VB6_SA_AT(BSTR,") != std::string::npos) return true;
        std::string lower = expr;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        if (knownBstrVars_.count(lower)) return true;
        return false;
    };

    if (node.isFormPrint) {
        // M22-Issue6: Form surface Print - "Print expr" in form module
        // Generate: vb6_Form_Print(formHwnd, bstrExpr) for each output item
        std::string formHwnd = isFormModule_ ? ("vb6_hwnd_" + cIdent(formName_)) : "NULL";
        for (auto& expr : node.outputList) {
            emitExpr(*expr);
            std::string val = lastExpr_;
            if (isBstrExpr(val)) {
                c_.emitLine("vb6_Form_Print(" + formHwnd + ", " + val + ");");
            } else {
                // Use wrapToBSTR for proper type conversion (Date→CStrDate, etc.)
                std::string bstrVal = wrapToBSTR(val, *expr);
                c_.emitLine("vb6_Form_Print(" + formHwnd + ", " + bstrVal + ");");
            }
        }
        return;
    }

    // File I/O Print: Print #fnum, expr1; expr2; ...
    emitExpr(*node.fileNumber);
    std::string fnum = std::move(lastExpr_);

    if (node.outputList.empty()) {
        c_.emitLine("vb6_Print(" + fnum + ", NULL);");
    } else {
        for (auto& expr : node.outputList) {
            emitExpr(*expr);
            std::string val = lastExpr_;
            if (isBstrExpr(val)) {
                // 已经是BSTR, 直接传给vb6_Print
                c_.emitLine("vb6_Print(" + fnum + ", " + val + ");");
            } else {
                // 非BSTR: 转换为BSTR后输出
                c_.emitLine("vb6_Print(" + fnum + ", vb6_Str((int32_t)(" + val + ")));");
            }
        }
    }
}

void CCodeGen::visit(WriteStmt& node) {
    emitExpr(*node.fileNumber);
    std::string fnum = std::move(lastExpr_);

    static const std::vector<std::string> bstrFuncs = {
        "vb6_BSTR_FromStr", "vb6_Left", "vb6_Right", "vb6_Mid",
        "vb6_UCase", "vb6_LCase", "vb6_UCase_str", "vb6_LCase_str",
        "vb6_Trim", "vb6_LTrim", "vb6_RTrim", "vb6_Chr",
        "vb6_Str", "vb6_CStr", "vb6_Format", "vb6_Hex", "vb6_Oct",
        "vb6_Replace", "vb6_Space", "vb6_String", "vb6_StrReverse",
        "vb6_BSTR_Concat", "vb6_BSTR_Empty", "vb6_App_Path", "vb6_App_EXEName", "vb6_App_Title", "vb6_App_HelpFile", "vb6_Command", "vb6_CurDir", "vb6_Environ", "vb6_Dir", "vb6_IIfBSTR", "vb6_GetControlText", "vb6_GetControlCaption"
    };
    auto isBstrExpr = [&](const std::string& expr) -> bool {
        for (auto& prefix : bstrFuncs) {
            if (expr.compare(0, prefix.size(), prefix) == 0) return true;
        }
        if (expr.compare(0, 8, "vb6_BSTR") == 0) return true;
        if (expr.find("VB6_SA_AT(BSTR,") != std::string::npos) return true;
        std::string lower = expr;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        if (knownBstrVars_.count(lower)) return true;
        return false;
    };

    if (node.outputList.empty()) {
        c_.emitLine("vb6_Write(" + fnum + ", NULL);");
    } else {
        for (auto& expr : node.outputList) {
            emitExpr(*expr);
            std::string val = lastExpr_;
            if (isBstrExpr(val)) {
                c_.emitLine("vb6_Write(" + fnum + ", " + val + ");");
            } else {
                c_.emitLine("vb6_Write(" + fnum + ", vb6_Str((int32_t)(" + val + ")));");
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
