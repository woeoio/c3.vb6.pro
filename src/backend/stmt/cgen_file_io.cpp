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

// 账 #608: 「这一枚实参是不是已经是一枚 BSTR」的那道闸门，Print # 与 Write # 各抄过一份
// （名单 + 查表，两份一字一样），而且两份都只问**局部**那张表 ——
// `knownBstrVars_` 在每个过程入口被 clear 后只由局部声明与形参重新填
// (cgen_decl_proc.cpp:51 / cgen_decl_func.cpp:47)，所以模块级那句 `Public gAll As String`
// 压根不在里面 ⇒ `Print #f, gAll` 落进下面的数值兜底，发出来是
// `vb6_Print(fn, vb6_Str((int32_t)(gAll)))`：把 BSTR 指针当数打印，产物文件里只有一枚指针数
// (实测 probe_out.txt = "710757560")。真语料 D:\vb_yqt4qPac 里这一族有 48 处
// （Print#/Write# 共 2158 行，实参是裸标识符的 245 行，其中同名在**本文件列 0** 声明成
// `As String` 的 48 行）。
//
// 修法沿用账 #113 那把尺：只对**裸标识符**补问一次符号表（inferExprType 的 IdentifierExpr
// 分支认得模块级变量），成员/控件属性/数组元素那一族的既有通道一个都不碰 —— 它们各自另有主人。
// 数值兜底那条口径（`(int32_t)` 折成串）保持不变。
bool CCodeGen::fileOutArgIsBstr(const std::string& expr, Expr& argNode) const {
    static const std::vector<std::string> bstrFuncs = {
        "vb6_BSTR_FromStr", "vb6_Left", "vb6_Right", "vb6_Mid",
        "vb6_UCase", "vb6_LCase", "vb6_UCase_str", "vb6_LCase_str",
        "vb6_Trim", "vb6_LTrim", "vb6_RTrim", "vb6_Chr",
        "vb6_Str", "vb6_CStr", "vb6_Format", "vb6_Hex", "vb6_Oct",
        "vb6_Replace", "vb6_Space", "vb6_String", "vb6_StrReverse",
        "vb6_BSTR_Concat", "vb6_BSTR_Empty", "vb6_App_Path", "vb6_App_EXEName", "vb6_App_Title", "vb6_App_HelpFile", "vb6_Command", "vb6_CurDir", "vb6_Environ", "vb6_Dir", "vb6_IIfBSTR", "vb6_GetControlText", "vb6_GetControlCaption"
    };
    // 账 #305 第331刀: 名单匹配从「裸前缀」收成「顶层函数名精确」—— 与 Fix 100a/100b/100d
    // 同一族修法，裸前缀会把**工程自己的函数**蒙进来：`Direction() As Double` 发成
    // `vb6_Direction()` 命中名单里的 `vb6_Dir`、`Stripped() As Long` 命中 `vb6_Str`，
    // 明明是数值却按 BSTR 直通 ⇒ 改前实测 `Print #f, Direction()` 直接
    // C2440「无法从 double 转换为 BSTR」(整棵工程编不过，不是读数错)。
    // 反向也一样：返回 String 的工程函数只有**撞名** (`StrFn` 命中 `vb6_Str`) 才蒙对，
    // 没撞名 (`MakeStr`) 就落进那条 int32 兜底 —— §B11 那一格真正的成因。
    const size_t pOpen = expr.find('(');
    const std::string fnName = (pOpen == std::string::npos) ? std::string() : expr.substr(0, pOpen);
    if (!fnName.empty()) {
        for (auto& name : bstrFuncs) {
            if (fnName == name) return true;
        }
    }
    if (expr.compare(0, 8, "vb6_BSTR") == 0) return true;
    if (expr.find("VB6_SA_AT(BSTR,") != std::string::npos) return true;
    std::string lower = expr;
    std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
    if (knownBstrVars_.count(lower)) return true;
    if (argNode.kind == ASTNodeKind::IdentifierExpr &&
        inferExprType(argNode) == Vb6Type::String) return true;
    return false;
}

// 账 #305 第331刀 (台账 §B11): 「这一发实参该以什么形状交给写文件的出口」—— 三户同一条。
// 上一刀把「是不是已经是 BSTR」收成一位主人，但**非 BSTR 那半**两份抄本各走各的：
// Form Print 走 `wrapToBSTR`（按类型分档：Variant→解包、Date→CStrDate、Double→CStrDbl、
// 整数→CStrLong、返回 String 的工程函数→原样直通），而 `Print #` / `Write #` 自留一条
// `vb6_Str((int32_t)(val))`。那一条同时塌了三档，实测（`.local/k611/`）：
//   ① `Print #f, MakeStr()`（`Public Function MakeStr() As String`）⇒ BSTR 指针被截成 int32
//      打成数（§B11 那格；同一句写成不带括号 `Print #f, MakeStr` 反倒是对的 —— 因为它走
//      IdentifierExpr 那一问，两条写法本该答同一个数）；
//   ② `Print #f, MakeVar()`（`As Variant`）⇒ **C2440 编不过**：`无法从 vb6_VARIANT 转换到 int32_t`
//      （c3-error.log 里就是这一句，整棵工程编不出来，不是读数错）；
//   ③ `As Date` / `As Double` / `As Boolean` 的返回档被 `(int32_t)` 截掉 —— 日期打成序列数、
//      小数打成就整数。
// 修法 = 非 BSTR 那半也走 Form Print 那位主人（`wrapToBSTR`），三户一条口。
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
            // 账 #305 第331刀: 两问收进一位主人，本户的形状一字未动。
            c_.emitLine("vb6_Form_Print(" + formHwnd + ", "
                        + fileOutArgBstrExpr(val, *expr) + ");");
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
            // 账 #305 第331刀 (台账 §B11): 原来非 BSTR 那支自留 `vb6_Str((int32_t)(val))`，
            // 把「返回 String 的工程函数」的指针当数打、把「返回 Variant 的」直接编不过 (C2440)、
            // 把 Date/Double 截成整数。现在与 Form Print 同一条口 (wrapToBSTR 按档分诊)。
            c_.emitLine("vb6_Print(" + fnum + ", " + fileOutArgBstrExpr(val, *expr) + ");");
        }
    }
}

void CCodeGen::visit(WriteStmt& node) {
    emitExpr(*node.fileNumber);
    std::string fnum = std::move(lastExpr_);

    if (node.outputList.empty()) {
        c_.emitLine("vb6_Write(" + fnum + ", NULL);");
    } else {
        for (auto& expr : node.outputList) {
            emitExpr(*expr);
            std::string val = lastExpr_;
            // 账 #305 第331刀: 与 Print # 同一句病灶（Write # 那支抄的是同一条 int32 兜底）。
            c_.emitLine("vb6_Write(" + fnum + ", " + fileOutArgBstrExpr(val, *expr) + ");");
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
