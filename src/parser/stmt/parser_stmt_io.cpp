// vb6c3 - 语句解析器
// VB6 块语句 + 单行语句

#include "parser/parser.hpp"

namespace vb6c3 {

// --- parser_stmt_io.cpp: 文件与 I-O 语句（Open / Close / Get / Put / Input / Print / Write / Seek / Lock / Name / Kill / MkDir …） ---


// ============================================================
// 文件 I/O 语句
// ============================================================

std::unique_ptr<OpenStmt> Parser::parseOpenStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Open'
    auto pathName = parseExpression();

    OpenMode mode = OpenMode::Random;
    if (match(TokenKind::For)) {
        if (match(TokenKind::Input))          mode = OpenMode::Input;
        else if (match(TokenKind::Output))    mode = OpenMode::Output;
        else if (match(TokenKind::Append))     mode = OpenMode::Append;
        else if (match(TokenKind::Binary))     mode = OpenMode::Binary;
        else if (match(TokenKind::Random))     mode = OpenMode::Random;
    }

    OpenAccess access = OpenAccess::Default;
    if (match(TokenKind::Access)) {
        if (match(TokenKind::Read)) {
            if (match(TokenKind::Write)) access = OpenAccess::ReadWrite;
            else access = OpenAccess::Read;
        } else if (match(TokenKind::Write)) {
            access = OpenAccess::Write;
        } else if (match(TokenKind::ReadWrite)) {
            access = OpenAccess::ReadWrite;
        }
    }

    LockType lock = LockType::Default;
    if (match(TokenKind::Shared)) {
        lock = LockType::Shared;
    } else if (cur_.kind == TokenKind::Lock) {
        advance();
        if (match(TokenKind::Read)) {
            if (match(TokenKind::Write)) lock = LockType::LockReadWrite;
            else lock = LockType::LockRead;
        } else if (match(TokenKind::Write)) {
            lock = LockType::LockWrite;
        }
    }

    expect(TokenKind::As, DiagnosticID::ParseExpectedToken,
           "expected 'As' in Open statement");
    if (cur_.kind == TokenKind::Hash) advance();  // 可选的 #
    auto fileNumber = parseExpression();

    ExprPtr recordLength;
    // Len=reclength (可选)
    if (cur_.kind == TokenKind::Identifier && toLower(cur_.text) == "len") {
        advance(); // consume 'Len'
        if (match(TokenKind::Equals)) {
            recordLength = parseExpression();
        }
    }

    return std::make_unique<OpenStmt>(loc, std::move(pathName), mode, access,
        lock, std::move(fileNumber), std::move(recordLength));
}

std::unique_ptr<CloseStmt> Parser::parseCloseStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Close'
    std::vector<ExprPtr> fileNumbers;
    if (cur_.kind != TokenKind::NewLine && cur_.kind != TokenKind::EndOfFile) {
        do {
            if (cur_.kind == TokenKind::Hash) advance();
            fileNumbers.push_back(parseExpression());
        } while (match(TokenKind::Comma));
    }
    return std::make_unique<CloseStmt>(loc, std::move(fileNumbers));
}

std::unique_ptr<GetStmt> Parser::parseGetStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Get'
    if (cur_.kind == TokenKind::Hash) advance();
    auto fileNumber = parseExpression();
    ExprPtr recordNumber;
    if (match(TokenKind::Comma)) {
        // VB6 允许省略记录号: Get #1, , data
        if (cur_.kind != TokenKind::Comma) {
            recordNumber = parseExpression();
        }
    }
    expect(TokenKind::Comma, DiagnosticID::ParseExpectedToken,
           "expected ',' in Get statement");
    auto varName = parseExpression();
    return std::make_unique<GetStmt>(loc, std::move(fileNumber),
        std::move(recordNumber), std::move(varName));
}

std::unique_ptr<PutStmt> Parser::parsePutStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Put'
    if (cur_.kind == TokenKind::Hash) advance();
    auto fileNumber = parseExpression();
    ExprPtr recordNumber;
    if (match(TokenKind::Comma)) {
        // VB6 允许省略记录号: Put #1, , data
        if (cur_.kind != TokenKind::Comma) {
            recordNumber = parseExpression();
        }
    }
    expect(TokenKind::Comma, DiagnosticID::ParseExpectedToken,
           "expected ',' in Put statement");
    auto varName = parseExpression();
    return std::make_unique<PutStmt>(loc, std::move(fileNumber),
        std::move(recordNumber), std::move(varName));
}

std::unique_ptr<InputStmt> Parser::parseInputStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Input'
    if (cur_.kind == TokenKind::Hash) advance();
    auto fileNumber = parseExpression();
    expect(TokenKind::Comma, DiagnosticID::ParseExpectedToken,
           "expected ',' in Input statement");
    std::vector<ExprPtr> varList;
    varList.push_back(parseExpression());
    while (match(TokenKind::Comma)) {
        varList.push_back(parseExpression());
    }
    return std::make_unique<InputStmt>(loc, std::move(fileNumber), std::move(varList));
}

std::unique_ptr<PrintStmt> Parser::parsePrintStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Print'
    
    // M22-Issue6: Distinguish "Print expr" (form surface) from "Print #n, expr" (file I/O)
    // If '#' is present, this is file I/O mode
    // If no '#', the first expression is the first output item (form print mode)
    bool hasHash = (cur_.kind == TokenKind::Hash);
    if (hasHash) advance();
    
    auto firstExpr = parseExpression();
    std::vector<ExprPtr> outputList;
    std::vector<char> separators;
    ExprPtr fileNumber;
    bool isFormPrint = false;
    
    if (hasHash) {
        // File I/O mode: "Print #n, expr1; expr2; ..."
        fileNumber = std::move(firstExpr);
        if (match(TokenKind::Comma) || match(TokenKind::Semicolon)) {
            while (cur_.kind != TokenKind::NewLine && cur_.kind != TokenKind::EndOfFile) {
                outputList.push_back(parseExpression());
                // 记录这一项**后面**的分隔符: ';'/'，' 或 0=语句尾。尾随分隔符
                // (Print #f, 1;) 只影响行尾, 也必须留下, 否则会被当成 0 补 CRLF。
                if (cur_.kind == TokenKind::Comma) { separators.push_back(','); advance(); }
                else if (cur_.kind == TokenKind::Semicolon) { separators.push_back(';'); advance(); }
                else { separators.push_back(0); break; }
                if (cur_.kind == TokenKind::NewLine || cur_.kind == TokenKind::EndOfFile) break;
            }
        }
    } else {
        // Form print mode: "Print expr1; expr2; ..."
        // The first expression IS the first output item
        isFormPrint = true;
        outputList.push_back(std::move(firstExpr));
        while (match(TokenKind::Semicolon) || match(TokenKind::Comma)) {
            if (cur_.kind == TokenKind::NewLine || cur_.kind == TokenKind::EndOfFile) break;
            outputList.push_back(parseExpression());
        }
    }
    
    auto stmt = std::make_unique<PrintStmt>(loc, std::move(fileNumber), std::move(outputList));
    stmt->isFormPrint = isFormPrint;
    stmt->separators = std::move(separators);
    return stmt;
}

std::unique_ptr<WriteStmt> Parser::parseWriteStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Write'
    if (cur_.kind == TokenKind::Hash) advance();
    auto fileNumber = parseExpression();
    std::vector<ExprPtr> outputList;
    if (match(TokenKind::Comma) || match(TokenKind::Semicolon)) {
        while (cur_.kind != TokenKind::NewLine && cur_.kind != TokenKind::EndOfFile) {
            outputList.push_back(parseExpression());
            if (!match(TokenKind::Comma) && !match(TokenKind::Semicolon)) break;
        }
    }
    return std::make_unique<WriteStmt>(loc, std::move(fileNumber), std::move(outputList));
}

std::unique_ptr<LineInputStmt> Parser::parseLineInputStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Line'
    expect(TokenKind::Input, DiagnosticID::ParseExpectedToken,
           "expected 'Input' after 'Line'");
    if (cur_.kind == TokenKind::Hash) advance();
    auto fileNumber = parseExpression();
    expect(TokenKind::Comma, DiagnosticID::ParseExpectedToken,
           "expected ',' in Line Input statement");
    auto varName = parseExpression();
    return std::make_unique<LineInputStmt>(loc, std::move(fileNumber), std::move(varName));
}

std::unique_ptr<WidthStmt> Parser::parseWidthStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Width'
    if (cur_.kind == TokenKind::Hash) advance();
    auto fileNumber = parseExpression();
    expect(TokenKind::Comma, DiagnosticID::ParseExpectedToken,
           "expected ',' in Width statement");
    auto width = parseExpression();
    return std::make_unique<WidthStmt>(loc, std::move(fileNumber), std::move(width));
}

std::unique_ptr<SeekStmt> Parser::parseSeekStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Seek'
    if (cur_.kind == TokenKind::Hash) advance();
    auto fileNumber = parseExpression();
    expect(TokenKind::Comma, DiagnosticID::ParseExpectedToken,
           "expected ',' in Seek statement");
    auto position = parseExpression();
    return std::make_unique<SeekStmt>(loc, std::move(fileNumber), std::move(position));
}

std::unique_ptr<LockStmt> Parser::parseLockStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Lock'
    if (cur_.kind == TokenKind::Hash) advance();
    auto fileNumber = parseExpression();
    ExprPtr start, end;
    if (match(TokenKind::Comma)) {
        start = parseExpression();
        if (match(TokenKind::To)) {
            end = parseExpression();
        }
    }
    return std::make_unique<LockStmt>(loc, std::move(fileNumber), std::move(start), std::move(end));
}

std::unique_ptr<UnlockStmt> Parser::parseUnlockStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Unlock'
    if (cur_.kind == TokenKind::Hash) advance();
    auto fileNumber = parseExpression();
    ExprPtr start, end;
    if (match(TokenKind::Comma)) {
        start = parseExpression();
        if (match(TokenKind::To)) {
            end = parseExpression();
        }
    }
    return std::make_unique<UnlockStmt>(loc, std::move(fileNumber), std::move(start), std::move(end));
}

std::unique_ptr<NameStmt> Parser::parseNameStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Name'
    auto oldPath = parseExpression();
    expect(TokenKind::As, DiagnosticID::ParseExpectedToken,
           "expected 'As' in Name statement");
    auto newPath = parseExpression();
    return std::make_unique<NameStmt>(loc, std::move(oldPath), std::move(newPath));
}

std::unique_ptr<FileCopyStmt> Parser::parseFileCopyStmt() {
    auto loc = currentLoc();
    advance(); // consume 'FileCopy'
    auto source = parseExpression();
    expect(TokenKind::Comma, DiagnosticID::ParseExpectedToken,
           "expected ',' in FileCopy statement");
    auto dest = parseExpression();
    return std::make_unique<FileCopyStmt>(loc, std::move(source), std::move(dest));
}

std::unique_ptr<KillStmt> Parser::parseKillStmt() {
    auto loc = currentLoc();
    advance(); // consume 'Kill'
    auto path = parseExpression();
    return std::make_unique<KillStmt>(loc, std::move(path));
}

std::unique_ptr<MkDirStmt> Parser::parseMkDirStmt() {
    auto loc = currentLoc();
    advance(); // consume 'MkDir'
    auto path = parseExpression();
    return std::make_unique<MkDirStmt>(loc, std::move(path));
}

std::unique_ptr<RmDirStmt> Parser::parseRmDirStmt() {
    auto loc = currentLoc();
    advance(); // consume 'RmDir'
    auto path = parseExpression();
    return std::make_unique<RmDirStmt>(loc, std::move(path));
}

std::unique_ptr<ChDirStmt> Parser::parseChDirStmt() {
    auto loc = currentLoc();
    advance(); // consume 'ChDir'
    auto path = parseExpression();
    return std::make_unique<ChDirStmt>(loc, std::move(path));
}

std::unique_ptr<ChDriveStmt> Parser::parseChDriveStmt() {
    auto loc = currentLoc();
    advance(); // consume 'ChDrive'
    auto drive = parseExpression();
    return std::make_unique<ChDriveStmt>(loc, std::move(drive));
}

// ============================================================
// Beep / DoEvents
// ============================================================

StmtPtr Parser::parseBeepOrDoEvents() {
    auto loc = currentLoc();
    if (cur_.kind == TokenKind::Beep) {
        advance();
        return std::make_unique<BeepStmt>(loc);
    }
    advance(); // consume 'DoEvents'
    return std::make_unique<DoEventsStmt>(loc);
}

// ============================================================
// Attribute (过程体内)
// ============================================================

std::unique_ptr<AttributeStmt> Parser::parseAttributeInBody() {
    return parseAttribute();
}

} // namespace vb6c3
