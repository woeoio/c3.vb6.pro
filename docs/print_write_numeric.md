# Print # / Write # 数值输出语义（VB6 实测口径）

本文件是 `Print #` / `Write #` 文本格式的**唯一口径说明**，也是 builtin 返回类型表与
`inferExprType` 边界清单的落点。代码里的引用点：

- `src/rtl/core/vb6rtl/vb6rtl_file.c`（`vb6_File*Text` / `vb6_PrintText` / `vb6_WriteText`）
- `src/backend/stmt/cgen_file_io.cpp`（`printWriteTextExpr`）
- `src/backend/detail/expr/cgen_builtin_return_type_table.inc`

## 1. 口径来源与测量方法

- 真 VB6：`C:\Program Files (x86)\VB6Mini\bin\VB6.EXE`（`/make` 路线），参照工程
  `.temp/m4/vb6num`（Print 分支）与同一工程的 Write 分支，逐项写盘。
- 对照：同一份 VB 源码用本机 `.build\C3.exe` 重编、运行、写盘。
- 判定：两个文件按**字节**比较；修复前 = 旧 `(int32_t)` 路径产物，修复后 = 本批产物，
  VB6 = 参照工程产物。采样 74 点（Print 46 + Write 28）。
- 证据：`.temp/m4/three_col.csv`（原始 6 列）、`.temp/m4/three_col.md`（本文件 §4 的来源）。
- 结果：`修复后 == VB6` **74/74**；`修复前 != VB6` **73/74**（唯一相同的是
  `p38_empty_str`——空串修前修后都是 `\r\n`）。74 点里有 7 点在旧路径上直接
  编译失败 `C2440/C2198`，不是"写错值"——见 §4 的 `<旧路径编译失败>` 标记。

**与官方文档的冲突点：0 条。** 公开文档只写了"数值用定点或指数表示"，没有规定
前导空格、尾部空格、有效位数、定点/指教的切换边界、`Date` 的 `[0,1)` 分支、`Decimal`
的输出位数——这些是文档留白，本文件按实测补齐。凡留白处一律以实测字节为准。

## 2. VB6 实测结论：`Print #`

| 项 | 实测口径 | 证据 |
|---|---|---|
| 正数前导空格 | **成立**：非负（含 0）输出 1 个前导空格 | ` 123 `、` 0 `、` 3.25 ` |
| 负数 | `-` 自己占那一格，**不再补**前导空格 | `-123 `、`-0.5 ` |
| 尾部空格 | 每个数值/日期/布尔项后恒 **1 个**空格 | 所有行 |
| 整数（Integer/Long/Byte…） | 十进制，无小数点 | ` 40000 `、`-32768 ` |
| Single | **7 位有效数字** | ` 0.3333333 ` |
| Double | **15 位有效数字** | ` 3.1415926535 `、` 0.333333333333333 ` |
| Currency | 定点、4 位小数 | ` 1.2345 `、` 1234.5678 ` |
| Decimal | 定点、按实际有效位（存储 28 位） | ` 1.234567890123456789 ` |
| 定点/指数切换 | 按**有效位数**决定，不按固定阈值；定点区间实测 `[1e-15, 1e15)`，之外用指数 | ` 1E+20 `、` 1E-20 `、` 999999999999999 `、` 0.000001 ` |
| 指数字母/符号 | 大写 `E`；指数带 `+`/`-`，不补零 | `1E+20`、`1E-20` |
| Date（含日期+时间） | 本地化 `CStr`（本机 zh-CN）：`yyyy/M/d H:MM:SS` + 尾空格 | `2020/1/2 15:04:05 ` |
| Date（纯日期） | `yyyy/M/d` | `2020/1/2 ` |
| Date（`[0,1)` 只有时间） | 只输出时间，**小时不补零** | `15:04:05 ` |
| Boolean | `True` / `False`（无前导空格、无尾空格） | `True`、`False` |
| Empty | 空（该位置不产生任何字节，只剩语句末 CRLF） | `\r\n` |
| Null | `Null` | `Null` |
| 字符串项 | 原样输出，不加引号 | `ABC` |
| 分隔 `;` | 直接续写（不补空格） | ` 1  2 `（两个数值各带前导空格） |
| 分隔 `,` | 补到下一个 14 列区（第 15 列起） | ` 1             2 ` |
| 语句末 | 一个 CRLF | 所有行 |

## 3. VB6 实测结论：`Write #`（独立语义，不与 Print 共用格式化）

| 项 | 实测口径 | 证据 |
|---|---|---|
| 数值前导/尾部空格 | **都没有**（去掉了 `Str` 的语义） | `123`、`3.25` |
| 小数 `|x|<1` | 去前导 0：`.25` / `-.25` | `.25`、`-.25`、`-.00001` |
| Single/Double 有效位 | 与 Print 相同（7 / 15） | `.3333333`、`.333333333333333` |
| 指数 | 大写 `E`，无空格 | `1E+20`、`1E-20` |
| Date | **恒** `#YYYY-MM-DD HH:MM:SS#`（不是本地化格式） | `#2020-01-02 15:04:05#` |
| Date（只有时间） | 落在 `1899-12-30`：`#1899-12-30 15:04:05#` | `#1899-12-30 15:04:05#` |
| Boolean | `#TRUE#` / `#FALSE#` | `#TRUE#`、`#FALSE#` |
| Null | `#NULL#` | `#NULL#` |
| Empty | 空 | `\r\n` |
| 字符串 | 双引号包裹，内部 `"` **双写** | `"ABC"`、`"A""B"` |
| 分隔 | 项间恒 `,`（源码写 `;` 也输出 `,`） | `1,"x"` |
| 语句末 | 一个 CRLF | 所有行 |

修复前的 Write # 走的是 Print 字符串格式器：`Write #f, 3.25` 写盘为 `"3",`——带引号、
逗号，且没有 CRLF。

## 4. 三列逐行对比（修复前 / 修复后 / VB6 参照）

`\r\n` 为字面转义；`<旧路径编译失败 C2440/C2198>` 表示旧路径在该形态上根本编不出
可执行文件（Variant 结构体被强转成标量），不是"写错值"。

| # | 夹具 | 修复前 (Before) | 修复后 (After) | VB6 实测 (Ref) |
|---|---|---|---|---|
| 1 | p01_lit_int.txt | `123\r\n` | ` 123 \r\n` | ` 123 \r\n` |
| 2 | p02_lit_neg.txt | `-123\r\n` | `-123 \r\n` | `-123 \r\n` |
| 3 | p03_zero.txt | `0\r\n` | ` 0 \r\n` | ` 0 \r\n` |
| 4 | p04_int_var.txt | `42\r\n` | ` 42 \r\n` | ` 42 \r\n` |
| 5 | p05_long_var.txt | `123456\r\n` | ` 123456 \r\n` | ` 123456 \r\n` |
| 6 | p06_single.txt | `3\r\n` | ` 3.25 \r\n` | ` 3.25 \r\n` |
| 7 | p07_single_third.txt | `0\r\n` | ` 0.3333333 \r\n` | ` 0.3333333 \r\n` |
| 8 | p08_double_325.txt | `3\r\n` | ` 3.25 \r\n` | ` 3.25 \r\n` |
| 9 | p09_double_pi.txt | `3\r\n` | ` 3.1415926535 \r\n` | ` 3.1415926535 \r\n` |
| 10 | p10_1e20.txt | `-2147483648\r\n` | ` 1E+20 \r\n` | ` 1E+20 \r\n` |
| 11 | p11_1e-20.txt | `0\r\n` | ` 1E-20 \r\n` | ` 1E-20 \r\n` |
| 12 | p12_double_third.txt | `0\r\n` | ` 0.333333333333333 \r\n` | ` 0.333333333333333 \r\n` |
| 13 | p13_double_negthird.txt | `0\r\n` | `-0.333333333333333 \r\n` | `-0.333333333333333 \r\n` |
| 14 | p14_currency_325.txt | `3\r\n` | ` 3.25 \r\n` | ` 3.25 \r\n` |
| 15 | p15_currency_12345.txt | `1\r\n` | ` 1.2345 \r\n` | ` 1.2345 \r\n` |
| 16 | p16_currency_big.txt | `1234\r\n` | ` 1234.5678 \r\n` | ` 1234.5678 \r\n` |
| 17 | p17_decimal_325.txt | `<旧路径编译失败 C2440/C2198>` | ` 3.25 \r\n` | ` 3.25 \r\n` |
| 18 | p18_decimal_pi.txt | `<旧路径编译失败 C2440/C2198>` | ` 3.1415926535 \r\n` | ` 3.1415926535 \r\n` |
| 19 | p19_date_time.txt | `43832\r\n` | `2020/1/2 15:04:05 \r\n` | `2020/1/2 15:04:05 \r\n` |
| 20 | p20_date_only.txt | `43832\r\n` | `2020/1/2 \r\n` | `2020/1/2 \r\n` |
| 21 | p21_time_only.txt | `0\r\n` | `15:04:05 \r\n` | `15:04:05 \r\n` |
| 22 | p22_true.txt | `-1\r\n` | `True\r\n` | `True\r\n` |
| 23 | p23_false.txt | `0\r\n` | `False\r\n` | `False\r\n` |
| 24 | p25_null.txt | `<旧路径编译失败 C2440/C2198>` | `Null\r\n` | `Null\r\n` |
| 25 | p26_long_lit.txt | `40000\r\n` | ` 40000 \r\n` | ` 40000 \r\n` |
| 26 | p27_int_max.txt | `32767\r\n` | ` 32767 \r\n` | ` 32767 \r\n` |
| 27 | p28_int_min.txt | `-32768\r\n` | `-32768 \r\n` | `-32768 \r\n` |
| 28 | p29_single_lit.txt | `3\r\n` | ` 3.25 \r\n` | ` 3.25 \r\n` |
| 29 | p30_double_lit.txt | `3\r\n` | ` 3.25 \r\n` | ` 3.25 \r\n` |
| 30 | p31_half.txt | `0\r\n` | ` 0.5 \r\n` | ` 0.5 \r\n` |
| 31 | p32_neg_half.txt | `0\r\n` | `-0.5 \r\n` | `-0.5 \r\n` |
| 32 | p33_100.txt | `100\r\n` | ` 100 \r\n` | ` 100 \r\n` |
| 33 | p34_multi_semi.txt | `1\r\n2\r\n` | ` 1  2 \r\n` | ` 1  2 \r\n` |
| 34 | p35_multi_comma.txt | `1\r\n2\r\n` | ` 1             2 \r\n` | ` 1             2 \r\n` |
| 35 | p36_mixed.txt | `A\r\n1\r\nB\r\n` | `A 1 B\r\n` | `A 1 B\r\n` |
| 36 | p24_empty.txt | `<旧路径编译失败 C2440/C2198>` | `\r\n` | `\r\n` |
| 37 | p38_empty_str.txt | `\r\n` | `\r\n` | `\r\n` |
| 38 | p42_0_25.txt | `0\r\n` | ` 0.25 \r\n` | ` 0.25 \r\n` |
| 39 | p43_0_0001.txt | `0\r\n` | ` 0.0001 \r\n` | ` 0.0001 \r\n` |
| 40 | p45_1e-5.txt | `0\r\n` | ` 0.00001 \r\n` | ` 0.00001 \r\n` |
| 41 | p47_123456789012345.txt | `-2147483648\r\n` | ` 123456789012345 \r\n` | ` 123456789012345 \r\n` |
| 42 | p49_dec_long.txt | `<旧路径编译失败 C2440/C2198>` | ` 1.234567890123456789 \r\n` | ` 1.234567890123456789 \r\n` |
| 43 | p50_1e-6.txt | `0\r\n` | ` 0.000001 \r\n` | ` 0.000001 \r\n` |
| 44 | p55_999999999999999.txt | `-2147483648\r\n` | ` 999999999999999 \r\n` | ` 999999999999999 \r\n` |
| 45 | p60_dec_quarter.txt | `<旧路径编译失败 C2440/C2198>` | ` 0.25 \r\n` | ` 0.25 \r\n` |
| 46 | p61_dec_neg.txt | `<旧路径编译失败 C2440/C2198>` | `-0.25 \r\n` | `-0.25 \r\n` |
| 47 | w01_lit_int.txt | `"123",` | `123\r\n` | `123\r\n` |
| 48 | w02_lit_neg.txt | `"-123",` | `-123\r\n` | `-123\r\n` |
| 49 | w03_zero.txt | `"0",` | `0\r\n` | `0\r\n` |
| 50 | w04_double_325.txt | `"3",` | `3.25\r\n` | `3.25\r\n` |
| 51 | w05_double_pi.txt | `"3",` | `3.1415926535\r\n` | `3.1415926535\r\n` |
| 52 | w06_1e20.txt | `"-2147483648",` | `1E+20\r\n` | `1E+20\r\n` |
| 53 | w07_1e-20.txt | `"0",` | `1E-20\r\n` | `1E-20\r\n` |
| 54 | w08_single.txt | `"3",` | `3.25\r\n` | `3.25\r\n` |
| 55 | w09_currency.txt | `"3",` | `3.25\r\n` | `3.25\r\n` |
| 56 | w10_decimal.txt | `<旧路径编译失败 C2440/C2198>` | `3.25\r\n` | `3.25\r\n` |
| 57 | w11_date_time.txt | `"43832",` | `#2020-01-02 15:04:05#\r\n` | `#2020-01-02 15:04:05#\r\n` |
| 58 | w12_time_only.txt | `"0",` | `#1899-12-30 15:04:05#\r\n` | `#1899-12-30 15:04:05#\r\n` |
| 59 | w13_true.txt | `"-1",` | `#TRUE#\r\n` | `#TRUE#\r\n` |
| 60 | w14_false.txt | `"0",` | `#FALSE#\r\n` | `#FALSE#\r\n` |
| 61 | w16_null.txt | `<旧路径编译失败 C2440/C2198>` | `#NULL#\r\n` | `#NULL#\r\n` |
| 62 | w17_str.txt | `"ABC",` | `"ABC"\r\n` | `"ABC"\r\n` |
| 63 | w18_quote.txt | `"A"B",` | `"A""B"\r\n` | `"A""B"\r\n` |
| 64 | w19_multi_comma.txt | `"1","x",` | `1,"x"\r\n` | `1,"x"\r\n` |
| 65 | w20_multi_semi.txt | `"1","x",` | `1,"x"\r\n` | `1,"x"\r\n` |
| 66 | w21_half.txt | `"0",` | `.5\r\n` | `.5\r\n` |
| 67 | w15_empty.txt | `<旧路径编译失败 C2440/C2198>` | `\r\n` | `\r\n` |
| 68 | w23_empty_str.txt | `"",` | `""\r\n` | `""\r\n` |
| 69 | w26_0_25.txt | `"0",` | `.25\r\n` | `.25\r\n` |
| 70 | w32_third.txt | `"0",` | `.333333333333333\r\n` | `.333333333333333\r\n` |
| 71 | w33_single_third.txt | `"0",` | `.3333333\r\n` | `.3333333\r\n` |
| 72 | w42_neg_small.txt | `"0",` | `-.00001\r\n` | `-.00001\r\n` |
| 73 | w44_dec_quarter.txt | `<旧路径编译失败 C2440/C2198>` | `.25\r\n` | `.25\r\n` |
| 74 | w45_dec_neg.txt | `<旧路径编译失败 C2440/C2198>` | `-.25\r\n` | `-.25\r\n` |

## 5. 修复方式与影响面

### 5.1 RTL（`src/rtl/core/vb6rtl/vb6rtl_file.c`）

- `vb6_file_fmt_number`：按有效位数在定点/指数之间选（Single 7 / Double 15）。
- `vb6_FileNumText(v, isSingle, mode)` / `vb6_FileLongText` / `vb6_FileBoolText` /
  `vb6_FileDateText` / `vb6_FileVariantText`：**只产出文本**，分隔与 CRLF 不在这里决定。
  `mode=0` = Print #（`Str` 语义：非负前导空格 + 尾部空格），`mode=1` = Write #（无
  前导/尾空格，`|x|<1` 去前导 0）。
- `vb6_PrintText(f, s, sep)` / `vb6_WriteText(f, s, isString, sep)`：两套分隔器分别实现
  `;` / `,` / 语句末 CRLF；Write # 的引号与双写只在这里做。**Print 与 Write 不共用格式化
  函数**，避免一套代码套两个格式。

### 5.2 代码生成（`src/backend/stmt/cgen_file_io.cpp`）

`printWriteTextExpr` 按 `inferExprType` 的答案发对应的 `vb6_File*Text`：

| 静态类型 | 发射 |
|---|---|
| Byte/Integer/Long/SByte/UInteger/ULong/LongLong/ULongLong/LongPtr | `vb6_FileLongText((int64_t)(…), mode)` |
| Single | `vb6_FileNumText((double)(…), 1, mode)` |
| Double / Currency | `vb6_FileNumText((double)(…), 0, mode)` |
| Date | `vb6_FileDateText((double)(…), mode)` |
| Boolean | `vb6_FileBoolText((int32_t)(…), mode)` |
| Decimal | `vb6_FileVariantText(…, mode)`（C 侧 Decimal 就是整只 `vb6_VARIANT`） |
| Variant/Empty/Null **且有静态"整只槽"证据** | `vb6_FileVariantText(…, mode)` |
| 其余（类型答不准） | 旧路径 `vb6_Str((int32_t)(…))` |

"静态整只槽证据"= `cExprIsVariant`（RTL 返回 `vb6_VARIANT` 的出口名前缀、类字段槽、
数组元素槽）或 `isDefinitelyVariantExpr`。只看到类型名 `Variant` **不等于**有证据：
那可能只是 `inferExprType` 的兜底答案（见 §7）。

影响面：RTL +239 行、`cgen_file_io.cpp` +116/-13 行、`cgen_util_type.cpp` +10 行；另含
解析/克隆侧记录 `;` 与 `,` 的少量改动（`parser_stmt_io.cpp`、`ast_stmt_io.hpp`、
`ast_clone.cpp`）。既有 `String/BSTR` 站点与 `Variant 装字符串`、`UDT String 字段`两条
旧路径未动（A/B 证据见 §6.2 的 `0` 行）。

## 6. builtin 返回类型表与字符串兜底名单（第一步 + 第二步）

### 6.1 补齐的表

新增 `src/backend/detail/expr/cgen_builtin_return_type_table.inc`：给**语义层没有符号**、
但 cgen 会发射的内置函数一个返回档兜底。键按小写、剥掉尾缀 `$` 查表，所以
`Input`/`Input$`、`Command`/`Command$` 都命中同一行。共 16 个名字 / 15 个键：

| 名字 | 返回档 | 依据 |
|---|---|---|
| `CByte`（别名 `Cbyt`） | Byte | 语义层只登记 `CByte`，cgen 发射表里 `cbyt` 与 `cbyte` 双键 |
| `Kill` | Long | RTL `vb6_Kill` 返回 `int32_t` |
| `Randomize` | Void | RTL `vb6_Randomize` 返回 `void` |
| `Input` / `Input$` | String | `vb6_InputString` 返回 `BSTR` |
| `InputB` / `LeftB` / `RightB` / `MidB` | String | RTL 声明 `BSTR` |
| `LoadResString` | Variant | RTL 签名 `vb6_VARIANT vb6_LoadResString(...)` |
| `InStrB` / `LenB` | Long | RTL 交 `int32_t`（`vb6_LenB` 每一档都是 `int32_t`） |
| `TypeOf` | Boolean | `int32_t vb6_TypeOf(...)`，VB 语义是 Boolean |
| `LoadResData` / `LoadResPicture` | Variant | RTL 签名 `vb6_VARIANT` |
| `Command` / `Command$` | String | 语义层已登记 `command`；这行只为名字形状完整 |

**不登记多态内置**（`IIf` / `Choose` / `Switch` / `If`）：它们的 C 返回档由发码现场决定
（`vb6_IIfBSTR` / `vb6_IIfLong` / 裸三目…），固定成任何一档都会答错。

### 6.2 A/B 等价性（换判据前后逐字节对比 `--emit-c`）

方法：同一份源码、两台 C3.exe（PRE = 未含本表与 Print/Write 分派的树；POST = 当前树），
在 580 份 emit 输入上各跑一次，比较 `ascii256 + rc + bytes`。

结果：**580/580 按路径一致；14 个路径有差异**（13 个真实发码变化 + 1 个运行期抖动）。
变化处逐条判定：

| 路径 | bytes (PRE→POST) | 变化的形状 | 该不该变 |
|---|---|---|---|
| `tests/b174_numstrarg/B174.vbp` | 6586→6440 | `LenB` 答 Long：`vb6_CStr(vb6_VariantFromValue(...))`→`vb6_CStrLong(...)`；两处 `vb6_VARIANT _vcmp_* = vb6_VariantFromValue(...)` 中间量消失（直接比较）。本对 emit 里 `vb6_ChkLong` 两侧均为 0 | **该变**（旧发码把已知 Long 当 Variant 包了一层，行为等价但更慢且错档） |
| `tests/b174_numstrarg/B174Main.bas` | 6511→6365 | 同上（同一工程的模块入口；实测 `vb6_VariantFromValue` 3→0、`vb6_CStrLong` 16→17） | **该变** |
| `tests/BalloonTooltips/prjBalloonTooltips.vbp` | 71439→71400 | `vb6_ChkLong((int32_t)sizeof(...))` → 去闸（`LenB` 已是 Long） | **该变** |
| `tests/guisample/GUISample.vbp` | 305759→305756 | `vb6_VariantFromValue(vb6_InputString(...))`→`vb6_VariantString(...)` | **该变**（同义，后者省一层 `_Generic` 分派） |
| `tests/guisample/vb6ref/src/vc6/GUISample.vbp` | 288334→288331 | 同上 | **该变** |
| `tests/guisample/vb6ref/src/vs2015/GUISample.vbp` | 286979→286976 | 同上 | **该变** |
| `tests/VBFlexGridDemo/Builds/VBFlexGrid/VBFlexGridBase.bas` | 61126→61048 | `LenB` 答 Long 后去哨兵/装箱 | **该变** |
| `tests/VBFlexGridDemo/Builds/VTableHandle.bas` | 89424→89422 | 同上 | **该变** |
| `tests/VBFlexGridDemo/Common/Common.bas` | 195107→195042 | 同上 | **该变** |
| `tests/VBFlexGridDemo/Common/VisualStyles.bas` | 73862→73823 | 同上 | **该变** |
| `tests/VBFlexGridDemo/VBFlexGridDemo.vbp` | 5587740→5587101 | 同上（聚合工程） | **该变** |
| `tests/VbQRCodegen-master/src/mdQRCodegen.bas` | 259249→259062 | `LenB` 去闸：`vb6_ChkLong` 55→52、`vb6_VariantFromValue` 27→25（Input 侧计数未变） | **该变** |
| `tests/VbQRCodegen-master/test/Project1.vbp` | 280393→280206 | 同上（聚合工程） | **该变** |
| `tests/ctrltoolbar/TbApp.vbp` | 25904→0（rc 0→1） | 无发码变化：PRE/POST 单独复跑均 rc=0、emit 逐字节相同 | **不该变**（manifest 运行期抖动；已单独复跑证否） |

结论：**只有该变的地方变**。两个独立读数：
1. **走 `Print #` / `Write #` 的 18 份产物、87 个项站点：0 处因本表而变化**
   （b598/b599 的 emit 与 PRE 逐字节相同）。87 站的类型分布见 §7.2。
2. 13 个真实变化全部可追到 `LenB`/`Input` 一个原因：`LenB` 从 Variant 兜底变成 Long，
   `Input` 从 Variant 兜底变成 String。其余 566 份逐字节未动、缺席 0 / 多出 0。

### 6.3 符号表里仍答 Variant 的 builtin（只列清单，名单逻辑不动）

**(A) 语义层登记为 `Vb6Type::Variant`（12 枚）**：

| 名字 | 为什么保持 Variant |
|---|---|
| `Split` | 返回 `Variant()`（数组），元素档由运行时决定 |
| `CDec` | Decimal 在 C 侧就是整只 `vb6_VARIANT`；且实参是 Variant |
| `CVar` | 就是"转成 Variant" |
| `CVErr` | 返回错误值 Variant |
| `Array` | 返回 `Variant()` |
| `GetAllSettings` | 返回 `Variant`（注册表值档运行时决定） |
| `IIf` / `If` | 多态：返回档由两支实参决定 |
| `CallByName` | 调用结果的档运行时决定 |
| `Filter` | 返回 `Variant()` |
| `Choose` / `Switch` | 多态：返回档由被选中的那一支决定 |

**(B) 语义层没有符号、但 RTL 事实返回 `vb6_VARIANT`（3 枚）**：
`LoadResData`、`LoadResString`、`LoadResPicture`。

### 6.4 第二步：字符串兜底名单已删除（`variantPrefixes` 仍保留）

**已删**（`src/backend/stmt/cgen_file_io.cpp`）：PrintStmt 与 WriteStmt 各一份 34 名
`bstrFuncs` 名字表 + `isBstrExpr` lambda。分流改由 `isStringItem` 两问回答：

1. `inferExprType(e) == Vb6Type::String`——builtin 的返回档由 §6.1 的表答出，表本身查
   RTL 声明；
2. 发射形状本身证明是 BSTR 槽：`vb6_BSTR*` 构造/读取族、`vb6_IIfBSTR`、
   `VB6_SA_AT(BSTR,` 数组元素、`knownBstrVars_` 静态 String 变量。

名字表不再存在：**类型答不出时不许猜**。

删表前落实的两件事：
- `Print #`/`Write #` 会碰到的 builtin 返回档，语义层或 §6.1 的表都能答（§6.3(A) 的
  12 枚只有 `IIf`/`Choose`/`Switch` 这类多态会落到形状判据，而它们的串支形状是
  `vb6_IIfBSTR`/`VB6_SA_AT` 一族，两问仍能判准）；
- `App.Title`/`EXEName`/`Path`/`HelpFile` 四个 BSTR 成员在 `cgen_util_type.cpp` 的
  `MemberAccessExpr` 里直接答 `String`（权威 = `vb6rtl_system.c` 的四个 BSTR 出口），
  不再靠名字表认。删表后 `Print #f, App.Title` 曾写出 `999655096`（BSTR 指针截断），
  就是这一条补上后才写回 `[B600 Title]`。

**未删**：`cExprIsVariant` 的 `variantPrefixes` 前缀表（91 个调用点 / 31 个文件）。
它回答的是"C 表达式是不是**整只槽**"，而 C 文本回答不了这个问题：`Variant` 数组元素读
交回的是**指针**，`VariantFromValue` 又按变目决定交槽与否，只能按发射形状认。删它需要
AST 级槽 oracle，按"不许硬上 / 不许猜"的口径保留，本轮不动。

兜住第二步的断言：
- `scripts/check_builtin_slot_returns.ps1`：R6（`bstrFuncs|isBstrExpr` 回潮即红）+
  R7（两处 `isStringItem` 必须①问 `inferExprType` ②四个形状判据齐全）；
- `scripts/check_app_title_authority.ps1`：R8 名单 6→4 站、`cgen_file_io.cpp` 登记 0
  （回潮即红）+ R8b 四条运行期证据（含 B600 的 `App.Title` 写盘行）；
- `tests/b600_print_string_dispatch/`：22 项，19 项钉曾经靠名字兜底的 builtin 逐个走
  字符串路径 + `App.Title`，末尾两项是**反用例**（数值实参仍走数值路径，保留
  `Str` 的前导/尾随空格）；
- 18 份走 `Print #`/`Write #` 的产物、删表前后逐字节 A/B：`diff=0`（87 站 0 处变化）。

## 7. `inferExprType` 适用边界（本次实际接上了哪几类）

统计对象：580 份 emit 输入展开出的 **476 份 VB 源文件**（`.bas`/`.cls`/`.frm`），按调用
名计数（大小写不敏感）。Print#/Write# 站点数与类型分布来自 18 份实际产物（87 站点）。

### 7.1 答不准/不能直接当数值的形态

| 形态 | 频度 | `inferExprType` 作答 | 本次路径 | 为什么判不准 |
|---|---|---|---|---|
| `IIf(...)` | 63 | Variant（兜底答案） | **旧路径** `vb6_Str((int32_t))` | 返回档由两支实参决定；发射形状是 `vb6_IIfLong/BSTR/...`，无法从类型判断装的是不是数值 |
| `Choose(...)` | 9 | Variant | 旧路径 | 同上（多态） |
| `Switch(...)` | 2 | Variant | 旧路径 | 同上（多态） |
| `CallByName(...)` | 1 | Variant | 旧路径 | 调用结果的档运行时决定 |
| `LoadResData(...)` | 1 | Variant | 有槽证据则 `vb6_FileVariantText`，否则旧路径 | 类型是 Variant（事实），但"是不是整只槽"要看发射形状 |
| `LoadResString(...)` | 2 | Variant | 同上 | 同上 |
| `LoadResPicture(...)` | 0 | Variant | 同上 | 同上 |
| 数组元素 `arr(i)` | 129 处 `VB6_SA_AT(`（18 份产物里的总出现数） | Variant（元素是槽） | 按发射形状：槽证据命中时 → `vb6_FileVariantText`；否则旧路径 | 类型是 Variant 且发射形状能看出"槽"时，两问都成立、可判准；看不到槽形状就不许猜 |
| 对象/外部类成员（无类型表） | — | Unknown | 旧路径 | 没有静态类型来源，不许猜 |
| 用户函数返回 Variant | — | Variant | 只有发射形状是槽时才走槽路径 | 同上 |

### 7.2 本次实际接上的三类（含 87 站的类型分布）

1. **静态标量 7 族 + Date + Boolean + Decimal** → `vb6_File*Text`（主任务）。
2. **有静态整只槽证据的 Variant**（Variant 变量、`LoadRes*`、COM 结果、数组元素）
   → `vb6_FileVariantText`。
3. **String/BSTR** → 字符串路径（第二步后由 `isStringItem` 两问判定，名字表已删）。

18 份走 `Print #` / `Write #` 的产物共 **87 个项站点**，按实际发射分类：
`String/BSTR 50`、`Long 12`、`Num 10`、`Date 5`、`Variant 5`、`Bool 3`、
**旧路径 2**。旧路径的 2 处都是数组元素：`vb6_PrintText(fnum, vb6_Str(VB6_SA_AT(int16_t, scores, i)), 0)`
——`inferExprType` 对数组元素只答 Variant，而发射形状已经写明元素是 `int16_t`，属于
"可判但本刀不接"（见 §7.1 的数组元素行）。这 2 处因此没有 `Str` 的前导/尾随空格，
是**已知残留**；接它需要类型侧认元素类型，属于下一刀。

未接上：`IIf/Choose/Switch/CallByName` 等（保持旧路径，代码注释已写明判不准的原因）。

### 7.3 正反用例（`.temp/m4/pw_bound/Bound.bas`）

整个用例按顺序写盘，`bound.txt` 实测 17 字节：`AB 42 XY 7 1aZZ\r\n`。

| 用例 | 期望路径 | 实际发射 | 该项写出的字节 |
|---|---|---|---|
| `Dim v As Variant: v = "AB": Print #f, v;` | 槽路径 | `vb6_FileVariantText(v, 0)` | `AB`（原样，无补空格） |
| `v = 42: Print #f, v;` | 槽路径（运行期分档） | 同上 | ` 42 `（前导+尾随） |
| `Dim u As T: u.s = "XY": Print #f, u.s;`（UDT String 字段） | 字符串路径 | `vb6_PrintText(f, u.s, 1)` | `XY` |
| `Print #f, u.n;`（UDT Long 字段，`u.n = 7`） | 数值路径 | `vb6_PrintText(f, vb6_FileLongText((int64_t)(u.n), 0), 1)` | ` 7 ` |
| `Print #f, IIf(True, 1, 2);` | **旧路径**（判不准） | `vb6_PrintText(f, vb6_Str((int32_t)(vb6_IIfLong(...))), 1)` | `1`（**无前导/尾随空格**） |
| `Print #f, IIf(True, "a", "b");` | 字符串路径（旧名单兜底） | `vb6_IIfBSTR(...)` → `vb6_PrintText` | `a` |
| `P "ZZ"` → `Print #1, s;`（ByRef String 形参） | 字符串路径（指针不拆） | `vb6_PrintText(1, (*s), 1)` | `ZZ` |

正反判定的意思：该走字符串的走字符串（`u.s`、ByRef String、IIf 字符串支），不该走的
（`u.n`、IIf 数值支）证明仍走数值/旧路径。`u.n` 走数值后拿到 ` 7 `，证明"静态 Long
字段没有被当成 Variant 槽"；`v` 走槽后运行期分档出 ` 42 `，证明"有槽证据时按值语义"。

## 8. 夹具与登记

- `tests/b598_print_numeric/`（Print #，20 条断言，含 `[ 123 ]` 这种方括号包住读回行的
  形式，前导/尾随空格可见）；`tests/b599_write_numeric/`（Write #，独立断言集，16 条）。
  两例均已在 `tests/run_tests.ps1` 登记。
- emit 语料登记：`emit-manifest.expected.txt` 追加 4 行（各工程的 `.vbp` + 主模块）。
  §6.2 的 13 个存量行的哈希来自 CI 那台（本地 PRE 与仓里期望一致），本刀的 13 处发码变化
  **尚未重登记**，等 CI 侧 `rebless_emit_manifest.ps1 -Bless`。
- 回归门（本批）：`tests\baseline.ps1` 3/0/0、`test.bat gui` 8/0/0、`test.bat syntax`
  179/0/0、`test.bat compile` 131/0/0（= 129 存量 + b598/b599 两条）、`manifest_coverage`
  K1..K3 绿（580/580）。
</content>
