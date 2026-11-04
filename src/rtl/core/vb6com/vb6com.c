// vb6com.c - VB6 COM互操作运行时实现 (P6)
// 使用Windows原生COM API, 独立于vb6rtl.h避免VARIANT冲突

#include "vb6com.h"
#include "vb6com_internal.h"
#include "vb6forms_controls.h"   /* Vb6OcxRef + vb6_OcxRefGetTable (manifest 诊断用)
                                     * 注: 产物编译时 RTL 源是平铺在同一 rtl 目录的（见 c3rtl.rc
                                     * 155 / rtl_embedded.cpp 的 RTL_VB6FORMS_CONTROLS_H），
                                     * 故这里用平铺名而不是树内的 ../vb6forms/ 相对路径。 */
#include <stdio.h>

#include <stdlib.h>
#include <string.h>
#include <oleauto.h>   // ITypeLib / ITypeInfo / LoadTypeLibEx / REGKIND_NONE


// P24-08: COM错误传播 — 将HRESULT/EXCEPINFO转换为VB6运行时错误
// VB6行为: COM方法失败时自动触发Err.Raise, 可被On Error捕获
extern void vb6_RaiseError(int32_t errNum, void* description);
extern int32_t vb6_err_resume_next;

// COM错误→VB6错误转换辅助函数
// hr: Invoke返回的HRESULT
// excep: EXCEPINFO结构 (可能包含scode/bstrDescription)
// context: 调用上下文 (用于默认错误描述, 如L"ComCall" / L"ComSetProp")
// VB6标准错误号→描述映射
// 数据源: Microsoft官方Trappable Errors列表 + MSVBVM60.DLL STRINGTABLE交叉验证
// DLL STRINGTABLE ID 规则:
//   - 经典运行时错误(seq 0~39): ID = 10000 + 序列号 (序列号≠错误号, 需交叉匹配)
//   - 窗体/控件/扩展错误: ID部分直接对应错误号, 部分需对照Microsoft文档
// 翻译策略: 中文描述为主, 英文原文保留在注释中供参考
// 注: 735/744/746 和 485/486/487 都是合法的VB6错误号, 不可互相替代
// 非 static: vb6com_collection.c 的内建 Collection 也要用它填 EXCEPINFO.bstrDescription
const wchar_t* vb6_StdErrorDesc(int32_t errNum) {
    switch (errNum) {
    // ---- VB6 经典运行时错误 (ErrNum 3~94) ----
    // 来源: Microsoft官方Trappable Errors + DLL STRINGTABLE ID 10000~10039
    case   3: return L"没有GoSub时的Return";              // Return without GoSub
    case   5: return L"无效的过程调用或参数";               // Invalid procedure call or argument
    case   6: return L"溢出";                              // Overflow
    case   7: return L"内存不足";                           // Out of memory
    case   9: return L"下标越界";                           // Subscript out of range
    case  10: return L"此数组被固定或暂时锁定";              // This array is fixed or temporarily locked
    case  11: return L"除数为零";                           // Division by zero
    case  13: return L"类型不匹配";                          // Type mismatch
    case  14: return L"字符串空间不足";                       // Out of string space
    case  16: return L"表达式太复杂";                        // Expression too complex
    case  17: return L"不能执行请求的操作";                   // Can't perform requested operation
    case  18: return L"发生用户中断";                         // User interrupt occurred
    case  20: return L"没有错误时Resume";                     // Resume without error
    case  28: return L"栈空间不足";                           // Out of stack space
    case  35: return L"Sub、Function或Property未定义";        // Sub, Function, or Property not defined
    case  47: return L"DLL应用程序客户端过多";                 // Too many DLL application clients
    case  48: return L"加载DLL时出错";                        // Error in loading DLL
    case  49: return L"DLL调用约定错误";                      // Bad DLL calling convention
    case  51: return L"内部错误";                             // Internal error
    case  52: return L"文件名或文件号错误";                     // Bad file name or number
    case  53: return L"找不到文件";                            // File not found
    case  54: return L"文件模式错误";                          // Bad file mode
    case  55: return L"文件已打开";                            // File already open
    case  57: return L"设备I/O错误";                          // Device I/O error
    case  58: return L"文件已存在";                            // File already exists
    case  59: return L"记录长度错误";                          // Bad record length
    case  61: return L"磁盘已满";                             // Disk full
    case  62: return L"输入超出文件尾";                        // Input past end of file
    case  63: return L"记录号错误";                            // Bad record number
    case  67: return L"文件过多";                             // Too many files
    case  68: return L"设备不可用";                            // Device unavailable
    case  70: return L"权限被拒绝";                            // Permission denied
    case  71: return L"磁盘未准备好";                          // Disk not ready
    case  74: return L"不能用不同的驱动器重命名";                // Can't rename with different drive
    case  75: return L"路径/文件访问错误";                      // Path/File access error
    case  76: return L"找不到路径";                            // Path not found
    case  91: return L"对象变量或With块变量未设置";              // Object variable or With block variable not set
    case  92: return L"For循环未初始化";                       // For loop not initialized
    case  93: return L"无效的模式字符串";                       // Invalid pattern string
    case  94: return L"Null的使用无效";                        // Invalid use of Null

    // ---- VB6 扩展运行时错误 (ErrNum 96~98) ----
    case  96: return L"无法接收对象事件,已达最大接收数";         // Unable to sink events of object because it's already firing events to the maximum number of event receivers
    case  97: return L"不能在不是定义类实例的对象上调用Friend过程"; // Can't call Friend procedure on an object that is not an instance of the defining class
    case  98: return L"属性或方法调用不能包含对私有对象的引用";     // A property or method call cannot include a reference to a private object, either as an argument or as a return value

    // ---- 系统DLL/文件格式错误 (ErrNum 298~328) ----
    // 来源: Microsoft官方Trappable Errors + DLL STRINGTABLE交叉验证
    case 298: return L"不能加载系统DLL";                      // System DLL could not be loaded
    case 320: return L"不能在指定文件名中使用字符设备名";         // Can't use character device names in specified file names
    case 321: return L"无效的文件格式";                         // Invalid file format
    case 322: return L"无法创建必要的临时文件";                  // Can't create necessary temporary file
    case 325: return L"资源文件中格式无效";                      // Invalid format in resource file
    case 327: return L"找不到命名的数据值";                      // Data value named not found
    case 328: return L"非法参数;不能写入数组";                   // Illegal parameter; can't write arrays

    // ---- 注册表/组件错误 (ErrNum 335~338) ----
    case 335: return L"不能访问系统注册表";                      // Could not access system registry
    case 336: return L"ActiveX部件未正确注册";                  // Component not correctly registered
    case 337: return L"找不到ActiveX部件";                      // Component not found
    case 338: return L"ActiveX部件不能正确运行";                // Component did not run correctly

    // ---- 控件数组错误 (ErrNum 339~345) ----
    // DLL STRINGTABLE ID 1339~1345, VB6错误号 = ID - 1000
    case 339: return L"组件或其依赖项未正确注册:文件缺失或无效";   // Component '|' or one of its dependencies not correctly registered: a file is missing or invalid
    case 340: return L"控件数组元素不存在";                      // Control array element '|' doesn't exist
    case 341: return L"控件数组索引无效";                         // Invalid control array index
    case 342: return L"没有足够空间分配控件数组";                  // Not enough room to allocate control array
    case 343: return L"对象不是数组";                             // Object not an array
    case 344: return L"必须指定对象数组的索引";                    // Must specify index for object array
    case 345: return L"已达上限:无法为此窗体创建更多控件";          // Reached limit: cannot create any more controls for this form

    // ---- 窗体/对象加载错误 (ErrNum 360~374) ----
    case 360: return L"对象已加载";                              // Object already loaded
    case 361: return L"不能加载或卸载此对象";                     // Can't load or unload this object
    case 362: return L"不能卸载设计时创建的控件";                  // Can't unload controls created at design time
    case 363: return L"找不到指定的ActiveX控件";                  // Control specified not found
    case 364: return L"对象已卸载";                               // Object was unloaded
    case 365: return L"无法在此上下文中卸载";                      // Unable to unload within this context
    case 366: return L"没有可用的MDI窗体来加载";                   // No MDI form available to load
    case 367: return L"不能加载或注册ActiveX控件";                 // Can't load (or register) ActiveX control: '|'
    case 368: return L"指定文件已过期;此程序需要较新版本";          // The specified file is out of date. This program requires a later version
    case 369: return L"在ActiveX DLL中操作无效";                  // Operation not valid in an ActiveX DLL
    case 371: return L"指定的对象不能用作Show的所有者窗体";         // The specified object can't be used as an owner form for Show
    case 372: return L"从文件加载控件失败;版本可能已过期";          // Failed to load control '|' from |2. Your version may be outdated
    case 373: return L"编译环境与运行时组件之间的交互不支持";       // Interaction between compiled and design environment components is not supported
    case 374: return L"激活控件失败;控件可能不兼容";               // Failed to activate control '|'. This control may be incompatible

    // ---- 属性错误 (ErrNum 378~399) ----
    case 378: return L"加载或保存时不能设置此属性";                // '|' cannot be set while loading or saving
    case 379: return L"不能在属性页上放置默认或取消按钮";           // You can't put a Default or Cancel button on a Property Page
    case 380: return L"属性值无效";                               // Invalid property value
    case 381: return L"属性数组索引无效";                          // Invalid property array index
    case 382: return L"运行时不能设置此属性";                      // Property Set can't be executed at run time
    case 383: return L"只读属性不能设置";                          // Property Set can't be used with a read-only property
    case 384: return L"窗体最小化或最大化时不能移动或调整大小";      // A form can't be moved or sized while minimized or maximized
    case 385: return L"需要属性数组索引";                          // Need property array index
    case 387: return L"不允许设置此属性";                          // Property Set not permitted
    case 388: return L"不能从父菜单设置Visible属性";               // Can't set Visible property from a parent menu
    case 393: return L"运行时不能获取此属性";                      // Property Get can't be executed at run time
    case 394: return L"只写属性不能读取";                          // Property Get can't be executed on write-only property
    case 395: return L"不能将分隔条用作此控件的菜单名";             // Cannot use separator bar as menu name for this control
    case 396: return L"在属性页中不能设置此属性";                   // Property cannot be set within a page
    case 397: return L"合并时不能加载、卸载或设置顶层菜单的Visible属性"; // Can't load, unload, or set Visible property for top level menus while they are merged
    case 398: return L"客户端站点不可用";                           // Client Site not available

    // ---- 模态窗体错误 (ErrNum 400~406) ----
    case 400: return L"窗体已显示;不能以模态方式显示";              // Form already displayed; can't show modally
    case 401: return L"显示模态窗体时不能显示非模态窗体";           // Can't show non-modal form when modal form is displayed
    case 402: return L"必须先关闭或隐藏最上层模态窗体";             // Must close or hide topmost modal form first
    case 403: return L"MDI窗体不能以模态方式显示";                  // MDI forms cannot be shown modally
    case 404: return L"MDI子窗体不能以模态方式显示";                // MDI child forms cannot be shown modally
    case 405: return L"无法在此上下文中显示模态窗体";               // Unable to show modal form within this context
    case 406: return L"不能从ActiveX DLL中显示非模态窗体";          // Non-modal forms cannot be displayed from an ActiveX DLL

    // ---- 对象权限错误 (ErrNum 419~425) ----
    case 419: return L"拒绝使用对象的权限";                        // Permission to use object denied
    case 422: return L"找不到属性";                                // Property not found
    case 423: return L"找不到属性或方法";                          // Property or method not found
    case 424: return L"需要对象";                                  // Object required
    case 425: return L"无效的对象使用";                            // Invalid object use

    // ---- COM/ActiveX错误 (ErrNum 429~463) ----
    case 429: return L"ActiveX部件不能创建对象或返回引用";          // Component can't create object or return reference to this object
    case 430: return L"类不支持自动化";                             // Class doesn't support Automation
    case 432: return L"自动化操作期间找不到文件名或类名";            // File name or class name not found during Automation operation
    case 438: return L"对象不支持此属性或方法";                      // Object doesn't support this property or method
    case 440: return L"自动化错误";                                 // Automation error
    case 442: return L"远程处理的类型库或对象库的连接已丢失";         // Connection to type library or object library for remote process has been lost
    case 443: return L"自动化对象没有默认值";                        // Automation object doesn't have a default value
    case 445: return L"对象不支持此操作";                            // Object doesn't support this action
    case 446: return L"对象不支持命名参数";                          // Object doesn't support named arguments
    case 447: return L"对象不支持当前区域设置";                      // Object doesn't support current locale setting
    case 448: return L"找不到命名参数";                              // Named argument not found
    case 449: return L"参数不是可选的或属性赋值无效";                 // Argument not optional or invalid property assignment
    case 450: return L"参数个数错误或属性赋值无效";                   // Wrong number of arguments or invalid property assignment
    case 451: return L"对象不是集合";                                // Object not a collection
    case 452: return L"无效的序号";                                  // Invalid ordinal
    case 453: return L"找不到指定的DLL函数";                         // Specified DLL function not found
    case 454: return L"找不到代码资源";                              // Code resource not found
    case 455: return L"代码资源锁定错误";                            // Code resource lock error
    case 457: return L"此键已经与该集合的一个元素关联";               // This key is already associated with an element of this collection
    case 458: return L"变量使用了Visual Basic不支持的类型";          // Variable uses a type not supported in Visual Basic
    case 459: return L"此组件不支持事件集";                          // This component doesn't support the set of events
    case 460: return L"剪贴板格式无效";                              // Invalid Clipboard format
    case 461: return L"找不到方法或数据成员";                         // Method or data member not found
    case 462: return L"远程服务器机器不存在或不可用";                 // The remote server machine does not exist or is unavailable
    case 463: return L"类未在本地机器上注册";                         // Class not registered on local machine

    // ---- 图片/打印机错误 (ErrNum 480~487) ----
    case 480: return L"不能创建AutoRedraw图像";                      // Can't create AutoRedraw image
    case 481: return L"图片无效";                                    // Invalid picture
    case 482: return L"打印机错误";                                  // Printer error
    case 483: return L"打印机驱动程序不支持指定属性";                 // Printer driver does not support specified property
    case 484: return L"从系统获取打印机信息时出错";                   // Problem getting printer information from the system
    case 485: return L"无效的图片类型";                               // Invalid picture type
    case 486: return L"不能用这种类型的打印机打印窗体图像";            // Can't print form image to this type of printer
    case 487: return L"不能打印最小化的窗体图像";                     // Can't print minimized form image

    // ---- 剪贴板错误 (ErrNum 520~521) ----
    case 520: return L"不能清空剪贴板";                               // Can't empty Clipboard
    case 521: return L"不能打开剪贴板";                               // Can't open Clipboard

    // ---- 其他运行时错误 (ErrNum 735~746) ----
    // 注: 735/744/746 与 485/486/487 是不同的VB6错误号, 两者都合法
    case 735: return L"不能将文件保存到TEMP目录";                      // Can't save file to TEMP directory
    case 744: return L"找不到搜索文本";                               // Search text not found
    case 746: return L"替换内容太长";                                 // Replacements too long

    // ---- OLE/嵌入对象错误 (ErrNum 31001~31037) ----
    // 来源: Microsoft官方Trappable Errors列表 (1-1000范围外)
    case 31001: return L"内存不足";                                   // Out of memory
    case 31004: return L"无对象";                                     // No object
    case 31018: return L"未设置类";                                   // Class is not set
    case 31027: return L"不能激活对象";                               // Unable to activate object
    case 31032: return L"不能创建嵌入对象";                            // Unable to create embedded object
    case 31036: return L"保存到文件时出错";                            // Error saving to file
    case 31037: return L"从文件加载时出错";                            // Error loading from file

    default: return NULL;
    }
}


/* 账 #305 §B185/§B186: 「一条 VB 运行期错误 crossed 服务边界该长成哪个 HRESULT」的答案, 与它的反向。
 *   VB6 那台机器交回的就是 0x800A0000|错误号 —— 造与认这两问从此只住下面这两枚函数:
 *   发腿 (vb6_ComFailExcep 填 EXCEPINFO.scode、类工厂那趟没有 EXCEPINFO 通道的路) 问前者,
 *   收腿 (vb6_ComCheckError 与 vb6_CreateObject 那条 create 腿) 问后者。
 * vb6_ComVbErrFromHr 认不出那一档就答 0: 0x8004xxxx (REGDB_E_CLASSNOTREG / CLASS_E_NOTLICENSED)
 * 与 0x8007xxxx 这些**真·激活失败**不许被低 16 位当成 VB 错误号读出一串垃圾。 */
#define VB6_COM_VB_HR_FACILITY 0x800A0000u

HRESULT vb6_ComVbHresult(int32_t errNum) {
    return (HRESULT)(VB6_COM_VB_HR_FACILITY | (unsigned)(errNum & 0xFFFF));
}

int32_t vb6_ComVbErrFromHr(HRESULT hr) {
    unsigned long u = (unsigned long)hr;
    if ((u & 0xFFFF0000UL) != VB6_COM_VB_HR_FACILITY) return 0;
    return (int32_t)(u & 0xFFFFUL);
}

// COM错误→VB6错误转换辅助函数
// hr: Invoke返回的HRESULT
// excep: EXCEPINFO结构 (可能包含scode/bstrDescription)
// context: 调用上下文 (用于默认错误描述, 如L"ComCall" / L"ComSetProp")
// Fix 177: HRESULT → VB6 错误号。
// "低16位即VB6错误号" 只对 0x8004xxxx (标准 COM) 与 0x800Axxxx (VB6 运行时) 成立,
// 对 DISP_E_* 系列 (0x8002xxxx) 不成立 —— 其低16位落在 DISPID 子码空间: 例如
// DISP_E_MEMBERNOTFOUND (0x80020003) 的低16位 = 3, 会被当成 VB6 错误 3
// "没有GoSub时的Return", 与真实原因 (对象不支持该属性/方法) 毫无关系。
// 实测: VBMAN 对 Scripting.Dictionary 调自有扩展 API 时即报出这个假象,
// 把排查方向引到 GoSub/Return 上。此处显式映射 DISP_E_* 到 VB6 官方错误号
// (描述表 vb6_StdErrorDesc 中 438/449/450/451 均已有条目)。
static int32_t vb6_HresultToErrNum(long hr) {
    /* 账 #305 §B186: 「这一枚 hr 是不是服务端交回的 VB 错误、是哪一号」只问上面那唯一一枚认腿,
     * 不在这里再抄一遍低 16 位的算法 (抄出来的两份会漂)。认不出那一档才往下走原路。 */
    int32_t vbErr = vb6_ComVbErrFromHr((HRESULT)hr);
    if (vbErr) return vbErr;
    switch ((unsigned long)hr) {
    case 0x80020003UL: return 438;  // DISP_E_MEMBERNOTFOUND → 对象不支持此属性或方法
    case 0x80020006UL: return 438;  // DISP_E_UNKNOWNNAME    → 同上
    case 0x80020005UL: return 13;   // DISP_E_TYPEMISMATCH   → 类型不匹配
    case 0x8002000EUL: return 450;  // DISP_E_BADPARAMCOUNT  → 参数个数错误
    case 0x8002000FUL: return 449;  // DISP_E_PARAMNOTOPTIONAL → 参数不是可选的
    case 0x8002000BUL: return 451;  // DISP_E_BADINDEX       → 对象不是集合
    default: {
        int32_t errNum = (int32_t)((unsigned long)hr & 0xFFFF);  // VB6错误号 = HRESULT低16位
        if (errNum == 0) errNum = (int32_t)hr;  // 非标准HRESULT直接用整个值
        return errNum;
    }
    }
}

void vb6_ComCheckError(HRESULT hr, EXCEPINFO* excep, const wchar_t* context) {
    int32_t errNum = vb6_HresultToErrNum((long)hr);

    // 优先使用EXCEPINFO中的scode (OLE自动化错误码; DISP_E_EXCEPTION 时携带真实错误号)
    if (excep && excep->scode) {
        errNum = vb6_HresultToErrNum((long)excep->scode);
    }

    // 构造错误描述BSTR — 四级fallback:
    //   1. EXCEPINFO.bstrDescription (COM对象直接提供)
    //   2. IErrorInfo::GetDescription (COM错误接口)
    //   3. VB6标准错误号中文查表 (scode有效时优先)
    //   4. FormatMessageW(FROM_SYSTEM) (系统HRESULT描述)
    //   5. 默认: "COM error in <context>: 0xXXXXXXXX"
    BSTR desc = NULL;
    if (excep && excep->bstrDescription) {
        // 1. EXCEPINFO有描述: 直接使用
        desc = excep->bstrDescription;
        excep->bstrDescription = NULL;  // 转移所有权
    } else {
        // 2. 尝试IErrorInfo
        IErrorInfo* perror = NULL;
        if (SUCCEEDED(GetErrorInfo(0, &perror)) && perror) {
            BSTR eiDesc = NULL;
            perror->lpVtbl->GetDescription(perror, &eiDesc);
            DWORD helpCtx = 0;
            perror->lpVtbl->GetHelpContext(perror, &helpCtx);
            if (helpCtx) errNum = (int32_t)(helpCtx & 0xFFFF);
            perror->lpVtbl->Release(perror);
            if (eiDesc) desc = eiDesc;
        }
        if (!desc) {
            // 3. VB6标准错误号中文查表 (scode提取的VB6错误号优先)
            const wchar_t* stdDesc = vb6_StdErrorDesc(errNum);
            if (stdDesc) desc = SysAllocString(stdDesc);
        }
        if (!desc) {
            // 4. FormatMessageW(FROM_SYSTEM) — 仅当VB6表查不到时使用
            //    注意: DISP_E_EXCEPTION(0x80020009)会返回"发生意外"等无用通用描述
            //    所以当scode有效(VB6错误号)时不应走到这里
            wchar_t sysBuf[512] = {0};
            DWORD fmLen = FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                NULL, (DWORD)hr, 0, sysBuf, 512, NULL);
            if (fmLen > 0) {
                while (fmLen > 0 && (sysBuf[fmLen-1] == L'\r' || sysBuf[fmLen-1] == L'\n')) sysBuf[--fmLen] = L'\0';
                if (fmLen > 0) desc = SysAllocString(sysBuf);
            }
        }
        if (!desc) {
            // 5. 默认描述
            wchar_t buf[256];
            swprintf(buf, 256, L"COM error in %ls: 0x%08lX", context ? context : L"?", (unsigned long)hr);
            desc = SysAllocString(buf);
        }
    }

    // 清理EXCEPINFO中的其他BSTR (bstrDescription已转移)
    if (excep) {
        if (excep->bstrSource) { SysFreeString(excep->bstrSource); excep->bstrSource = NULL; }
        if (excep->bstrHelpFile) { SysFreeString(excep->bstrHelpFile); excep->bstrHelpFile = NULL; }
    }

    vb6_RaiseError(errNum, desc);
}

HRESULT vb6_ComFailExcep(EXCEPINFO* pexcep, int32_t errNum,
                         const wchar_t* desc, const wchar_t* fallback) {
    if (pexcep) {
        pexcep->scode = (SCODE)vb6_ComVbHresult(errNum);
        pexcep->bstrSource = NULL;
        pexcep->bstrHelpFile = NULL;
        pexcep->dwHelpContext = 0;
        /* 三级缺省:  raiser 自己带的那句(服务端传进来的 vb6_ErrDescription) -> 那张 VB 标准错误号
         *  描述表 -> 调用方给的兜底。中间这一级不许跳: 没有 Err.Raise 描述的那一形(内在错误,
         *  如服务端里 raise 的 5)在真 VB6 里交回的不是空串, 而是那一号的默认描述。
         *  表交的是中文, 所以判据只钉 scode, 不钉这一句的字节。 */
        if (!desc) desc = vb6_StdErrorDesc(errNum);
        if (!desc) desc = fallback;
        pexcep->bstrDescription = SysAllocString(desc ? desc : L"");
    }
    return DISP_E_EXCEPTION;
}
// C-style vtable helpers for IUnknown (avoid C++ IUnknown method call issues)
// ============================================================
HRESULT vb6_UnknownQI(void* obj, REFIID riid, void** ppv) {
    IUnknown* pUnk = (IUnknown*)obj;
    return pUnk->lpVtbl->QueryInterface(pUnk, riid, ppv);
}
ULONG vb6_UnknownRelease(void* obj) {
    IUnknown* pUnk = (IUnknown*)obj;
    return pUnk->lpVtbl->Release(pUnk);
}

// ============================================================
// COM初始化/退出
// ============================================================

// 模式策略（实现在下方 ComLib 表之后：需要读组件的内嵌 manifest）
static void vb6_ComActCtxEnter(void);
static void vb6_ComActCtxLeave(void);

void vb6_ComInit(void) {
    // 免注册 COM 的 ActCtx 必须**先于** CoInitializeEx（因而也先于 OleInitialize 与
    // 任何 CoCreateInstance）就位；direct 模式下这一行是空操作，直加载逐字不变。
    vb6_ComActCtxEnter();
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
}

void vb6_ComExit(void) {
    CoUninitialize();
    // ActCtx 句柄的唯一释放点：进程退出的正常路径（vb6_Exit → 这里）。
    // 非正常退出 (ExitProcess/崩溃) 由操作系统回收，不存在泄漏口径问题。
    vb6_ComActCtxLeave();
}

// ============================================================
// Fix 160: 免注册 COM — vbp ComLib= 声明的组件 DLL
// ============================================================

#define VB6_MAX_COMLIBS 256
static Vb6ComLib g_comLibs[VB6_MAX_COMLIBS];
static int g_comLibCount = 0;

/* C3_COM_TRACE=1 缓存 */
static int comLibTraceEnabled(void) {
    static int cached = -1;
    if (cached < 0) cached = (GetEnvironmentVariableW(L"C3_COM_TRACE", NULL, 0) > 0);
    return cached;
}

/* ============================================================
 * Fix 160-com: 免注册 typelib 缓存 — 让 typelib 驱动的 IDispatch
 * (如 Chilkat) 在 TLB 未注册时也能 GetIDsOfNames/Invoke 成功.
 * C3 在 vb6_ComLibRegister 时按 DLL 路径去重 LoadTypeLib(REGKIND_NONE)
 * (不写注册表), 取出每个 coclass 的默认调度接口 ITypeInfo; 对象创建时
 * 把 对象指针→ITypeInfo 记进 g_comlibObjTI; vb6_getDispid 据此解析
 * DISPID, 绕开对象自身因 TLB 缺失而失败的 GetIDsOfNames.
 * ============================================================ */
#define VB6_COMLIB_TLIBMEM_MAX 32
static struct { wchar_t path[MAX_PATH * 2]; ITypeLib* tlib; } g_comLibTLBCache[VB6_COMLIB_TLIBMEM_MAX];
static int g_comLibTLBCacheCount = 0;

#define VB6_COMLIB_OBJTI_MAX 1024
static struct { void* disp; ITypeInfo* ti; } g_comlibObjTI[VB6_COMLIB_OBJTI_MAX];
static int g_comlibObjTICount = 0;

static ITypeLib* comLibGetTypeLib(const wchar_t* dllPath) {
    if (!dllPath || !*dllPath) return NULL;
    for (int i = 0; i < g_comLibTLBCacheCount; i++)
        if (_wcsicmp(g_comLibTLBCache[i].path, dllPath) == 0)
            return g_comLibTLBCache[i].tlib;
    if (g_comLibTLBCacheCount >= VB6_COMLIB_TLIBMEM_MAX) return NULL;
    ITypeLib* pTLib = NULL;
    HRESULT hr = LoadTypeLibEx((LPOLESTR)dllPath, REGKIND_NONE, &pTLib);
    if (FAILED(hr) || !pTLib) {
        if (comLibTraceEnabled())
            fwprintf(stderr, L"[C3_COM]   LoadTypeLibEx(%ls) failed: 0x%08lX (DISPID 将退回对象自身 GetIDsOfNames)\n",
                     dllPath, (unsigned long)hr);
        return NULL;
    }
    wcscpy(g_comLibTLBCache[g_comLibTLBCacheCount].path, dllPath);
    g_comLibTLBCache[g_comLibTLBCacheCount].tlib = pTLib;  /* 进程期持有, 退出时随进程释放 */
    g_comLibTLBCacheCount++;
    return pTLib;
}

/* 取 coclass 的默认(非 source)调度接口 ITypeInfo —— 用于免注册解析 DISPID. */
static ITypeInfo* comLibLoadDispIface(const wchar_t* dllPath, REFCLSID rclsid) {
    ITypeLib* pTLib = comLibGetTypeLib(dllPath);
    if (!pTLib) return NULL;
    ITypeInfo* pCoclass = NULL;
    if (FAILED(pTLib->lpVtbl->GetTypeInfoOfGuid(pTLib, rclsid, &pCoclass)) || !pCoclass)
        return NULL;
    ITypeInfo* result = NULL;
    TYPEATTR* pTA = NULL;
    if (SUCCEEDED(pCoclass->lpVtbl->GetTypeAttr(pCoclass, &pTA))) {
        UINT cImpl = pTA->cImplTypes;
        for (UINT i = 0; i < cImpl && !result; i++) {
            INT flags = 0;
            if (SUCCEEDED(pCoclass->lpVtbl->GetImplTypeFlags(pCoclass, i, &flags))) {
                if ((flags & IMPLTYPEFLAG_FDEFAULT) && !(flags & IMPLTYPEFLAG_FSOURCE)) {
                    HREFTYPE ref = 0;
                    if (SUCCEEDED(pCoclass->lpVtbl->GetRefTypeOfImplType(pCoclass, i, &ref)))
                        pCoclass->lpVtbl->GetRefTypeInfo(pCoclass, ref, &result);
                }
            }
        }
        pCoclass->lpVtbl->ReleaseTypeAttr(pCoclass, pTA);
    }
    pCoclass->lpVtbl->Release(pCoclass);
    if (result) result->lpVtbl->AddRef(result);  /* 返回独立引用, 由 g_comLibs 表项持有 */
    return result;
}

/* 免注册对象 → 默认调度接口 ITypeInfo (getDispid 解析用). 命中返回 AddRef 过的
 * ITypeInfo*, 调用方须 Release; 未命中返回 NULL. */
ITypeInfo* vb6_ComLibLookupTypeInfo(void* disp) {
    if (!disp) return NULL;
    for (int i = 0; i < g_comlibObjTICount; i++) {
        if (g_comlibObjTI[i].disp == disp && g_comlibObjTI[i].ti) {
            g_comlibObjTI[i].ti->lpVtbl->AddRef(g_comlibObjTI[i].ti);
            return g_comlibObjTI[i].ti;
        }
    }
    return NULL;
}

/* 记录 免注册对象 → 默认调度接口 ITypeInfo (COM 单线程 STA, 无需加锁). */
static void comLibRegisterObjTI(void* disp, ITypeInfo* ti) {
    if (!disp || !ti) return;
    if (g_comlibObjTICount >= VB6_COMLIB_OBJTI_MAX) return;
    ti->lpVtbl->AddRef(ti);  /* 映射项持有独立引用 */
    g_comlibObjTI[g_comlibObjTICount].disp = disp;
    g_comlibObjTI[g_comlibObjTICount].ti = ti;
    g_comlibObjTICount++;
}

void vb6_ComLibRegister(const Vb6ComLib* libs, int count) {
    if (!libs || count <= 0) return;
    if (count > VB6_MAX_COMLIBS) {
        fprintf(stderr, "[C3_COM]   %d 条组件超过上限 %d, 超出部分被截断 (需注册表注册)\n",
                count, VB6_MAX_COMLIBS);
        count = VB6_MAX_COMLIBS;
    }
    memcpy(g_comLibs, libs, (size_t)count * sizeof(Vb6ComLib));
    g_comLibCount = count;
    if (comLibTraceEnabled())
        fprintf(stderr, "[C3_COM]   registered %d component(s) from ComLib=\n", count);

    /* Fix 160-com: 预载每个 coclass 的默认调度接口 ITypeInfo (免注册 TLB 解析用).
     * 仅按 DLL 路径去重 LoadTypeLibEx 一次; 失败则 dispIface=NULL, 后续退回对象自身
     * GetIDsOfNames (对 TLB 已注册/标准组件零影响). */
    {
        wchar_t exeDir[MAX_PATH]; exeDir[0] = 0;
        DWORD n = GetModuleFileNameW(NULL, exeDir, MAX_PATH);
        if (n > 0 && n < MAX_PATH) {
            wchar_t* es = wcsrchr(exeDir, L'\\');
            if (es) es[1] = 0; else exeDir[0] = 0;
        }
        for (int i = 0; i < g_comLibCount; i++) {
            g_comLibs[i].dispIface = NULL;
            const wchar_t* fn = g_comLibs[i].fileName;
            if (!fn || !*fn) continue;
            wchar_t full[MAX_PATH * 2]; full[0] = 0;
            int isAbs = (wcschr(fn, L':') != NULL) || (fn[0] == L'\\');
            if (exeDir[0] && !isAbs && wcslen(exeDir) + wcslen(fn) + 1 < MAX_PATH * 2) {
                wcscpy(full, exeDir); wcscat(full, fn);
            } else {
                wcscpy(full, fn);
            }
            CLSID clsid;
            if (SUCCEEDED(CLSIDFromString((LPOLESTR)g_comLibs[i].clsidStr, &clsid)))
                g_comLibs[i].dispIface = comLibLoadDispIface(full, &clsid);
            if (comLibTraceEnabled() && g_comLibs[i].dispIface)
                fwprintf(stderr, L"[C3_COM]   cached dispIface for %ls\n", g_comLibs[i].progId);
        }
    }
}

// ============================================================
// 免注册 COM 运行期模式：direct | manifest | auto
// ------------------------------------------------------------
// 纪律（本批的硬约束）：
//   * direct  = 既有 LoadLibrary + DllGetClassObject 直加载 —— 一条判定都不改；
//   * manifest = 建 ActCtx；identity 缺失 / 组件无内嵌 assembly manifest ⇒ **明确失败**
//                （打印缺项并退出，绝不静默回退直加载）；
//   * auto    = 默认；能用 manifest 就用，不能用就走 direct 并在 stderr 说明原因/哪条。
// 注入：cgen 只在显式给 -ComMode 时于入口点发 vb6_ComSetMode()（未给 ⇒ 一行不发，
// 产物逐字节不变）；另支持 C3_COM_MODE / C3_COM_MANIFEST 环境变量（等价注入方式，
// 便于同一枚产物在不重编的情况下做正反用例）。
// ============================================================
static int      g_comMode = VB6_COM_MODE_AUTO;
static int      g_comModeExplicit = 0;          /* 程序里显式设过 = 环境变量不覆盖 */
static wchar_t  g_comManifestPath[MAX_PATH * 2] = L"";

void vb6_ComSetMode(int mode) {
    if (mode >= VB6_COM_MODE_AUTO && mode <= VB6_COM_MODE_MANIFEST) {
        g_comMode = mode;
        g_comModeExplicit = 1;
    }
}

void vb6_ComSetManifest(const wchar_t* path) {
    if (!path) { g_comManifestPath[0] = 0; return; }
    wcsncpy(g_comManifestPath, path, MAX_PATH * 2 - 1);
    g_comManifestPath[MAX_PATH * 2 - 1] = 0;
}

int vb6_ComMode(void) {
    if (g_comModeExplicit) return g_comMode;
    static int envMode = -1;
    if (envMode < 0) {
        envMode = VB6_COM_MODE_AUTO;
        wchar_t buf[64];
        DWORD n = GetEnvironmentVariableW(L"C3_COM_MODE", buf, 64);
        if (n > 0 && n < 64) {
            if (_wcsicmp(buf, L"direct") == 0)        envMode = VB6_COM_MODE_DIRECT;
            else if (_wcsicmp(buf, L"manifest") == 0) envMode = VB6_COM_MODE_MANIFEST;
            else if (_wcsicmp(buf, L"auto") == 0)     envMode = VB6_COM_MODE_AUTO;
            /* 其它取值保持 auto（不猜） */
        }
    }
    return envMode;
}

/* 客户端 manifest 路径：显式 > C3_COM_MANIFEST > <exe>.manifest。
 * 返回 1 = 有路径可试（不代表文件存在，存在性由校验函数回答）。 */
static int comManifestPathResolve(wchar_t* out, int cch) {
    out[0] = 0;
    if (g_comManifestPath[0]) {
        wcsncpy(out, g_comManifestPath, (size_t)cch - 1);
        out[cch - 1] = 0;
        return 1;
    }
    static wchar_t envPath[MAX_PATH * 2];
    static int envGot = -1;
    if (envGot < 0)
        envGot = (GetEnvironmentVariableW(L"C3_COM_MANIFEST", envPath, MAX_PATH * 2) > 0) ? 1 : 0;
    if (envGot) {
        wcsncpy(out, envPath, (size_t)cch - 1);
        out[cch - 1] = 0;
        return 1;
    }
    wchar_t exe[MAX_PATH];
    DWORD n = GetModuleFileNameW(NULL, exe, MAX_PATH);
    if (n == 0 || n >= MAX_PATH) return 0;
    wchar_t* dot = wcsrchr(exe, L'.');
    if (dot) *dot = 0;
    _snwprintf(out, (size_t)cch - 1, L"%ls.manifest", exe);
    out[cch - 1] = 0;
    return 1;
}

/* 组件文件路径 → 绝对路径：先按 <exe目录>\<相对路径>，再按原样（与直加载候选同序）。
 * 只为诊断服务（读 RT_MANIFEST 用，不执行组件代码）。 */
static const wchar_t* comComponentPath(const wchar_t* fileName, wchar_t* out, int cch) {
    out[0] = 0;
    if (!fileName || !*fileName) return out;
    int isAbs = (wcschr(fileName, L':') != NULL) || (fileName[0] == L'\\');
    if (!isAbs) {
        wchar_t exeDir[MAX_PATH];
        DWORD n = GetModuleFileNameW(NULL, exeDir, MAX_PATH);
        if (n > 0 && n < MAX_PATH) {
            wchar_t* es = wcsrchr(exeDir, L'\\');
            if (es) {
                es[1] = 0;
                _snwprintf(out, (size_t)cch - 1, L"%ls%ls", exeDir, fileName);
                out[cch - 1] = 0;
                if (GetFileAttributesW(out) != INVALID_FILE_ATTRIBUTES) return out;
            }
        }
    }
    wcsncpy(out, fileName, (size_t)cch - 1);
    out[cch - 1] = 0;
    return out;
}

/* 缺项拼接：多条缺项用 "; " 串起来（诊断口径与 vb6com_manifest.c 的 missing 一致）。 */
static void comMissingJoin(wchar_t* out, int cch, const wchar_t* item) {
    if (!out || cch <= 0 || !item) return;
    if (out[0]) {
        size_t used = wcslen(out);
        if (used + 3 < (size_t)cch) wcscat(out, L"; ");
    }
    size_t used = wcslen(out);
    if (used + wcslen(item) >= (size_t)cch) return;
    wcscat(out, item);
}

/* ------------------------------------------------------------------
 * manifest 聚合（工具包口径，实测结论）：
 *   SxS 认的是**进程级清单**里的 <file name="x.ocx"> + <comClass clsid=...>，
 *   **不需要**厂商 identity（实测无 identity 的清单 CreateActCtx 成功且能建对象）；
 *   而外挂的 <组件文件名>.manifest 单放不收（实测 CoCreateInstance = 0x80040154），
 *   它只是素材 —— 必须并进进程清单才生效（工具包用 .res 把清单嵌进 EXE，同理）。
 *   <file name> 必须是**相对清单所在目录**的路径（绝对路径实测 14001）。
 *   因此 manifest 模式：逐组件（优先复用现成 <组件文件名>.manifest 的 <file> 块，
 *   没有则按表里的 progid/clsid/相对路径合成），把全部块聚合成一份清单写到 exe 目录旁，
 *   再 CreateActCtx 激活。注：外部清单按 ASCII 处理（工具包与厂商清单都是 ASCII）。
 * ------------------------------------------------------------------ */
static wchar_t g_comComposedManifest[MAX_PATH * 2] = L"";   /* 合成清单，退出时删 */

static int comComponentCount(void) {
    const Vb6OcxRef* tbl = NULL;
    int ocxCount = vb6_OcxRefGetTable(&tbl);
    return g_comLibCount + (ocxCount > 0 ? ocxCount : 0);
}

static int comExeDir(wchar_t* out, int cch) {
    DWORD n = GetModuleFileNameW(NULL, out, (DWORD)cch);
    if (n == 0 || n >= (DWORD)cch) return 0;
    wchar_t* slash = wcsrchr(out, L'\\');
    if (!slash) return 0;
    slash[1] = 0;
    return 1;
}

/* 表里的 fileName 本来就相对 exe 目录；绝对路径只有在 exe 目录之下才压得成相对。 */
static int comRelName(const wchar_t* fileName, const wchar_t* exeDir,
                      wchar_t* out, int cch) {
    out[0] = 0;
    if (!fileName || !*fileName) return 0;
    int isAbs = (wcschr(fileName, L':') != NULL) || (fileName[0] == L'\\');
    if (!isAbs) {
        wcsncpy(out, fileName, (size_t)cch - 1);
        out[cch - 1] = 0;
        return 1;
    }
    size_t dl = wcslen(exeDir);
    /* exeDir 带尾反斜杠；只有不带尾分隔符时才要求 fileName 在这里再有一个分隔符。 */
    int needSep = (dl > 0 && exeDir[dl - 1] != L'\\' && exeDir[dl - 1] != L'/');
    if (_wcsnicmp(fileName, exeDir, dl) == 0 &&
        (!needSep || fileName[dl] == L'\\' || fileName[dl] == L'/')) {
        const wchar_t* r = fileName + dl;
        if (*r == L'\\' || *r == L'/') r++;
        wcsncpy(out, r, (size_t)cch - 1);
        out[cch - 1] = 0;
        return 1;
    }
    return 0;
}

static int comAppend(wchar_t* buf, int cch, const wchar_t* text, int* used) {
    size_t n = wcslen(text);
    if (*used + (int)n >= cch) return 0;
    memcpy(buf + *used, text, (n + 1) * sizeof(wchar_t));
    *used += (int)n;
    return 1;
}

/* 把 <file ...> 开标签里的 name="..." 改成 rel（SxS 只吃相对路径）。 */
static int comRewriteNameAttr(wchar_t* block, int cch, const wchar_t* rel) {
    wchar_t* tagEnd = wcschr(block, L'>');
    if (!tagEnd) return 0;
    for (wchar_t* p = block; p < tagEnd; p++) {
        if ((p[0] == L'n' || p[0] == L'N') && _wcsnicmp(p, L"name", 4) == 0) {
            wchar_t* q = p + 4;
            while (q < tagEnd && (*q == L' ' || *q == L'\t')) q++;
            if (q >= tagEnd || *q != L'=') continue;
            q++;
            while (q < tagEnd && (*q == L' ' || *q == L'\t')) q++;
            if (q >= tagEnd || *q != L'"') return 0;
            q++;
            wchar_t* e = wcschr(q, L'"');
            if (!e) return 0;
            size_t relLen = wcslen(rel);
            size_t tail = wcslen(e);
            if (q + relLen >= block + cch - 1) return 0;
            memmove(q + relLen, e, (tail + 1) * sizeof(wchar_t));
            memcpy(q, rel, relLen * sizeof(wchar_t));
            return 1;
        }
    }
    return 0;
}

static int comAppendComClass(wchar_t* buf, int cch, const wchar_t* clsid,
                             const wchar_t* progId, int* used) {
    wchar_t blk[2048];
    if (progId && *progId) {
        _snwprintf(blk, 2047,
                   L"<comClass clsid=\"%ls\" threadingModel=\"Apartment\" progid=\"%ls\"></comClass>",
                   clsid, progId);
    } else {
        _snwprintf(blk, 2047,
                   L"<comClass clsid=\"%ls\" threadingModel=\"Apartment\"></comClass>",
                   clsid);
    }
    blk[2047] = 0;
    return comAppend(buf, cch, blk, used);
}

/* 聚合组件清单 → outPath（exe 目录旁的 c3com.<pid>.<tick>.manifest）。
 * 返回 1 = 已写出；0 = 失败，problems 写清每个组件的具体缺项。 */
#define VB6_COM_COMPOSE_MAX (64 * 1024)
static int comComposeManifest(wchar_t* outPath, int pathCch,
                              wchar_t* problems, int probCch) {
    outPath[0] = 0;
    wchar_t exeDir[MAX_PATH];
    if (!comExeDir(exeDir, MAX_PATH)) {
        comMissingJoin(problems, probCch, L"无法确定 exe 目录");
        return 0;
    }
    wchar_t* buf = (wchar_t*)malloc(VB6_COM_COMPOSE_MAX * sizeof(wchar_t));
    if (!buf) { comMissingJoin(problems, probCch, L"内存不足"); return 0; }
    int used = 0;
    comAppend(buf, VB6_COM_COMPOSE_MAX,
              L"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
              L"<assembly xmlns=\"urn:schemas-microsoft-com:asm.v1\" manifestVersion=\"1.0\">\r\n",
              &used);

    const Vb6OcxRef* tbl = NULL;
    int ocxCount = vb6_OcxRefGetTable(&tbl);
    if (!tbl) ocxCount = 0;
    int total = g_comLibCount + ocxCount;
    const wchar_t** uniq = (const wchar_t**)calloc((size_t)total + 1, sizeof(wchar_t*));
    wchar_t item[1024];
    if (!uniq) { comMissingJoin(problems, probCch, L"内存不足"); free(buf); return 0; }
    int uniqCount = 0;
    int emitted = 0;
    for (int pass = 0; pass < 2; pass++) {
        int count = (pass == 0) ? g_comLibCount : ocxCount;
        for (int i = 0; i < count; i++) {
            const wchar_t* fileName = (pass == 0) ? g_comLibs[i].fileName : tbl[i].fileName;
            if (!fileName || !*fileName) {
                _snwprintf(item, 1023, L"组件表第 %d 条没有文件名，无法写 <file name>", i);
                item[1023] = 0;
                comMissingJoin(problems, probCch, item);
                continue;
            }
            int dup = 0;
            for (int k = 0; k < uniqCount; k++)
                if (_wcsicmp(uniq[k], fileName) == 0) { dup = 1; break; }
            if (!dup) uniq[uniqCount++] = fileName;
        }
    }

    /* 同一 DLL 声明多个 coclass ⇒ 只能有**一个** <file>，comClass 都挂在它下面
     * （实测重复 <file name> 块会让整份清单解析不出 ProgID 映射）。 */
    for (int u = 0; u < uniqCount; u++) {
        const wchar_t* fileName = uniq[u];
        wchar_t rel[MAX_PATH * 2], abs[MAX_PATH * 2];
        if (!comRelName(fileName, exeDir, rel, MAX_PATH * 2)) {
            _snwprintf(item, 1023, L"%ls: 组件不在 exe 目录下，写不成相对 <file name>", fileName);
            item[1023] = 0;
            comMissingJoin(problems, probCch, item);
            continue;
        }
        const wchar_t* p = comComponentPath(fileName, abs, MAX_PATH * 2);
        if (GetFileAttributesW(p) == INVALID_FILE_ATTRIBUTES) {
            _snwprintf(item, 1023, L"%ls: 组件文件不存在 (%ls)", rel, p);
            item[1023] = 0;
            comMissingJoin(problems, probCch, item);
            continue;
        }
        /* 优先：现成的 <组件文件名>.manifest（工具包 per-file 清单） */
        wchar_t ext[MAX_PATH * 2 + 16];
        _snwprintf(ext, MAX_PATH * 2 + 15, L"%ls.manifest", p);
        ext[MAX_PATH * 2 + 15] = 0;
        if (GetFileAttributesW(ext) != INVALID_FILE_ATTRIBUTES) {
            wchar_t* block = (wchar_t*)malloc(8192 * sizeof(wchar_t));
            if (block) {
                int got = vb6_ComManifestExtractFileBlock(ext, block, 8192);
                int appended = got && comRewriteNameAttr(block, 8192, rel) &&
                               comAppend(buf, VB6_COM_COMPOSE_MAX, block, &used);
                free(block);
                if (appended) { emitted++; continue; }
                /* 抽到但改写/追加失败 → 退回合成 */
            }
        }
        wchar_t head[2048];
        _snwprintf(head, 2047, L"<file name=\"%ls\" hashalg=\"SHA1\">", rel);
        head[2047] = 0;
        if (!comAppend(buf, VB6_COM_COMPOSE_MAX, head, &used)) {
            comMissingJoin(problems, probCch, L"聚合清单超出 64KB 缓冲上限");
            free(uniq); free(buf);
            return 0;
        }
        for (int pass = 0; pass < 2; pass++) {
            int count = (pass == 0) ? g_comLibCount : ocxCount;
            for (int i = 0; i < count; i++) {
                const wchar_t* fn     = (pass == 0) ? g_comLibs[i].fileName : tbl[i].fileName;
                const wchar_t* clsid  = (pass == 0) ? g_comLibs[i].clsidStr : tbl[i].clsidStr;
                const wchar_t* progId = (pass == 0) ? g_comLibs[i].progId   : tbl[i].progId;
                if (!fn || _wcsicmp(fn, fileName) != 0) continue;
                if (!clsid || !*clsid) {
                    _snwprintf(item, 1023, L"%ls: 缺 CLSID（类型库导入没给出），无法合成 <comClass>", rel);
                    item[1023] = 0;
                    comMissingJoin(problems, probCch, item);
                    continue;
                }
                if (!comAppendComClass(buf, VB6_COM_COMPOSE_MAX, clsid, progId, &used)) {
                    comMissingJoin(problems, probCch, L"聚合清单超出 64KB 缓冲上限");
                    free(uniq); free(buf);
                    return 0;
                }
            }
        }
        if (!comAppend(buf, VB6_COM_COMPOSE_MAX, L"</file>", &used)) {
            comMissingJoin(problems, probCch, L"聚合清单超出 64KB 缓冲上限");
            free(uniq); free(buf);
            return 0;
        }
        emitted++;
    }
    free(uniq);
    if (emitted == 0) {
        /* 一份只能解析出 0 个 <file> 的合成清单没有意义 —— 明确失败，不建空 ActCtx。 */
        comMissingJoin(problems, probCch, L"没有任何组件能写进清单");
        free(buf);
        return 0;
    }
    comAppend(buf, VB6_COM_COMPOSE_MAX, L"</assembly>\r\n", &used);

    wchar_t name[96];
    _snwprintf(name, 95, L"c3com.%lu.%I64u.manifest",
               (unsigned long)GetCurrentProcessId(), (unsigned __int64)GetTickCount64());
    name[95] = 0;
    _snwprintf(outPath, (size_t)pathCch - 1, L"%ls%ls", exeDir, name);
    outPath[pathCch - 1] = 0;

    HANDLE hf = CreateFileW(outPath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS,
                            FILE_ATTRIBUTE_NORMAL, NULL);
    if (hf == INVALID_HANDLE_VALUE) {
        _snwprintf(item, 1023, L"合成清单写不出去 (目录只读?): %ls", outPath);
        item[1023] = 0;
        comMissingJoin(problems, probCch, item);
        free(buf);
        return 0;
    }
    int ok = 0;
    char* utf8 = (char*)malloc(VB6_COM_COMPOSE_MAX);
    if (utf8) {
        int bytes = WideCharToMultiByte(CP_UTF8, 0, buf, used, utf8,
                                        VB6_COM_COMPOSE_MAX, NULL, NULL);
        if (bytes > 0) {
            DWORD wrote = 0;
            ok = WriteFile(hf, utf8, (DWORD)bytes, &wrote, NULL) && wrote == (DWORD)bytes;
        }
        free(utf8);
    }
    CloseHandle(hf);
    free(buf);
    if (!ok) {
        DeleteFileW(outPath);
        outPath[0] = 0;
        comMissingJoin(problems, probCch, L"合成清单写入失败");
        return 0;
    }
    wcsncpy(g_comComposedManifest, outPath, MAX_PATH * 2 - 1);
    g_comComposedManifest[MAX_PATH * 2 - 1] = 0;
    return 1;
}

/* manifest 模式的"明确失败"：打印**具体缺什么**并终止进程。
 * 为什么不继续跑：继续 = 静默回退直加载，正是本模式明令禁止的那条。
 * 退出码 1 由夹具/脚本断言；direct 模式不受影响。 */
static void comManifestFatal(const wchar_t* what) {
    // 诊断刻意用 ASCII: 控制台/重定向文件里中文随代码页变成乱码, 夹具要能 grep 到具体缺项
    fwprintf(stderr, L"[C3_COM] manifest mode FAILED -- missing: %ls\n", what ? what : L"(unknown)");
    fwprintf(stderr, L"[C3_COM] manifest mode does NOT fall back to direct; use -ComMode direct explicitly\n");
    ExitProcess(1);
}

static void vb6_ComActCtxEnter(void) {
    int mode = vb6_ComMode();
    if (mode == VB6_COM_MODE_DIRECT) {
        if (comLibTraceEnabled())
            fprintf(stderr, "[C3_COM] mode=direct: 直加载 (LoadLibrary+DllGetClassObject), 不建 ActCtx\n");
        return;
    }

    /* 1) 显式清单（-ComManifest / C3_COM_MANIFEST）优先；2) exe 旁 <exe>.manifest；
     * 3) 都没有且声明了组件 ⇒ 按工具包口径合成聚合清单。 */
    wchar_t manifest[MAX_PATH * 2];
    int hasManifest = 0;
    if (g_comManifestPath[0]) {
        wcsncpy(manifest, g_comManifestPath, MAX_PATH * 2 - 1);
        manifest[MAX_PATH * 2 - 1] = 0;
        hasManifest = 1;
    } else if (GetEnvironmentVariableW(L"C3_COM_MANIFEST", manifest, MAX_PATH * 2) > 0) {
        hasManifest = 1;
    }
    if (!hasManifest) {
        wchar_t guessed[MAX_PATH * 2];
        if (comManifestPathResolve(guessed, MAX_PATH * 2) &&
            GetFileAttributesW(guessed) != INVALID_FILE_ATTRIBUTES) {
            wcsncpy(manifest, guessed, MAX_PATH * 2 - 1);
            manifest[MAX_PATH * 2 - 1] = 0;
            hasManifest = 1;
        }
    }

    wchar_t missing[2048];
    missing[0] = 0;
    if (hasManifest) {
        if (!vb6_ComManifestValidateClient(manifest, missing, 2048))
            goto unusable;
    } else {
        if (comComponentCount() <= 0) {
            if (mode == VB6_COM_MODE_MANIFEST)
                fprintf(stderr, "[C3_COM] manifest mode: no ComLib/OCX components declared; nothing to activate\n");
            return;   /* 没有组件可覆盖：不建空 ActCtx，也不改直加载行为 */
        }
        if (mode != VB6_COM_MODE_MANIFEST) {
            /* auto 且没人显式给清单：保持既有行为（走直加载），不主动合成 ——
             * 默认模式对存量产物必须零行为变化，合成留给显式的 -ComMode manifest。 */
            if (comLibTraceEnabled())
                fprintf(stderr, "[C3_COM] auto: no explicit manifest; direct path only\n");
            return;
        }
        if (!comComposeManifest(manifest, MAX_PATH * 2, missing, 2048))
            goto unusable;
    }

    wchar_t err[512];
    if (!vb6_ComActCtxCreate(manifest, err, 512)) {
        if (mode == VB6_COM_MODE_MANIFEST) comManifestFatal(err);
        fwprintf(stderr, L"[C3_COM] auto: fallback to direct -- %ls\n", err);
        return;
    }
    if (comLibTraceEnabled() || hasManifest)
        fwprintf(stderr, L"[C3_COM] using manifest (ActCtx activated): %ls\n", manifest);
    return;

unusable:
    if (mode == VB6_COM_MODE_MANIFEST) comManifestFatal(missing);
    fwprintf(stderr, L"[C3_COM] auto: fallback to direct -- manifest unusable: %ls\n", missing);
}

static void vb6_ComActCtxLeave(void) {
    vb6_ComActCtxRelease();
    if (g_comComposedManifest[0]) {
        /* C3_COM_KEEP_MANIFEST=1 保留合成清单，便于人工核对（默认退出即删）。 */
        if (GetEnvironmentVariableW(L"C3_COM_KEEP_MANIFEST", NULL, 0) == 0)
            DeleteFileW(g_comComposedManifest);
        g_comComposedManifest[0] = 0;
    }
}

/* 表查找:
 *   1) 精确 progId 匹配 (大小写不敏感)
 *   2) progId 无点号 (未限定名, 如 "FileSystemObject") → coclassName 匹配
 *      覆盖 `Dim x As New ClassName` 传的是类名而非完整 ProgID 的情形.
 *      只用在无点号时: 有 "Lib.Class" 形式的调用不做末段匹配, 避免
 *      例如 CreateObject("Word.Application") 误命中名为 "Application" 的无关 coclass.
 */
static const Vb6ComLib* vb6_ComLibLookup(const wchar_t* progId) {
    if (!progId || !*progId || g_comLibCount <= 0) return NULL;
    for (int i = 0; i < g_comLibCount; i++) {
        if (g_comLibs[i].progId && _wcsicmp(g_comLibs[i].progId, progId) == 0)
            return &g_comLibs[i];
    }
    if (!wcschr(progId, L'.')) {   /* 无点号 */
        for (int i = 0; i < g_comLibCount; i++) {
            if (g_comLibs[i].coclassName &&
                _wcsicmp(g_comLibs[i].coclassName, progId) == 0)
                return &g_comLibs[i];
        }
    }
    return NULL;
}

/* 免注册激活: LoadLibrary → GetProcAddress("DllGetClassObject") →
 * IClassFactory::CreateInstance. 流程与 OCX 控件 (vb6forms_axsite.c
 * ocxCreateFromPath) 一致, 并把 Fix 159 的失败原因区分一并搬来:
 *   HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT)  PE 位宽不匹配 — 始终打日志
 *   HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND)   DLL 未导出 DllGetClassObject
 *   CO_E_CLASSNOTREGISTERED (0x800401F1)      路径缺失/无效等通用失败
 * 三者都是 FAILED 码; 调用方只做 SUCCEEDED/FAILED 判定, 区分只体现在返回值与 stderr. */
static HRESULT comLibCreateFromPath(const wchar_t* dllPath, REFCLSID rclsid, void** ppUnk) {
    *ppUnk = NULL;
    if (!dllPath || !*dllPath) return ((HRESULT)0x800401F1L);
    HMODULE hMod = LoadLibraryW(dllPath);
    if (!hMod) {
        DWORD le = GetLastError();
        if (le == ERROR_BAD_EXE_FORMAT) {
            /* 宿主与组件位宽不一致: 用 --arch x86 重编宿主, 或换同位宽组件 */
            fprintf(stderr, "[C3_COM]   %ls: PE 位宽不匹配 (ERROR_BAD_EXE_FORMAT); "
                            "宿主与组件位数必须一致 — 用 --arch x86 重编, 或换同位宽组件\n",
                    dllPath);
            return HRESULT_FROM_WIN32(ERROR_BAD_EXE_FORMAT);
        }
        if (comLibTraceEnabled())
            fprintf(stderr, "[C3_COM]   LoadLibrary(%ls) failed: %lu\n",
                    dllPath, (unsigned long)le);
        return ((HRESULT)0x800401F1L);
    }
    typedef HRESULT (__stdcall *PFN_DllGetClassObject)(REFCLSID, REFIID, void**);
    PFN_DllGetClassObject pGet = (PFN_DllGetClassObject)GetProcAddress(hMod, "DllGetClassObject");
    if (!pGet) {
        FreeLibrary(hMod);
        if (comLibTraceEnabled())
            fprintf(stderr, "[C3_COM]   %ls: no DllGetClassObject export\n", dllPath);
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }
    IClassFactory* cf = NULL;
    HRESULT hr = pGet(rclsid, &IID_IClassFactory, (void**)&cf);
    if (FAILED(hr) || !cf) {
        /* FreeLibrary 不做: DllGetClassObject 成功后对象可能回引 DLL */
        if (comLibTraceEnabled())
            fprintf(stderr, "[C3_COM]   DllGetClassObject failed: 0x%08lX\n",
                    (unsigned long)hr);
        return hr;
    }
    hr = cf->lpVtbl->CreateInstance(cf, NULL, &IID_IUnknown, ppUnk);
    cf->lpVtbl->Release(cf);
    return hr;
}

/* 表命中后的本地激活总入口.
 * 路径候选与 OCX 的 vb6_ocxCreateAny 一致:
 *   1) <exe目录>\<fileName 相对部分>   便携分发, 不依赖 CWD
 *   2) <exe目录>\<裸文件名>            同目录
 *   3) 原样                            开发机/绝对路径
 *   4) 裸文件名                        系统 DLL 搜索路径
 * 绝对路径只在候选 3 出现一次, 不重复拼 exe 目录.
 * 返回 IDispatch*; 失败 NULL — 调用方继续注册表路径。
 * pFailHr (账 #305 §B186): 交回**最后那一趟**的 HRESULT —— 调用方要分得出「DLL 里有人 raise 了一条
 * 自己的 VB 错误」(0x800A0000|n, 不许再咽进注册表回退) 与「这台机器上根本没这枚 DLL」。 */
static void* comLibCreateLocal(const Vb6ComLib* e, HRESULT* pFailHr) {
    if (pFailHr) *pFailHr = E_FAIL;
    if (!e || !e->clsidStr || !*e->clsidStr) return NULL;
    CLSID clsid;
    if (FAILED(CLSIDFromString((LPOLESTR)e->clsidStr, &clsid))) {
        if (comLibTraceEnabled())
            fwprintf(stderr, L"[C3_COM]   bad CLSID in table: %ls\n", e->clsidStr);
        return NULL;
    }

    wchar_t exeDir[MAX_PATH];
    exeDir[0] = 0;
    DWORD n = GetModuleFileNameW(NULL, exeDir, MAX_PATH);
    if (n > 0 && n < MAX_PATH) {
        wchar_t* es = wcsrchr(exeDir, L'\\');
        if (es) es[1] = 0;       /* 保留结尾 '\' */
        else exeDir[0] = 0;
    }

    const wchar_t* fname = e->fileName ? e->fileName : L"";
    if (fname && *fname) {
        const wchar_t* s = wcsrchr(fname, L'\\');
        if (!s) s = wcsrchr(fname, L'/');
        if (s) fname = s + 1;
    }

    int isAbs = (e->fileName && wcschr(e->fileName, L':') != NULL)
                || (e->fileName && e->fileName[0] == L'\\');

    wchar_t exeRel[MAX_PATH * 2];
    exeRel[0] = 0;
    if (exeDir[0] && e->fileName && *e->fileName && !isAbs &&
        wcslen(exeDir) + wcslen(e->fileName) + 1 < MAX_PATH * 2) {
        wcscpy(exeRel, exeDir);
        wcscat(exeRel, e->fileName);
    }
    wchar_t exeSide[MAX_PATH * 2];
    exeSide[0] = 0;
    if (exeDir[0] && fname[0] && wcslen(exeDir) + wcslen(fname) + 1 < MAX_PATH * 2) {
        wcscpy(exeSide, exeDir);
        wcscat(exeSide, fname);
    }

    IUnknown* pUnk = NULL;
    HRESULT hr = E_FAIL;
    const wchar_t* cands[4];
    int nc = 0;
    if (exeRel[0]) cands[nc++] = exeRel;
    if (exeSide[0]) cands[nc++] = exeSide;
    if (e->fileName && *e->fileName) cands[nc++] = e->fileName;
    if (fname[0] && fname != e->fileName) cands[nc++] = fname;   /* 无目录部分时 fname==fileName, 不重复试 */
    for (int ci = 0; ci < nc && !pUnk; ci++) {
        hr = comLibCreateFromPath(cands[ci], &clsid, (void**)&pUnk);
        if (SUCCEEDED(hr) && pUnk) break;
    }
    if (!pUnk) {
        if (pFailHr) *pFailHr = hr;
        if (comLibTraceEnabled())
            fwprintf(stderr, L"[C3_COM]   local activation failed: %ls (0x%08lX)\n",
                     e->fileName ? e->fileName : L"(no path)", (unsigned long)hr);
        return NULL;
    }
    IDispatch* pDisp = NULL;
    hr = pUnk->lpVtbl->QueryInterface(pUnk, &IID_IDispatch, (void**)&pDisp);
    pUnk->lpVtbl->Release(pUnk);
    if (FAILED(hr) || !pDisp) {
        if (pFailHr) *pFailHr = hr;
        return NULL;
    }
    /* Fix 160-com: 登记 对象→默认调度接口 ITypeInfo, 供 vb6_getDispid /
     * vb6_ComInvoke 免注册解析 DISPID 与分派 (TLB 未注册时 Chilkat 等仍可用). */
    if (e->dispIface) comLibRegisterObjTI((void*)pDisp, e->dispIface);
    return (void*)pDisp;
}

// ============================================================
// CreateObject / GetObject
// ============================================================

/* 账 #305 §B186: create 那一趟失败该交回哪一条 VB 错误, 从此只住这一枚函数 ——
 * 三条腿 (ProgID 解析失败 / 免注册本地激活失败 / 注册表 CoCreateInstance 失败) 都问它,
 * 不再各抄一遍「一律 429」。
 * 服务端里没人接住的运行期错误现在由类工厂交回 0x800A0000|错误号 (§B185 同一决定的第二站),
 * 那一档折回它自己的号 —— 与「同一个类在客户进程里直接 New、调用方 On Error 拿到的就是那一号」
 * 对齐；认不出那一档 (0x8004xxxx 没注册 / 0x8007xxxx 起不动进程) 才照旧 429。
 * 描述三级: IErrorInfo (既有那条通道) -> VB 标准错误号表 (只有折回来的那一号才查) ->
 * 那句 ASCII 的 can't create object。第二级不许跳: 真·内存不够以外的失败交的是服务端自己的错号,
 * 拿 429 那句描述去配它会指错方向。 */
static void vb6_ComFailCreate(HRESULT hr, const wchar_t* progId) {
    int32_t vbErr = vb6_ComVbErrFromHr(hr);
    BSTR desc = NULL;
    IErrorInfo* perr = NULL;
    if (SUCCEEDED(GetErrorInfo(0, &perr)) && perr) {
        perr->lpVtbl->GetDescription(perr, &desc);
        perr->lpVtbl->Release(perr);
    }
    if (!desc && vbErr) {
        const wchar_t* std = vb6_StdErrorDesc(vbErr);
        if (std) desc = SysAllocString(std);
    }
    if (!desc) {
        wchar_t buf[512];
        swprintf(buf, 512, L"ActiveX component can't create object (ProgID: %ls, HRESULT: 0x%08lX)",
                 progId ? progId : L"?", (unsigned long)hr);
        desc = SysAllocString(buf);
    }
    vb6_RaiseError(vbErr ? vbErr : 429, desc);
}

void* vb6_CreateObject(const wchar_t* progId) {
    if (!progId) return NULL;

    // Fix 160: ComLib= 免注册 — 本地 DLL 优先 (LoadLibrary+DllGetClassObject, 绕开注册表).
    // 表未声明 (ComLib= 为空) 时 vb6_ComLibLookup 恒返 NULL, 该块整体跳过,
    // 后续注册表路径与改动前完全一致 — 未声明组件零变化.
    // C3_COM_PREFER_REG=1 强制注册表优先 (对齐 C3_OCX_PREFER_REG).
    {
        static int preferReg = -1;
        if (preferReg < 0)
            preferReg = (GetEnvironmentVariableW(L"C3_COM_PREFER_REG", NULL, 0) > 0) ? 1 : 0;
        if (preferReg == 0) {
            const Vb6ComLib* hit = vb6_ComLibLookup(progId);
            if (hit) {
                HRESULT localHr = S_OK;
                void* p = comLibCreateLocal(hit, &localHr);
                if (p) return p;
                /* 账 #305 §B186: 本地那一趟要是交回服务端自己 raise 的错误, 就地折给调用方 ——
                 * 再回退注册表会把「DLL 里抛了 7」读成「这台机器上没这个组件」(429)。
                 * 真·找不到 DLL/起不动的那些照旧往下回退, 便携分发那条路一行没动。 */
                if (vb6_ComVbErrFromHr(localHr)) {
                    vb6_ComFailCreate(localHr, progId);
                    return NULL;
                }
                if (comLibTraceEnabled())
                    fwprintf(stderr, L"[C3_COM]   %ls: local activation failed, trying registry\n",
                             progId);
            }
        }
    }

    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(progId, &clsid);
    if (FAILED(hr)) {
        // VB6 Error 429: ActiveX component can't create object
        vb6_ComFailCreate(hr, progId);
        return NULL;
    }

    IDispatch* pDisp = NULL;
    hr = CoCreateInstance(&clsid, NULL, CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER,
                          &IID_IDispatch, (void**)&pDisp);
    if (FAILED(hr)) {
        vb6_ComFailCreate(hr, progId);
        return NULL;
    }

    return (void*)pDisp;
}

void* vb6_ComGlobalInstance(const wchar_t* progId) {
    /* typelib appobject 单例缓存: 线性表足够 (每个引用库至多几枚 appobject)。
     * 实例持有到进程结束 (VB6 语义: 全局对象与工程同生命周期, 不显式释放)。 */
    static struct { const wchar_t* progId; void* disp; } cache[32];
    static int count = 0;
    if (progId) {
        for (int i = 0; i < count; i++) {
            if (wcscmp(cache[i].progId, progId) == 0) return cache[i].disp;
        }
    }
    void* disp = vb6_CreateObject(progId);
    if (disp && count < 32) {
        cache[count].progId = progId;
        cache[count].disp = disp;
        count++;
    }
    return disp;
}

void* vb6_GetObject(const wchar_t* pathName, const wchar_t* progId) {
    if (!progId) return NULL;

    CLSID clsid;
    // Fix 160: ComLib= 免注册 — GetObject 只做 GetActiveObject / IPersistFile::Load,
    // 不做本地实例化, 所以只借用表把 ProgID 解析成 CLSID (免注册可得), 未命中则原样走注册表.
    HRESULT hr = E_FAIL;
    {
        const Vb6ComLib* hit = vb6_ComLibLookup(progId);
        if (hit && hit->clsidStr && *hit->clsidStr)
            hr = CLSIDFromString((LPOLESTR)hit->clsidStr, &clsid);
    }
    if (FAILED(hr)) hr = CLSIDFromProgID(progId, &clsid);
    if (FAILED(hr)) {
        fwprintf(stderr, L"vb6_GetObject: CLSIDFromProgID(\"%ls\") failed: 0x%08lX\n",
                 progId, (unsigned long)hr);
        return NULL;
    }

    IUnknown* pUnk = NULL;

    if (pathName && wcslen(pathName) > 0) {
        // 从文件加载: GetObject(pathName, progId)
        IPersistFile* pFile = NULL;
        hr = CoCreateInstance(&clsid, NULL, CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER,
                              &IID_IPersistFile, (void**)&pFile);
        if (SUCCEEDED(hr) && pFile) {
            hr = pFile->lpVtbl->Load(pFile, pathName, STGM_READ);
            if (SUCCEEDED(hr)) {
                hr = pFile->lpVtbl->QueryInterface(pFile, &IID_IDispatch, (void**)&pUnk);
            }
            pFile->lpVtbl->Release(pFile);
        }
        // 回退: 尝试GetActiveObject
        if (!pUnk) {
            hr = GetActiveObject(&clsid, NULL, &pUnk);
        }
    } else {
        // 获取运行中的对象: GetObject(, progId)
        hr = GetActiveObject(&clsid, NULL, &pUnk);
    }

    if (FAILED(hr) || !pUnk) {
        fwprintf(stderr, L"vb6_GetObject: failed for \"%ls\": 0x%08lX\n",
                 progId, (unsigned long)hr);
        return NULL;
    }

    IDispatch* pDisp = NULL;
    hr = pUnk->lpVtbl->QueryInterface(pUnk, &IID_IDispatch, (void**)&pDisp);
    pUnk->lpVtbl->Release(pUnk);

    if (FAILED(hr)) {
        fwprintf(stderr, L"vb6_GetObject: QueryInterface IDispatch failed: 0x%08lX\n",
                 (unsigned long)hr);
        return NULL;
    }

    return (void*)pDisp;
}

// ============================================================
// IsNothing / ReleaseObject
// ============================================================

int32_t vb6_IsNothing(void* obj) {
    return (obj == NULL) ? -1 : 0;  // VB6 True=-1, False=0
}

void vb6_ReleaseObject(void** objPtr) {
    if (!objPtr || !*objPtr) return;

    /* Fix 191: 只有真是 COM 对象才谈得上 Release。VB6 里 `Set .OriginalIOleIPAO = Me`
     * 存的是接口的 COM 身份, 而生成码把类实例 `me` 直接塞进了接口字段, 于是
     * `Set ... = Nothing` 走到这里就是对 `me` 解引用 lpVtbl (VBFlexGridDemo 左键
     * 点击 → DeActivateIPAO → 读 0x10 崩溃)。非 COM 接收者按 VB6 的
     * "对象不支持此属性或方法" 处理: 只清指针, 不越权调用。 */
    IUnknown* pUnk = (IUnknown*)(*objPtr);
    if (vb6_ComIsDispatchable(pUnk)) {
        pUnk->lpVtbl->Release(pUnk);
    }
    *objPtr = NULL;
}

/* Fix <vbeclipse>: vb6_ReleaseObject 的反向操作 —— 与它**成对**维护, 判据必须是同一个。
 *
 * 起因: vb6_VariantClear() 对 vb6_vtDispatch 调 vb6_ReleaseObject (认为 Variant 持有引用),
 * 但 Variant 的构造侧 (vb6rtl_variant.h 的 vb6_VariantObject) 原先**不 AddRef** ——
 * 构造不取引用、析构却释放, 于是任何「对象既被原主 `Set x = Nothing` 释放、又被放进
 * Variant 容器 (List/Collection/数组)」的路径都是**过度释放**。实测 ucPerspective.AddView:
 *     l_View = vb6_ComObject_FromInstance(...);          // 1 个引用
 *     vb6_VARIANT t = vb6_VariantFromValue((void*)l_View); // 原先不 AddRef
 *     vb6_List_Add(m_Views, ViewId, &t, -1, 0);            // 容器存了裸引用
 *     vb6_ReleaseObject(&l_View);                          // refcount→0, 对象销毁
 * → 容器里留下指向已释放对象的 Variant, 之后取出当 IDispatch 用 / 容器清理时再 Release
 *   一次 → **0xC0000374 堆损坏**(检测点随堆布局漂移, 总在后面某次分配/释放才爆)。
 * VB6 对照: `coll.Add key, obj` 会 AddRef, 所以随后的 `Set obj = Nothing` 不销毁对象。
 *
 * 判据与 vb6_ReleaseObject 完全一致: 只对**真 COM 接收者** AddRef。裸结构体/UDT 地址
 * (vb6_cls_* 直传、Set .OriginalIOleIPAO = Me 那类) 一律跳过 —— 否则解引用 lpVtbl 即 AV。
 * 定义放这里而非 header: vb6_VariantObject 是 static inline, 需要一个外部符号,
 * 前向声明在 vb6rtl_variant.h (避免头文件反向依赖 vb6com.h)。 */
void vb6_ComAddRefDispatch(void* p) {
    if (!p) return;
    if (!vb6_ComIsDispatchable(p)) return;
    IUnknown* pUnk = (IUnknown*)p;
    pUnk->lpVtbl->AddRef(pUnk);
}
