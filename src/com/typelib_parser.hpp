#pragma once
// VB6 TypeLib解析器 - P6.3 前期绑定支持
// 编译期使用Windows LoadTypeLib/ITypeInfo API提取类型库信息
// 注册COM coclass/接口/方法签名到符号表

#include "common/types.hpp"
#include "common/diagnostics.hpp"
#include <algorithm>
#include <string>
#include <vector>
#include <unordered_map>
#include <memory>

// Windows COM类型前置声明 (编译期仅Windows可用)
typedef struct tagTYPEATTR TYPEATTR;
typedef struct tagFUNCDESC FUNCDESC;
typedef struct tagVARDESC VARDESC;
typedef struct tagELEMDESC ELEMDESC;
typedef struct tagTYPEDESC TYPEDESC;
typedef struct tagPARAMDESC PARAMDESC;

namespace vb6c3 {

// ============================================================
// COM方法/属性描述 (从ITypeInfo提取)
// ============================================================

enum class ComMemberKind : uint8_t {
    Method,         // 普通方法 (INVOKE_FUNC)
    PropertyGet,    // 属性读取 (INVOKE_PROPERTYGET)
    PropertyPut,    // 属性赋值 (INVOKE_PROPERTYPUT)
    PropertyPutRef, // 属性引用赋值 (INVOKE_PROPERTYPUTREF)
};

enum class ComParamDir : uint8_t {
    In,     // [in] 参数
    Out,    // [out] 参数
    InOut,  // [in,out] 参数
    RetVal, // [out, retval] 返回值参数
};

struct ComParamInfo {
    std::string name;           // 参数名
    Vb6Type type = Vb6Type::Variant;    // VB6类型 (带初值: 这枚字段会被发码期当"这成员是什么型"的答案读)
    ComParamDir direction = ComParamDir::In;  // 参数方向
    bool isOptional = false;   // Optional参数
    bool hasDefault = false;   // 有默认值
};

struct ComMemberInfo {
    std::string name;           // 方法/属性名 (小写, 用于查找)
    std::string realName;       // 原始名称 (保留大小写)
    ComMemberKind kind = ComMemberKind::Method;      // 成员类别
    // 账 #245: returnType 会被 driver_semantics 复制进 Symbol::ComMethodSig, 再被发码期
    // 当「这枚 COM 属性是什么型」的唯一答案读。曾经它没有初值, 而 parseVarDesc 只在
    // VAR_PERINSTANCE 那支赋值 ⇒ VAR_PROPERTY (dual 接口的属性 VARDESC, 如 stdole 的
    // StdFont.Name) 交出去的是没赋过值的 16 位字段: 同一枚二进制连跑三次读到
    // 29620/6971/50156, 换台工具链编出的二进制就落进另一个档 ⇒ 同一笔提交的发码分家
    // (泛型 ComGetProp ⇄ vb6_ComGetIntProp, 而 Name 是 BSTR)。初值 = Variant = "不知道"。
    Vb6Type returnType = Vb6Type::Variant;  // 返回类型 (Method/PropertyGet)
    std::vector<ComParamInfo> params;  // 参数列表
    int32_t memid = 0;          // DISPID (成员ID)
    int32_t vtableIndex = -1;   // vtable偏移 (前期绑定用)
    CallConv callConv = CallConv::StdCall;  // 调用约定
};

// ============================================================
// COM接口描述
// ============================================================

struct ComInterfaceInfo {
    std::string name;           // 接口名 (如 "IFileSystem3")
    std::string iidStr;         // IID字符串 (如 "{2A0A3E20-...}")
    bool isDual = false;        // 双重接口 (dispinterface + vtable)
    bool isDispatch = false;    // 纯IDispatch接口
    std::vector<ComMemberInfo> members;  // 方法/属性列表

    // P24-10: 默认成员 (DISPID_VALUE=0), 如 Collection.Item / Dictionary.Item
    // VB6语义: obj(args) 等价于 obj.DefaultMember(args)
    std::string defaultMemberName;       // 小写, 用于查找
    std::string defaultMemberRealName;   // 原始大小写, 用于代码生成

    // 按名称查找成员 (小写)
    const ComMemberInfo* findMember(const std::string& lowerName) const {
        for (auto& m : members) {
            if (m.name == lowerName) return &m;
        }
        return nullptr;
    }

};

// ============================================================
// COM coclass描述
// ============================================================

struct ComCoClassInfo {
    std::string name;           // coclass名 (如 "FileSystemObject")
    std::string progId;         // ProgID (如 "Scripting.FileSystemObject")
    std::string clsidStr;       // CLSID字符串
    std::string defaultIfaceName;  // 默认接口名
    const ComInterfaceInfo* defaultIface = nullptr;  // 默认接口指针 (解析后填充)
    // P13.20: Event source interface (IMPLTYPEFLAG_FSOURCE)
    std::string defaultSourceIfaceName;   // 默认事件源接口名
    const ComInterfaceInfo* defaultSourceIface = nullptr;  // 默认事件源接口指针
    std::vector<std::string> sourceIfaceNames;  // 所有事件源接口名列表
    // P24-04: VB_GlobalNameSpace = True 的coclass, 其默认接口的Public方法提升为全局符号
    // 检测启发式: TYPEFLAG_FPREDECLID + 默认接口有与TypeLib项目名同名的方法, 或名字含"Global"
    bool isGlobalNamespace = false;
};

// ============================================================
// COM模块描述 (P24-04: TKIND_MODULE → ActiveX DLL全局函数)
// VB6 ActiveX DLL中的标准模块(.bas) Public函数在TypeLib中记录为TKIND_MODULE
// 调用方式: VBMAN.Version() — 工程名为命名空间, 非COM对象创建
// ============================================================

struct ComModuleInfo {
    std::string name;           // 模块名 (= VB6工程名, 如 "VBMAN")
    std::string dllPath;        // 源DLL路径 (从TypeLib文件路径推断)
    std::vector<ComMemberInfo> functions;  // 模块级全局函数
    std::vector<ComMemberInfo> constants;  // 模块级常量

    // 按名称查找函数 (小写)
    const ComMemberInfo* findFunction(const std::string& lowerName) const {
        for (auto& f : functions) {
            if (f.name == lowerName) return &f;
        }
        return nullptr;
    }
};

// ============================================================
// Fix 018: COM枚举类型描述 (TKIND_ENUM → 枚举成员)
// VB6 引用 COM 类型库后, 枚举成员作为全局可见的命名常量
// (如 Scripting.Runtime 的 TextCompare, ADO 的 adStateClosed)
// ============================================================

struct ComEnumMemberInfo {
    std::string name;           // 成员名 (原始大小写, 如 "TextCompare")
    std::string lowerName;      // 小写 (查找用)
    int64_t value = 0;          // 枚举值 (来自 VARDESC.lpvarValue)
};

struct ComEnumInfo {
    std::string name;           // 枚举类型名 (如 "CompareMethod")
    std::string lowerName;      // 小写
    std::vector<ComEnumMemberInfo> members;
};

// ============================================================
// TypeLib解析结果
// ============================================================

struct TypeLibResult {
    std::string name;           // 类型库名 (如 "Microsoft Scripting Runtime")
    std::string version;        // 版本 (如 "1.0")
    std::string tlbPath;        // 类型库文件路径
    std::string canonPath;      // P24-05: 规范化路径(小写长路径, 用于缓存去重)
    std::string typeLibProjectName;  // P24-04: TypeLib项目名 (VB6工程Name=, 如"VBMANLIB")

    std::vector<std::unique_ptr<ComInterfaceInfo>> interfaces;
    std::vector<std::unique_ptr<ComCoClassInfo>> coclasses;
    std::vector<std::unique_ptr<ComModuleInfo>> modules;  // P24-04: TKIND_MODULE
    std::vector<std::unique_ptr<ComEnumInfo>> enums;      // Fix 018: TKIND_ENUM

    // 按名称查找coclass (不区分大小写)
    ComCoClassInfo* findCoClass(const std::string& name) const {
        std::string lower = name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        for (auto& cc : coclasses) {
            std::string ccLower = cc->name;
            std::transform(ccLower.begin(), ccLower.end(), ccLower.begin(), ::tolower);
            if (ccLower == lower) return cc.get();
            // 也匹配ProgID
            std::string progLower = cc->progId;
            std::transform(progLower.begin(), progLower.end(), progLower.begin(), ::tolower);
            if (progLower == lower) return cc.get();
        }
        return nullptr;
    }

    // 按名称查找接口
    ComInterfaceInfo* findInterface(const std::string& name) const {
        std::string lower = name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        for (auto& iface : interfaces) {
            std::string iLower = iface->name;
            std::transform(iLower.begin(), iLower.end(), iLower.begin(), ::tolower);
            if (iLower == lower) return iface.get();
        }
        return nullptr;
    }

    // P24-04: 按名称查找模块 (不区分大小写)
    ComModuleInfo* findModule(const std::string& name) const {
        std::string lower = name;
        std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
        for (auto& mod : modules) {
            std::string mLower = mod->name;
            std::transform(mLower.begin(), mLower.end(), mLower.begin(), ::tolower);
            if (mLower == lower) return mod.get();
        }
        return nullptr;
    }
};

// ============================================================
// TypeLib解析器
// ============================================================

class TypeLibParser {
public:
    TypeLibParser(Diagnostics& diag);
    ~TypeLibParser();

    // 按ProgID从注册表加载类型库
    // progId: "Scripting.FileSystemObject" 或 "Scripting"
    // 返回解析结果, 失败返回nullptr
    // silent=true: 不报告"TypeLib not found for ProgID"警告 (用于auto-load常用组件,
    //               用户未显式引用时缺失是预期行为, 不应污染c3-error.log)
    std::unique_ptr<TypeLibResult> loadByProgId(const std::string& progId, bool silent = false);

    // 按类型库文件路径加载
    std::unique_ptr<TypeLibResult> loadByPath(const std::string& tlbPath);

    // 按CLSID从注册表加载
    std::unique_ptr<TypeLibResult> loadByClsid(const std::string& clsidStr);

    // 按TypeLib名称+版本从注册表加载
    // name: "Microsoft Scripting Runtime", version: "1.0"
    std::unique_ptr<TypeLibResult> loadByName(const std::string& name,
                                               const std::string& version = "1.0");

    // 获取已解析的TypeLib缓存
    const std::vector<std::unique_ptr<TypeLibResult>>& cachedResults() const {
        return cache_;
    }

    // 查找已缓存中某个coclass (跨所有已加载TypeLib)
    ComCoClassInfo* findCachedCoClass(const std::string& name) const;

private:
    Diagnostics& diag_;

    // TypeLib缓存 (避免重复加载)
    std::vector<std::unique_ptr<TypeLibResult>> cache_;

    // ---- 内部实现 (Windows API) ----

    // 从ITypeLib解析所有类型
    bool parseTypeLib(void* pTypeLib, TypeLibResult& result);

    // 从ITypeInfo解析接口
    std::unique_ptr<ComInterfaceInfo> parseInterface(void* pTypeInfo,
                                                     const std::string& name);

    // 从ITypeInfo解析coclass
    std::unique_ptr<ComCoClassInfo> parseCoClass(void* pTypeInfo,
                                                  const std::string& name);

    // P24-04: 从ITypeInfo解析模块 (TKIND_MODULE)
    std::unique_ptr<ComModuleInfo> parseModule(void* pTypeInfo,
                                                 const std::string& name,
                                                 const std::string& dllPath);

    // Fix 018: 从ITypeInfo解析枚举 (TKIND_ENUM)
    std::unique_ptr<ComEnumInfo> parseEnum(void* pTypeInfo,
                                            const std::string& name);

    // 从FUNCDESC解析方法/属性
    ComMemberInfo parseFuncDesc(void* pTypeInfo, void* pFuncDesc, int index);

    // 从VARDESC解析属性
    ComMemberInfo parseVarDesc(void* pTypeInfo, void* pVarDesc);

    // TYPEDESC → Vb6Type 映射
    Vb6Type mapTypeDesc(void* pTypeDesc, void* pTypeInfo);

    // ELEMDESC → ComParamInfo
    ComParamInfo mapElemDesc(void* pElemDesc, const std::string& name, void* pTypeInfo);

    // IID → 字符串
    static std::string iidToString(const uint8_t* iidBytes);

    // ProgID → CLSID → TypeLib路径 (从注册表查找)
    std::string findTypeLibPathForProgId(const std::string& progId);

    // Fix 098: 当前正在解析的 TypeLib 库名 (= VB6 工程 Name=, 如 "VBMANLIB").
    // parseCoClass 里 ProgIDFromCLSID 反查失败时用它拼 "<库名>.<coclass名>".
    // 由 parseTypeLib 在枚举类型前设置.
    std::string currentProjectName_;
};

} // namespace vb6c3