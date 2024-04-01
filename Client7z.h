/*
从7z中抽象出的压缩、解压缩独立库
此库依赖7z.dll执行压缩、解压缩操作，需要在参数中指示7z.dll的路径
发布日期：2024-04-01
作者：hed10ne
*/
#pragma once
#include <windows.h>
#ifndef  _SIMPLE_7Z_LIB
#define _SIMPLE_7Z_LIB
extern "C" {
  /// @brief 执行压缩文件夹任务
  /// @param wszArchiveName         目标压缩文件路径及名称（全路径）
  /// @param wszPassword            压缩密码，当第一字节（wszPassword[0]）为\0时表示不加密
  /// @param wszDirectory           待压缩文件夹路径，请自行确保路径尾部存在路径分隔符
  /// @param wsz7zDllPath           7z.dll路径
  /// @param funcSetTotal           回调函数指针，此回调函数在压缩之前调用，指示压缩文件总大小。
  /// @param funcSetCompleted       回调函数指针，此回调函数在压缩过程中调用，指示压缩进度。
  /// @param bRecursive             是否递归压缩子文件夹
  /// @param bSolid                 是否固实压缩，后续参数为7z压缩属性
  /// @param nCompressionLevel      压缩级别，范围0-9，0为不压缩，9为最大压缩。
  /// @param nMultithread           多线程数
  /// @param bStoresLastModifiedTime 是否存储最后修改时间 
  /// @param bStoresCreationTime    是否存储创建时间 
  /// @param bStoresLastAccessTime  是否存储最后访问时间 
  /// @param bStoresAttributes      是否存储文件属性
  /// @param bCompressHeader        是否压缩头部
  /// @param bEncryptHeader         是否加密头部
  /// @return HRESULT，有关错误信息将输出到调试控制台
  /// @retval ERROR_DLL_NOT_FOUND   未找到7z.dll
  /// @retval ERROR_PROC_NOT_FOUND  从7z.dll中未找到CreateObject函数，此函数是OLE对象创建函数
  /// @retval ERROR_PATH_NOT_FOUND  未找到指定的文件夹
  /// @retval ERROR_FILE_EXISTS     目标压缩文件已存在
  /// @retval ERROR_CREATE_FAILED   创建压缩文件失败
  /// @retval ERROR_OPERATION_ABORTED CreateObject时出现错误
  /// @retval ERROR_INVALID_PARAMETER 在创建回调函数时出现错误
  /// @retval ERROR_NOT_SUPPORTED   ISetProperties接口失败或不支持
  /// @retval ERROR_BAD_ARGUMENTS   压缩器未能正确的修改所要求的属性
  /// @retval ERROR_FUNCTION_FAILED 压缩失败
  /// @retval ERROR_SUCCESS         压缩成功
  /// 回调函数原型：
  // HRESULT __stdcall SetTotal(UINT64 size)
  // HRESULT __stdcall SetCompleted(const UINT64 * completeValue)
  /// @note 函数使用lzma2算法执行压缩，且文件类型总为7z。其他类型暂未测试，不保证可用性。

  HRESULT CompressDirectory(
                    const LPWSTR wszArchiveName = L"archive.7z",
                    const LPWSTR wszPassword = L"",
                    const LPWSTR wszDirectory = L".",
                    const LPWSTR wsz7zDllPath = L"7z.dll",
                    void * funcSetTotal = NULL,
                    void * funcSetCompleted = NULL,
                    bool bRecursive = true,
                //  const char* szAlgorithm = "LZMA2",
                    bool bSolid = false,
                    unsigned int nCompressionLevel = 5,
                    unsigned int nMultithread = 4,
                    bool bStoresLastModifiedTime = true,
                    bool bStoresCreationTime = false,
                    bool bStoresLastAccessTime = false,
                    bool bStoresAttributes = true,
                    bool bCompressHeader = true,
                    bool bEncryptHeader = false);

};

extern "C" {
    /// @brief 压缩文件组
    /// @param wszArchiveName           目标压缩文件路径及名称（全路径）
    /// @param wszPassword              压缩密码，当第一字节（wszPassword[0]）为\0时表示不加密
    /// @param wszFiles                 待压缩文件路径数组
    /// @param nFiles                   文件数量，应该与数组长度一致
    /// @param wsz7zDllPath             7z.dll路径
    /// @param funcSetTotal             回调函数指针，此回调函数在压缩之前调用，指示压缩文件总大小。
    /// @param funcSetCompleted         回调函数指针，此回调函数在压缩过程中调用，指示压缩进度。
    /// @param bIgnoreFailedFiles       是否忽略压缩失败的文件
    /// @param bSolid                   是否固实压缩，后续参数为7z压缩属性
    /// @param nCompressionLevel        压缩级别，范围0-9，0为不压缩，9为最大压缩。
    /// @param nMultithread             多线程数
    /// @param bStoresLastModifiedTime  是否存储最后修改时间
    /// @param bStoresCreationTime      是否存储创建时间
    /// @param bStoresLastAccessTime    是否存储最后访问时间    
    /// @param bStoresAttributes        是否存储文件属性
    /// @param bCompressHeader          是否压缩头部
    /// @param bEncryptHeader           是否加密头部
    /// @return HRESULT，有关错误信息将输出到调试控制台
    /// @retval ERROR_DLL_NOT_FOUND   未找到7z.dll
    /// @retval ERROR_PROC_NOT_FOUND  从7z.dll中未找到CreateObject函数，此函数是OLE对象创建函数
    /// @retval ERROR_FILE_NOT_FOUND  未找到指定的文件，如果bIgnoreFailedFiles为true，则不会返回此错误，但是会于LastError中记录错误信息。有关的错误文件列表将输出到调试控制台
    /// @retval ERROR_FILE_EXISTS     目标压缩文件已存在
    /// @retval ERROR_CREATE_FAILED   创建压缩文件失败
    /// @retval ERROR_OPERATION_ABORTED CreateObject时出现错误
    /// @retval ERROR_INVALID_PARAMETER 在创建回调函数时出现错误
    /// @retval ERROR_NOT_SUPPORTED   ISetProperties接口失败或不支持
    /// @retval ERROR_BAD_ARGUMENTS   压缩器未能正确的修改所要求的属性
    /// @retval ERROR_FUNCTION_FAILED 压缩失败
    /// @retval ERROR_SUCCESS         压缩成功
    /// 回调函数原型：
    // HRESULT __stdcall SetTotal(UINT64 size)
    // HRESULT __stdcall SetCompleted(const UINT64 * completeValue)
    /// @note 函数使用lzma2算法执行压缩，且文件类型总为7z。其他类型暂未测试，不保证可用性。
    HRESULT CompressFile(
                        const LPWSTR wszArchiveName = L"archive.7z",
                        const LPWSTR wszPassword = L"",
                        const LPWSTR* wszFiles = NULL,
                        const int nFiles = 0,
                        const LPWSTR wsz7zDllPath = L"7z.dll",
                        void * funcSetTotal = NULL,
                        void * funcSetCompleted = NULL,
                        bool bIgnoreFailedFiles = false,
                    //  const char* szAlgorithm = "LZMA2",
                        bool bSolid = false,
                        unsigned int nCompressionLevel = 5,
                        unsigned int nMultithread = 4,
                        bool bStoresLastModifiedTime = true,
                        bool bStoresCreationTime = false,
                        bool bStoresLastAccessTime = false,
                        bool bStoresAttributes = true,
                        bool bCompressHeader = true,
                        bool bEncryptHeader = false
                        );
};
extern "C" {
    /// @brief 解压压缩文件
    /// @param wszArchiveName   待解压文件路径及名称（全路径）
    /// @param wszPassword      解压密码
    /// @param wszOutputFolder  输出文件夹路径，请自行确保路径尾部存在路径分隔符
    /// @param wsz7zDllPath     7z.dll路径
    /// @param funcSetTotal     回调函数指针，此回调函数在解压之前调用，指示解压文件总大小。
    /// @param funcSetCompleted 回调函数指针，此回调函数在解压过程中调用，指示解压进度。    
    /// @param nMultithread     多线程数
    /// @return HRESULT，有关错误信息将输出到调试控制台
    /// @retval ERROR_DLL_NOT_FOUND   未找到7z.dll
    /// @retval ERROR_PROC_NOT_FOUND  从7z.dll中未找到CreateObject函数，此函数是OLE对象创建函数
    /// @retval ERROR_OPERATION_ABORTED CreateObject时出现错误
    /// @retval ERROR_INVALID_PARAMETER 在创建回调函数时出现错误
    /// @retval ERROR_FILE_NOT_FOUND  未找到指定的文件
    /// @retval ERROR_FILE_CORRUPT    压缩文件损坏
    /// @retval ERROR_NOT_SUPPORTED   ISetProperties接口失败或不支持
    /// @retval ERROR_BAD_ARGUMENTS   压缩器未能正确的修改所要求的属性
    /// @retval ERROR_FUNCTION_FAILED 解压失败
    /// @retval ERROR_SUCCESS         解压成功
    /// 回调函数原型：
    // HRESULT __stdcall SetTotal(UINT64 size)
    // HRESULT __stdcall SetCompleted(const UINT64 * completeValue)
    
    HRESULT ExtractFile(const LPWSTR wszArchiveName = L"archive.7z",
                        const LPWSTR wszPassword = L"",
                        const LPWSTR wszOutputFolder = L".",
                        const LPWSTR wsz7zDllPath = L"7z.dll",
                        void * funcSetTotal = NULL,
                        void * funcSetCompleted = NULL,
                        unsigned int nMultithread = 4);
}

extern "C"{
    /// @brief 格式化打印错误信息到调试控制台
    /// @param format 要打印的格式化字符串
    /// @param  参数
    /// @return VOID
    VOID OutputDebugStringFormat(const wchar_t *format, ...);
}
#endif