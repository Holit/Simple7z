# Simple7z

从Client7z创建的简单7z压缩/解压缩接口。将繁琐的COM方法、大量的数据结构抽象并隐藏，最终提供三个主要函数供压缩。

目前只支持lzma2方法。

示例代码：

```C

// 这是一个利用Client7z静态链接库完成压缩、解压缩的例子。

// 允许使用L宏、宽字符集、打印字符串。
#include <stdio.h>

// 引入Client7z.h头文件
#include "Client7z.h"
// 引入Client7z.lib静态链接库
#pragma comment(lib, "Client7z.lib")

// 此回调函数在动作前被调用，用于指示作业大小
__declspec(nothrow) HRESULT __stdcall setTotal(UINT64 size)
{
    // 此处的size由回调函数提供
    printf("SetTotal: %d\n", size);
    return S_OK;
}
// 此回调函数在作业过程中被调用，用于指示作业进度
__declspec(nothrow) HRESULT __stdcall setCompleted(const UINT64* completeValue)
{
    printf("SetCompleted: %d\n", *completeValue);
    return S_OK;
}
int main(){  

    // 指定文件列表
    LPWSTR* wszFiles = new LPWSTR[1];
    wszFiles[0] = L"D:\\temp.py";
    // 删除已存在文件，仅用于调试
    DeleteFileW(L"archive.7z");
    // 执行压缩
    HRESULT hres = CompressFile(
        L"archive.7z",  // 压缩后的文件名
        L"",            // 压缩密码，留空代表不加密
        wszFiles,       // 文件列表
        1,              // 文件数量
        L"7z.dll",      // 7z.dll路径
        setTotal,       // 回调函数指针，用于指示作业大小
        setCompleted,   // 回调函数指针，用于指示作业进度
        false,          // 出现压缩失败的文件时返回错误
        true,           // 固实压缩
        9,              // 指定压缩级别：9 - 最大压缩
        8);             // 多线程数
                        // 还有一些参数，取默认值。
    // 如果压缩失败，输出错误信息
    // 有关其他错误类型，请参照头文件
    if (hres != ERROR_SUCCESS)
    {
        printf("Unexpected end of process, code = %d\n", hres);
    }
    printf("Compress file done\n");

    wszFiles[0] = L"D:\\tmp\\";
    DeleteFileW(L"archiveDir.7z");
    hres = CompressDirectory(L"archiveDir.7z", L"", wszFiles[0], L"7z.dll", setTotal, setCompleted, true);
    if (hres != ERROR_SUCCESS)
    {
        printf("Unexpected end of process, code = %d\n", hres);
    }
    printf("Compress directory done\n");

    hres = ExtractFile(L"archiveExtract.7z", L"", L"D:\\tmp\\extract", L"7z.dll", setTotal, setCompleted, 8);
    if (hres != ERROR_SUCCESS)
    {
        printf("Unexpected end of process, code = %d\n", hres);
    }
    printf("Extract file done\n");

    delete[] wszFiles;
}

```

相关文档请自行查阅`Client7z.h`，其中写的很详细了

`Client7z.cpp`来自`lzma2301\CPP\7zip\UI\Client7z`，替换即可继续开发，相关依赖不再提供。
