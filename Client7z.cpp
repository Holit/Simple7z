#define _CRT_SECURE_NO_WARNINGS
// Client7z.cpp

#include "StdAfx.h"
#include "Client7z.h"


#include "../../../Common/MyWindows.h"
#include "../../../Common/MyInitGuid.h"

#include "../../../Common/Defs.h"
#include "../../../Common/IntToString.h"
#include "../../../Common/StringConvert.h"

#include "../../../Windows/DLL.h"
#include "../../../Windows/FileDir.h"
#include "../../../Windows/FileFind.h"
#include "../../../Windows/FileName.h"
#include "../../../Windows/NtCheck.h"
#include "../../../Windows/PropVariant.h"
#include "../../../Windows/PropVariantConv.h"

#include "../../Common/FileStreams.h"

#include "../../Archive/IArchive.h"

#include "../../IPassword.h"
#include "../../../../C/7zVersion.h"

#include <stdio.h>

#ifdef _WIN32
extern
HINSTANCE g_hInstance;
HINSTANCE g_hInstance = NULL;
#endif

// You can find full list of all GUIDs supported by 7-Zip in Guid.txt file.
// 7z format GUID: {23170F69-40C1-278A-1000-000110070000}

#define DEFINE_GUID_ARC(name, id) Z7_DEFINE_GUID(name, \
  0x23170F69, 0x40C1, 0x278A, 0x10, 0x00, 0x00, 0x01, 0x10, id, 0x00, 0x00);

enum
{
  kId_Zip = 1,
  kId_BZip2 = 2,
  kId_7z = 7,
  kId_Xz = 0xC,
  kId_Tar = 0xEE,
  kId_GZip = 0xEF
};

// use another id, if you want to support other formats (zip, Xz, ...).
// DEFINE_GUID_ARC (CLSID_Format, kId_Zip)
// DEFINE_GUID_ARC (CLSID_Format, kId_BZip2)
// DEFINE_GUID_ARC (CLSID_Format, kId_Xz)
// DEFINE_GUID_ARC (CLSID_Format, kId_Tar)
// DEFINE_GUID_ARC (CLSID_Format, kId_GZip)
DEFINE_GUID_ARC (CLSID_Format, kId_7z)

using namespace NWindows;
using namespace NFile;
using namespace NDir;

#ifdef _WIN32
#define kDllName "7z.dll"
#else
#define kDllName "7z.so"
#endif


VOID OutputDebugStringFormat(const wchar_t *format, ...)
{
  va_list args;
  va_start(args, format);
  wchar_t buffer[1024];
  _vsnwprintf(buffer, sizeof(buffer) / sizeof(buffer[0]), format, args);
  OutputDebugStringW(buffer);
  va_end(args);
}

static void Convert_UString_to_AString(const UString &s, AString &temp)
{
  int codePage = CP_OEMCP;
  /*
  int g_CodePage = -1;
  int codePage = g_CodePage;
  if (codePage == -1)
    codePage = CP_OEMCP;
  if (codePage == CP_UTF8)
    ConvertUnicodeToUTF8(s, temp);
  else
  */
    UnicodeStringToMultiByte2(temp, s, (UINT)codePage);
}


static HRESULT IsArchiveItemProp(IInArchive *archive, UInt32 index, PROPID propID, bool &result)
{
  NCOM::CPropVariant prop;
  RINOK(archive->GetProperty(index, propID, &prop))
  if (prop.vt == VT_BOOL)
    result = VARIANT_BOOLToBool(prop.boolVal);
  else if (prop.vt == VT_EMPTY)
    result = false;
  else
    return E_FAIL;
  return S_OK;
}

static HRESULT IsArchiveItemFolder(IInArchive *archive, UInt32 index, bool &result)
{
  return IsArchiveItemProp(archive, index, kpidIsDir, result);
}


static const wchar_t * const kEmptyFileAlias = L"[Content]";


//////////////////////////////////////////////////////////////
// Archive Open callback class


class CArchiveOpenCallback Z7_final:
  public IArchiveOpenCallback,
  public ICryptoGetTextPassword,
  public CMyUnknownImp
{
  Z7_IFACES_IMP_UNK_2(IArchiveOpenCallback, ICryptoGetTextPassword)
public:

  bool PasswordIsDefined;
  UString Password;

  CArchiveOpenCallback() : PasswordIsDefined(false) {}
};

Z7_COM7F_IMF(CArchiveOpenCallback::SetTotal(const UInt64 * /* files */, const UInt64 * /* bytes */))
{
  return S_OK;
}

Z7_COM7F_IMF(CArchiveOpenCallback::SetCompleted(const UInt64 * /* files */, const UInt64 * /* bytes */))
{
  return S_OK;
}
  
Z7_COM7F_IMF(CArchiveOpenCallback::CryptoGetTextPassword(BSTR *password))
{
  if (!PasswordIsDefined)
  {
    // You can ask real password here from user
    // Password = GetPassword(OutStream);
    // PasswordIsDefined = true;
    OutputDebugStringA("Password is not defined");
    return E_ABORT;
  }
  return StringToBstr(Password, password);
}



static const char * const kIncorrectCommand = "incorrect command";

//////////////////////////////////////////////////////////////
// Archive Extracting callback class

static const char * const kTestingString    =  "Testing     ";
static const char * const kExtractingString =  "Extracting  ";
static const char * const kSkippingString   =  "Skipping    ";
static const char * const kReadingString    =  "Reading     ";

static const char * const kUnsupportedMethod = "Unsupported Method";
static const char * const kCRCFailed = "CRC Failed";
static const char * const kDataError = "Data Error";
static const char * const kUnavailableData = "Unavailable data";
static const char * const kUnexpectedEnd = "Unexpected end of data";
static const char * const kDataAfterEnd = "There are some data after the end of the payload data";
static const char * const kIsNotArc = "Is not archive";
static const char * const kHeadersError = "Headers Error";


struct CArcTime
{
  FILETIME FT;
  UInt16 Prec;
  Byte Ns100;
  bool Def;

  CArcTime()
  {
    Clear();
  }

  void Clear()
  {
    FT.dwHighDateTime = FT.dwLowDateTime = 0;
    Prec = 0;
    Ns100 = 0;
    Def = false;
  }

  bool IsZero() const
  {
    return FT.dwLowDateTime == 0 && FT.dwHighDateTime == 0 && Ns100 == 0;
  }

  int GetNumDigits() const
  {
    if (Prec == k_PropVar_TimePrec_Unix ||
        Prec == k_PropVar_TimePrec_DOS)
      return 0;
    if (Prec == k_PropVar_TimePrec_HighPrec)
      return 9;
    if (Prec == k_PropVar_TimePrec_0)
      return 7;
    int digits = (int)Prec - (int)k_PropVar_TimePrec_Base;
    if (digits < 0)
      digits = 0;
    return digits;
  }

  void Write_To_FiTime(CFiTime &dest) const
  {
   #ifdef _WIN32
    dest = FT;
   #else
    if (FILETIME_To_timespec(FT, dest))
    if ((Prec == k_PropVar_TimePrec_Base + 8 ||
         Prec == k_PropVar_TimePrec_Base + 9)
        && Ns100 != 0)
    {
      dest.tv_nsec += Ns100;
    }
   #endif
  }

  void Set_From_Prop(const PROPVARIANT &prop)
  {
    FT = prop.filetime;
    unsigned prec = 0;
    unsigned ns100 = 0;
    const unsigned prec_Temp = prop.wReserved1;
    if (prec_Temp != 0
        && prec_Temp <= k_PropVar_TimePrec_1ns
        && prop.wReserved3 == 0)
    {
      const unsigned ns100_Temp = prop.wReserved2;
      if (ns100_Temp < 100)
      {
        ns100 = ns100_Temp;
        prec = prec_Temp;
      }
    }
    Prec = (UInt16)prec;
    Ns100 = (Byte)ns100;
    Def = true;
  }
};



class CArchiveExtractCallback Z7_final:
  public IArchiveExtractCallback,
  public ICryptoGetTextPassword,
  public CMyUnknownImp
{
  Z7_IFACES_IMP_UNK_2(IArchiveExtractCallback, ICryptoGetTextPassword)
  Z7_IFACE_COM7_IMP(IProgress)

  CMyComPtr<IInArchive> _archiveHandler;
  FString _directoryPath;  // Output directory
  UString _filePath;       // name inside arcvhive
  FString _diskFilePath;   // full path to file on disk
  bool _extractMode;
  struct CProcessedFileInfo
  {
    CArcTime MTime;
    UInt32 Attrib;
    bool isDir;
    bool Attrib_Defined;
  } _processedFileInfo;

  COutFileStream *_outFileStreamSpec;
  CMyComPtr<ISequentialOutStream> _outFileStream;

typedef HRESULT (__stdcall *ProgressCallback_SetTotal)(UInt64 size);
typedef HRESULT (__stdcall *ProgressCallback_SetCompleted)(const UInt64 *completeValue);
  PVOID funcSetTotal = NULL;
  PVOID funcSetCompleted = NULL;

public:
  void Init(IInArchive *archiveHandler, const FString &directoryPath);

  void InitCallback(PVOID pSetTotal, PVOID pSetComplete)
  {
    if(pSetTotal != NULL && pSetComplete != NULL)
    {
      funcSetTotal = pSetTotal;
      funcSetCompleted = pSetComplete;
    }
    else
    {
      SetLastError(ERROR_INVALID_PARAMETER);
    }
  }
  UInt64 NumErrors;
  bool PasswordIsDefined;
  UString Password;

  CArchiveExtractCallback() : PasswordIsDefined(false) {}
};


void CArchiveExtractCallback::Init(IInArchive *archiveHandler, const FString &directoryPath)
{
  NumErrors = 0;
  _archiveHandler = archiveHandler;
  _directoryPath = directoryPath;
  NName::NormalizeDirPathPrefix(_directoryPath);
}



Z7_COM7F_IMF(CArchiveExtractCallback::SetTotal(UInt64 size))
{
  try {
    if(this->funcSetTotal != NULL)
    {
      ProgressCallback_SetTotal pSetTotal = (ProgressCallback_SetTotal)this->funcSetTotal;
      return pSetTotal(size);
    }
    return S_OK;
  } catch (...) {
    SetLastError(ERROR_FUNCTION_FAILED);
    return E_FAIL;
  }
}

Z7_COM7F_IMF(CArchiveExtractCallback::SetCompleted(const UInt64 * completeValue ))
{
  try {
    if(this->funcSetCompleted != NULL)
    {
      ProgressCallback_SetCompleted pSetCompleted = (ProgressCallback_SetCompleted)this->funcSetCompleted;
      return pSetCompleted(completeValue);
    }
    return S_OK;
  } catch (...) {
    SetLastError(ERROR_FUNCTION_FAILED);
    return E_FAIL;
  }
}

Z7_COM7F_IMF(CArchiveExtractCallback::GetStream(UInt32 index,
    ISequentialOutStream **outStream, Int32 askExtractMode))
{
  *outStream = NULL;
  _outFileStream.Release();

  {
    // Get Name
    NCOM::CPropVariant prop;
    RINOK(_archiveHandler->GetProperty(index, kpidPath, &prop))
    
    UString fullPath;
    if (prop.vt == VT_EMPTY)
      fullPath = kEmptyFileAlias;
    else
    {
      if (prop.vt != VT_BSTR)
        return E_FAIL;
      fullPath = prop.bstrVal;
    }
    _filePath = fullPath;
  }

  if (askExtractMode != NArchive::NExtract::NAskMode::kExtract)
    return S_OK;

  {
    // Get Attrib
    NCOM::CPropVariant prop;
    RINOK(_archiveHandler->GetProperty(index, kpidAttrib, &prop))
    if (prop.vt == VT_EMPTY)
    {
      _processedFileInfo.Attrib = 0;
      _processedFileInfo.Attrib_Defined = false;
    }
    else
    {
      if (prop.vt != VT_UI4)
        return E_FAIL;
      _processedFileInfo.Attrib = prop.ulVal;
      _processedFileInfo.Attrib_Defined = true;
    }
  }

  RINOK(IsArchiveItemFolder(_archiveHandler, index, _processedFileInfo.isDir))

  {
    _processedFileInfo.MTime.Clear();
    // Get Modified Time
    NCOM::CPropVariant prop;
    RINOK(_archiveHandler->GetProperty(index, kpidMTime, &prop))
    switch (prop.vt)
    {
      case VT_EMPTY:
        // _processedFileInfo.MTime = _utcMTimeDefault;
        break;
      case VT_FILETIME:
        _processedFileInfo.MTime.Set_From_Prop(prop);
        break;
      default:
        return E_FAIL;
    }

  }
  {
    // Get Size
    NCOM::CPropVariant prop;
    RINOK(_archiveHandler->GetProperty(index, kpidSize, &prop))
    UInt64 newFileSize;
    /* bool newFileSizeDefined = */ ConvertPropVariantToUInt64(prop, newFileSize);
  }

  
  {
    // Create folders for file
    int slashPos = _filePath.ReverseFind_PathSepar();
    if (slashPos >= 0)
      CreateComplexDir(_directoryPath + us2fs(_filePath.Left(slashPos)));
  }

  FString fullProcessedPath = _directoryPath + us2fs(_filePath);
  _diskFilePath = fullProcessedPath;

  if (_processedFileInfo.isDir)
  {
    CreateComplexDir(fullProcessedPath);
  }
  else
  {
    NFind::CFileInfo fi;
    if (fi.Find(fullProcessedPath))
    {
      if (!DeleteFileAlways(fullProcessedPath))
      {
        OutputDebugStringFormat(L"Cannot delete output file %s", fullProcessedPath);
        return E_ABORT;
      }
    }
    
    _outFileStreamSpec = new COutFileStream;
    CMyComPtr<ISequentialOutStream> outStreamLoc(_outFileStreamSpec);
    if (!_outFileStreamSpec->Open(fullProcessedPath, CREATE_ALWAYS))
    {
      OutputDebugStringFormat(L"Cannot open output file %s", fullProcessedPath);
      return E_ABORT;
    }
    _outFileStream = outStreamLoc;
    *outStream = outStreamLoc.Detach();
  }
  return S_OK;
}

Z7_COM7F_IMF(CArchiveExtractCallback::PrepareOperation(Int32 askExtractMode))
{
  _extractMode = false;
  switch (askExtractMode)
  {
    case NArchive::NExtract::NAskMode::kExtract:  _extractMode = true; break;
  }
  switch (askExtractMode)
  {
    case NArchive::NExtract::NAskMode::kExtract:  OutputDebugStringA(kExtractingString); break;
    case NArchive::NExtract::NAskMode::kTest:  OutputDebugStringA(kTestingString); break;
    case NArchive::NExtract::NAskMode::kSkip:  OutputDebugStringA(kSkippingString); break;
    case NArchive::NExtract::NAskMode::kReadExternal: OutputDebugStringA(kReadingString); break;
    default:
      OutputDebugStringA("??? "); break;
  }
  OutputDebugStringW(_filePath);
  return S_OK;
}

Z7_COM7F_IMF(CArchiveExtractCallback::SetOperationResult(Int32 operationResult))
{
  switch (operationResult)
  {
    case NArchive::NExtract::NOperationResult::kOK:
      break;
    default:
    {
      NumErrors++;
      OutputDebugStringA("  :  ");
      const char *s = NULL;
      switch (operationResult)
      {
        case NArchive::NExtract::NOperationResult::kUnsupportedMethod:
          s = kUnsupportedMethod;
          break;
        case NArchive::NExtract::NOperationResult::kCRCError:
          s = kCRCFailed;
          break;
        case NArchive::NExtract::NOperationResult::kDataError:
          s = kDataError;
          break;
        case NArchive::NExtract::NOperationResult::kUnavailable:
          s = kUnavailableData;
          break;
        case NArchive::NExtract::NOperationResult::kUnexpectedEnd:
          s = kUnexpectedEnd;
          break;
        case NArchive::NExtract::NOperationResult::kDataAfterEnd:
          s = kDataAfterEnd;
          break;
        case NArchive::NExtract::NOperationResult::kIsNotArc:
          s = kIsNotArc;
          break;
        case NArchive::NExtract::NOperationResult::kHeadersError:
          s = kHeadersError;
          break;
      }
      if (s)
      {
        OutputDebugStringA("Error : ");
        OutputDebugStringA(s);
      }
      else
      {
        char temp[16];
        ConvertUInt32ToString((UInt32)operationResult, temp);
        OutputDebugStringA("Error #");
        OutputDebugStringA(temp);
      }
    }
  }

  if (_outFileStream)
  {
    if (_processedFileInfo.MTime.Def)
    {
      CFiTime ft;
      _processedFileInfo.MTime.Write_To_FiTime(ft);
      _outFileStreamSpec->SetMTime(&ft);
    }
    RINOK(_outFileStreamSpec->Close())
  }
  _outFileStream.Release();
  if (_extractMode && _processedFileInfo.Attrib_Defined)
    SetFileAttrib_PosixHighDetect(_diskFilePath, _processedFileInfo.Attrib);
  return S_OK;
}


Z7_COM7F_IMF(CArchiveExtractCallback::CryptoGetTextPassword(BSTR *password))
{
  if (!PasswordIsDefined)
  {
    // You can ask real password here from user
    // Password = GetPassword(OutStream);
    // PasswordIsDefined = true;
    OutputDebugStringA("Password is not defined");
    return E_ABORT;
  }
  return StringToBstr(Password, password);
}



//////////////////////////////////////////////////////////////
// Archive Creating callback class

struct CDirItem: public NWindows::NFile::NFind::CFileInfoBase
{
  UString Path_For_Handler;
  FString FullPath; // for filesystem

  CDirItem(const NWindows::NFile::NFind::CFileInfo &fi):
      CFileInfoBase(fi)
    {}
};

class CArchiveUpdateCallback Z7_final:
  public IArchiveUpdateCallback2,
  public ICryptoGetTextPassword2,
  public CMyUnknownImp
{
  Z7_IFACES_IMP_UNK_2(IArchiveUpdateCallback2, ICryptoGetTextPassword2)
  Z7_IFACE_COM7_IMP(IProgress)
  Z7_IFACE_COM7_IMP(IArchiveUpdateCallback)

typedef HRESULT (__stdcall *ProgressCallback_SetTotal)(UInt64 size);
typedef HRESULT (__stdcall *ProgressCallback_SetCompleted)(const UInt64 *completeValue);
  PVOID funcSetTotal = NULL;
  PVOID funcSetCompleted = NULL;

public:
  CRecordVector<UInt64> VolumesSizes;
  UString VolName;
  UString VolExt;

  FString DirPrefix;
  const CObjectVector<CDirItem> *DirItems;

  bool PasswordIsDefined;
  UString Password;
  bool AskPassword;

  bool m_NeedBeClosed;

  FStringVector FailedFiles;
  CRecordVector<HRESULT> FailedCodes;

  CArchiveUpdateCallback():
      DirItems(NULL),
      PasswordIsDefined(false),
      AskPassword(false)
      {}

  ~CArchiveUpdateCallback() { Finilize(); }
  HRESULT Finilize();

  void Init(const CObjectVector<CDirItem> *dirItems)
  {
    DirItems = dirItems;
    m_NeedBeClosed = false;
    FailedFiles.Clear();
    FailedCodes.Clear();
  }
  void InitCallback(PVOID pSetTotal, PVOID pSetComplete)
  {
    if(pSetTotal != NULL && pSetComplete != NULL)
    {
      funcSetTotal = pSetTotal;
      funcSetCompleted = pSetComplete;
    }
    else
    {
      SetLastError(ERROR_INVALID_PARAMETER);
    }

  }
};
Z7_COM7F_IMF(CArchiveUpdateCallback::SetTotal(UInt64 size))
{
  try {
    if(this->funcSetTotal != NULL)
    {
      ProgressCallback_SetTotal pSetTotal = (ProgressCallback_SetTotal)this->funcSetTotal;
      return pSetTotal(size);
    }
    return S_OK;
  } catch (...) {
    SetLastError(ERROR_FUNCTION_FAILED);
    return E_FAIL;
  }
}
Z7_COM7F_IMF(CArchiveUpdateCallback::SetCompleted(const UInt64 * completeValue ))
{
  try {
    if(this->funcSetCompleted != NULL)
    {
      ProgressCallback_SetCompleted pSetCompleted = (ProgressCallback_SetCompleted)this->funcSetCompleted;
      return pSetCompleted(completeValue);
    }
    return S_OK;
  } catch (...) {
    SetLastError(ERROR_FUNCTION_FAILED);
    return E_FAIL;
  }
}

Z7_COM7F_IMF(CArchiveUpdateCallback::GetUpdateItemInfo(UInt32 /* index */,
      Int32 *newData, Int32 *newProperties, UInt32 *indexInArchive))
{
  if (newData)
    *newData = BoolToInt(true);
  if (newProperties)
    *newProperties = BoolToInt(true);
  if (indexInArchive)
    *indexInArchive = (UInt32)(Int32)-1;
  return S_OK;
}

Z7_COM7F_IMF(CArchiveUpdateCallback::GetProperty(UInt32 index, PROPID propID, PROPVARIANT *value))
{
  NCOM::CPropVariant prop;
  
  if (propID == kpidIsAnti)
  {
    prop = false;
    prop.Detach(value);
    return S_OK;
  }

  {
    const CDirItem &di = (*DirItems)[index];
    switch (propID)
    {
      case kpidPath:  prop = di.Path_For_Handler; break;
      case kpidIsDir:  prop = di.IsDir(); break;
      case kpidSize:  prop = di.Size; break;
      case kpidCTime:  PropVariant_SetFrom_FiTime(prop, di.CTime); break;
      case kpidATime:  PropVariant_SetFrom_FiTime(prop, di.ATime); break;
      case kpidMTime:  PropVariant_SetFrom_FiTime(prop, di.MTime); break;
      case kpidAttrib:  prop = (UInt32)di.GetWinAttrib(); break;
      case kpidPosixAttrib: prop = (UInt32)di.GetPosixAttrib(); break;
    }
  }
  prop.Detach(value);
  return S_OK;
}

HRESULT CArchiveUpdateCallback::Finilize()
{
  // if (m_NeedBeClosed)
  // {
  //   OutputDebugStringANewLine();
  //   m_NeedBeClosed = false;
  // }
  return S_OK;
}

static void GetStream2(const wchar_t *name)
{
  OutputDebugStringFormat(L"Compressing %s" ,name);
  // if (name[0] == 0)
  //   name = kEmptyFileAlias;
  // OutputDebugStringA(name);
}

Z7_COM7F_IMF(CArchiveUpdateCallback::GetStream(UInt32 index, ISequentialInStream **inStream))
{
  RINOK(Finilize())

  const CDirItem &dirItem = (*DirItems)[index];
  GetStream2(dirItem.Path_For_Handler);
 
  if (dirItem.IsDir())
    return S_OK;

  {
    CInFileStream *inStreamSpec = new CInFileStream;
    CMyComPtr<ISequentialInStream> inStreamLoc(inStreamSpec);
    FString path = DirPrefix + dirItem.FullPath;
    if (!inStreamSpec->Open(path))
    {
      const DWORD sysError = ::GetLastError();
      FailedCodes.Add(HRESULT_FROM_WIN32(sysError));
      FailedFiles.Add(path);
      // if (systemError == ERROR_SHARING_VIOLATION)
      {
        SetLastError(sysError);
        OutputDebugStringFormat(L"GetStream Error: %s\n", path);
        // OutputDebugStringA(NError::MyFormatMessageW(systemError));
        return S_FALSE;
      }
      // return sysError;
    }
    *inStream = inStreamLoc.Detach();
  }
  return S_OK;
}

Z7_COM7F_IMF(CArchiveUpdateCallback::SetOperationResult(Int32 /* operationResult */))
{
  m_NeedBeClosed = true;
  return S_OK;
}

Z7_COM7F_IMF(CArchiveUpdateCallback::GetVolumeSize(UInt32 index, UInt64 *size))
{
  if (VolumesSizes.Size() == 0)
    return S_FALSE;
  if (index >= (UInt32)VolumesSizes.Size())
    index = VolumesSizes.Size() - 1;
  *size = VolumesSizes[index];
  return S_OK;
}

Z7_COM7F_IMF(CArchiveUpdateCallback::GetVolumeStream(UInt32 index, ISequentialOutStream **volumeStream))
{
  wchar_t temp[16];
  ConvertUInt32ToString(index + 1, temp);
  UString res = temp;
  while (res.Len() < 2)
    res.InsertAtFront(L'0');
  UString fileName = VolName;
  fileName.Add_Dot();
  fileName += res;
  fileName += VolExt;
  COutFileStream *streamSpec = new COutFileStream;
  CMyComPtr<ISequentialOutStream> streamLoc(streamSpec);
  if (!streamSpec->Create(us2fs(fileName), false))
    return GetLastError_noZero_HRESULT();
  *volumeStream = streamLoc.Detach();
  return S_OK;
}

Z7_COM7F_IMF(CArchiveUpdateCallback::CryptoGetTextPassword2(Int32 *passwordIsDefined, BSTR *password))
{
  if (!PasswordIsDefined)
  {
    if (AskPassword)
    {
      // You can ask real password here from user
      // Password = GetPassword(OutStream);
      // PasswordIsDefined = true;
      OutputDebugStringA("Password is not defined");
      return E_ABORT;
    }
  }
  *passwordIsDefined = BoolToInt(PasswordIsDefined);
  return StringToBstr(Password, password);
}


// Main function

#if defined(_UNICODE) && !defined(_WIN64) && !defined(UNDER_CE)
#define NT_CHECK_FAIL_ACTION OutputDebugStringA("Unsupported Windows version"); return 1;
#endif


void AddDirFileInfo(
	const UString &prefix,
	const UString &fullPathName,
	NFind::CFileInfo &fileInfo,
	CObjectVector<CDirItem>& dirItems)
{
	CDirItem item(fileInfo);
	item.Attrib = fileInfo.Attrib;
	item.Size = fileInfo.Size;
	item.CTime = fileInfo.CTime;
	item.ATime = fileInfo.ATime;
	item.MTime = fileInfo.MTime;
  
	item.Path_For_Handler = prefix + fileInfo.Name;
	item.FullPath = fullPathName;
	dirItems.Add(item);
}

static void EnumerateDirectory(
	const UString &baseFolderPrefix,
	const UString &directory,
	const UString &prefix,
	CObjectVector<CDirItem>& dirItems)
{
	NFind::CEnumerator enumerator;
	enumerator.SetDirPrefix(baseFolderPrefix + directory);
	NFind::CFileInfo fileInfo;
	while (enumerator.Next(fileInfo))
	{
		AddDirFileInfo(prefix, directory + fileInfo.Name, fileInfo, dirItems);
		if (fileInfo.IsDir())
		{
			EnumerateDirectory(baseFolderPrefix, directory + fileInfo.Name + WCHAR_PATH_SEPARATOR,
			prefix + fileInfo.Name + WCHAR_PATH_SEPARATOR, dirItems);
		}
	}
}

HRESULT CompressDirectory(const LPWSTR wszArchiveName,
                     const LPWSTR wszPassword,
                     const LPWSTR wszDirectory,
                     const LPWSTR wsz7zDllPath,
                     void * funcSetTotal,
                     void * funcSetCompleted,
                     bool bRecursive,
                     bool bSolid,
                     unsigned int nCompressionLevel,
                     unsigned int nMultithread,
                     bool bStoresLastModifiedTime,
                     bool bStoresCreationTime,
                     bool bStoresLastAccessTime,
                     bool bStoresAttributes,
                     bool bCompressHeader,
                     bool bEncryptHeader)
{
  //Initialize
  // load 7z.dll
  NDLL::CLibrary lib;
  if (!lib.Load(wsz7zDllPath))
  {
    OutputDebugStringW(L"Cannot load 7-zip library");
    return ERROR_DLL_NOT_FOUND;
  }
  // get CreateObject
  Func_CreateObject
     f_CreateObject = Z7_GET_PROC_ADDRESS(
  Func_CreateObject, lib.Get_HMODULE(),
      "CreateObject");
  if (!f_CreateObject)
  {
    OutputDebugStringW(L"Cannot get CreateObject");
    return ERROR_PROC_NOT_FOUND;
  }
  // prepare password
  UString password = wszPassword;
  bool passwordIsDefined = wszPassword[0] != L'\0';
  // prepare files
  // if path is not a directory or not exists
  if(GetFileAttributesW(wszDirectory) == INVALID_FILE_ATTRIBUTES)
  {
    OutputDebugStringFormat(L"Directory %s not exists", wszDirectory);
    return ERROR_PATH_NOT_FOUND;
  }
  CObjectVector<CDirItem> dirItems = CObjectVector<CDirItem>();
  if(!bRecursive)
  {
    // not tested, may cause unexpected behavior
    AddDirFileInfo(UString(), wszDirectory, NFind::CFileInfo(), dirItems);
  }
  else
  {
    EnumerateDirectory(UString(), wszDirectory, L"", dirItems);
  }

  //Check if file is alreay exists
  if(GetFileAttributesW(wszArchiveName) != INVALID_FILE_ATTRIBUTES)
  {
    OutputDebugStringFormat(L"File %s already exists", wszArchiveName);
    return ERROR_FILE_EXISTS;
  }
  COutFileStream *outFileStreamSpec = new COutFileStream;
  CMyComPtr<IOutStream> outFileStream = outFileStreamSpec;
  if (!outFileStreamSpec->Create(wszArchiveName, false))
  {
    OutputDebugStringFormat(L"can't create archive file %s", wszArchiveName);
    return ERROR_CREATE_FAILED;
  }
  CMyComPtr<IOutArchive> outArchive;
  if (f_CreateObject(&CLSID_Format, &IID_IOutArchive, (void **)&outArchive) != S_OK)
  {
    OutputDebugStringW(L"Cannot get class object");
    return ERROR_OPERATION_ABORTED;
  }
  CArchiveUpdateCallback *updateCallbackSpec = new CArchiveUpdateCallback;
  CMyComPtr<IArchiveUpdateCallback2> updateCallback(updateCallbackSpec);
  updateCallbackSpec->Init(&dirItems);
  updateCallbackSpec->PasswordIsDefined = passwordIsDefined;
  updateCallbackSpec->Password = password;
  DWORD dwLastError = GetLastError();
  if(funcSetTotal != NULL && funcSetCompleted != NULL)
  {
    updateCallbackSpec->InitCallback(funcSetTotal, funcSetCompleted);
    DWORD dwLastError2 = GetLastError();
    if(dwLastError != dwLastError2)
    {
      OutputDebugStringFormat(L"InitCallback failed %d", dwLastError2);
      return ERROR_INVALID_PARAMETER;
    }
  }
    // set archive propeties
    {
      const wchar_t *names[] =
      {
        L"m",
        L"s",
        L"x",
        L"mt",
        L"tm",
        L"tc",
        L"ta",
        L"tr",
        L"hc",
        L"he",
      };
      const unsigned kNumProps = Z7_ARRAY_SIZE(names);
      NCOM::CPropVariant values[kNumProps] =
      {
        L"lzma2", // This is disabled in prototype, which may add in future
        bSolid,    // solid mode OFF
        (UInt32)nCompressionLevel, // compression level = 9 - ultra
        (UInt32)nMultithread, // multithread
        bStoresLastModifiedTime, // store last modified time
        bStoresCreationTime, // store creation time
        bStoresLastAccessTime, // store last access time
        bStoresAttributes, // store attributes
        bCompressHeader, // compress header
        bEncryptHeader // encrypt header
      };
      CMyComPtr<ISetProperties> setProperties;
      outArchive->QueryInterface(IID_ISetProperties, (void **)&setProperties);
      if (!setProperties)
      {
        OutputDebugStringW(L"ISetProperties unsupported");
        return ERROR_NOT_SUPPORTED;
      }
      if (setProperties->SetProperties(names, values, kNumProps) != S_OK)
      {
        OutputDebugStringW(L"SetProperties() error");
        return ERROR_BAD_ARGUMENTS;
      }
    }
    // update archive
    HRESULT result = outArchive->UpdateItems(outFileStream, dirItems.Size(), updateCallback);
    
    updateCallbackSpec->Finilize();

    if (result != S_OK)
    {
      OutputDebugStringW(L"Update Error");
      return ERROR_FUNCTION_FAILED;
    }
    
    FOR_VECTOR (i, updateCallbackSpec->FailedFiles)
    {
      OutputDebugStringFormat(L"Error for file %s", updateCallbackSpec->FailedFiles[i]);
    }
    
    if (updateCallbackSpec->FailedFiles.Size() != 0)
      return ERROR_FUNCTION_FAILED;

    return ERROR_SUCCESS;

}

                 
HRESULT CompressFile(const LPWSTR wszArchiveName,
                     const LPWSTR wszPassword,
                     const LPWSTR* wszFiles ,
                     const int nFiles,
                     const LPWSTR wsz7zDllPath,
                     /*progress callback*
__declspec(nothrow) HRESULT __stdcall CArchiveUpdateCallback::SetTotal(UInt64 size) throw()
__declspec(nothrow) HRESULT __stdcall CArchiveUpdateCallback::SetCompleted(const UInt64 * completeValue ) throw()
                      */
                     void * funcSetTotal,
                     void * funcSetCompleted,
                     bool bIgnoreFailedFiles,
                    //  const char* szAlgorithm = "LZMA2",
                     bool bSolid,
                     unsigned int nCompressionLevel,
                     unsigned int nMultithread,
                     bool bStoresLastModifiedTime,
                     bool bStoresCreationTime,
                     bool bStoresLastAccessTime,
                     bool bStoresAttributes,
                     bool bCompressHeader,
                     bool bEncryptHeader
                     )
{
  //Initialize 
  // load 7z.dll
  NDLL::CLibrary lib;
  if (!lib.Load(wsz7zDllPath))
  {
    OutputDebugStringW(L"Cannot load 7-zip library");
    return ERROR_DLL_NOT_FOUND;
  }
  // get CreateObject
  Func_CreateObject
     f_CreateObject = Z7_GET_PROC_ADDRESS(
  Func_CreateObject, lib.Get_HMODULE(),
      "CreateObject");
  if (!f_CreateObject)
  {
    OutputDebugStringW(L"Cannot get CreateObject");
    return ERROR_PROC_NOT_FOUND;
  }
  // prepare password 
  UString password = wszPassword;
  bool passwordIsDefined = wszPassword[0] != L'\0';
  // prepare files

  CObjectVector<CDirItem> dirItems;
  {
    unsigned i;
    for (i = 0; i < nFiles; i++)
    {
      const FString name = wszFiles[i];
      
      NFind::CFileInfo fi;
      if (!fi.Find(name))
      {
        OutputDebugStringFormat(L"Can't find file %s", name);
        SetLastError(ERROR_FILE_NOT_FOUND);
        if(!bIgnoreFailedFiles)
          return ERROR_FILE_NOT_FOUND;
      }

      CDirItem di(fi);
      
      di.Path_For_Handler = fi.Name;
      // change this name to basename

      di.FullPath = name;
      dirItems.Add(di);
    }
  }
  //Check if file is alreay exists
  if(GetFileAttributesW(wszArchiveName) != INVALID_FILE_ATTRIBUTES)
  {
    OutputDebugStringFormat(L"File %s already exists", wszArchiveName);
    return ERROR_FILE_EXISTS;
  }
   COutFileStream *outFileStreamSpec = new COutFileStream;
   CMyComPtr<IOutStream> outFileStream = outFileStreamSpec;
   if (!outFileStreamSpec->Create(wszArchiveName, false))
   {
    OutputDebugStringFormat(L"can't create archive file %s", wszArchiveName);
    return ERROR_CREATE_FAILED;
   }

    CMyComPtr<IOutArchive> outArchive;
    if (f_CreateObject(&CLSID_Format, &IID_IOutArchive, (void **)&outArchive) != S_OK)
    {
      OutputDebugStringW(L"Cannot get class object");
      return ERROR_OPERATION_ABORTED;
    }
    CArchiveUpdateCallback *updateCallbackSpec = new CArchiveUpdateCallback;
    CMyComPtr<IArchiveUpdateCallback2> updateCallback(updateCallbackSpec);
    updateCallbackSpec->Init(&dirItems);
    updateCallbackSpec->PasswordIsDefined = passwordIsDefined;
    updateCallbackSpec->Password = password;
    
    DWORD dwLastError = GetLastError();
    if(funcSetTotal != NULL && funcSetCompleted != NULL)
    {
      updateCallbackSpec->InitCallback(funcSetTotal, funcSetCompleted);
      DWORD dwLastError2 = GetLastError();
      if(dwLastError != dwLastError2)
      {
        OutputDebugStringFormat(L"InitCallback failed %d", dwLastError2);
        return ERROR_INVALID_PARAMETER;
      }
    }

    // set archive propeties
      /*
    Support properties(According to 7z.h):
    Parameter  Default  Description  
    x=[0 | 1 | 3 | 5 | 7 | 9 ]  5  Sets level of compression.  
    yx=[0 | 1 | 3 | 5 | 7 | 9 ]  5  Sets level of file analysis.  
    memuse=[ p{N_Percents} | {N}b | {N}k | {N}m | {N}g | {N}t]   Sets memory usage size.  
    s=[off | on | [e] [{N}f] [{N}b | {N}k | {N}m | {N}g | {N}t]]  on  Sets solid mode.  
    qs=[off | on]  off  Sort files by type in solid archives.  
    f=[off | on | FilterID]  on  Enables or disables filters. FilterID: Delta:{N}, BCJ, BCJ2, ARM64, ARM, ARMT, IA64, PPC, SPARC.  
    hc=[off | on]  on Enables or disables archive header compressing.  
    he=[off | on]  off Enables or disables archive header encryption.  
    b{C1}[s{S1}]:{C2}[s{S2}]   Sets binding between coders. 
    {N}={MethodID}[:param1][:param2][..] LZMA2 Sets a method: LZMA, LZMA2, PPMd, BZip2, Deflate, Delta, BCJ, BCJ2, Copy.  
    mt=[off | on | {N}]  on Sets multithreading mode.  
    mtf=[off | on]  on  Set multithreading mode for filters.  
    tm=[off | on]  on  Stores last Modified timestamps for files.  
    tc=[off | on]  off  Stores Creation timestamps for files.  
    ta=[off | on]  off  Stores last Access timestamps for files.  
    tr=[off | on]  on  Stores file attributes.  
    */
    {
      const wchar_t *names[] =
      {
        L"m",
        L"s",
        L"x",
        L"mt",
        L"tm",
        L"tc",
        L"ta",
        L"tr",
        L"hc",
        L"he",
      };
      const unsigned kNumProps = Z7_ARRAY_SIZE(names);
      NCOM::CPropVariant values[kNumProps] =
      {
        L"lzma2", // This is disabled in prototype, which may add in future
        bSolid,    // solid mode OFF
        (UInt32)nCompressionLevel, // compression level = 9 - ultra
        (UInt32)nMultithread, // multithread
        bStoresLastModifiedTime, // store last modified time
        bStoresCreationTime, // store creation time
        bStoresLastAccessTime, // store last access time
        bStoresAttributes, // store attributes
        bCompressHeader, // compress header
        bEncryptHeader // encrypt header
      };
      CMyComPtr<ISetProperties> setProperties;
      outArchive->QueryInterface(IID_ISetProperties, (void **)&setProperties);
      if (!setProperties)
      {
        OutputDebugStringW(L"ISetProperties unsupported");
        return ERROR_NOT_SUPPORTED;
      }
      if (setProperties->SetProperties(names, values, kNumProps) != S_OK)
      {
        OutputDebugStringW(L"SetProperties() error");
        return ERROR_BAD_ARGUMENTS;
      }
    }
    // update archive
    HRESULT result = outArchive->UpdateItems(outFileStream, dirItems.Size(), updateCallback);
    
    updateCallbackSpec->Finilize();

    if (result != S_OK)
    {
      OutputDebugStringW(L"Update Error");
      return ERROR_FUNCTION_FAILED;
    }
    
    FOR_VECTOR (i, updateCallbackSpec->FailedFiles)
    {
      OutputDebugStringFormat(L"Error for file %s", updateCallbackSpec->FailedFiles[i]);
    }
    
    if (updateCallbackSpec->FailedFiles.Size() != 0)
      return ERROR_FUNCTION_FAILED;

    return ERROR_SUCCESS;
}

HRESULT ExtractFile(const LPWSTR wszArchiveName,
                     const LPWSTR wszPassword,
                     const LPWSTR wszOutputFolder,
                     const LPWSTR wsz7zDllPath,
                     void * funcSetTotal,
                     void * funcSetCompleted,
                     unsigned int nMultithread)
{
  
  //Initialize 
  // load 7z.dll
  NDLL::CLibrary lib;
  if (!lib.Load(wsz7zDllPath))
  {
    OutputDebugStringW(L"Cannot load 7-zip library");
    return ERROR_DLL_NOT_FOUND;
  }
  // get CreateObject
  Func_CreateObject
     f_CreateObject = Z7_GET_PROC_ADDRESS(
  Func_CreateObject, lib.Get_HMODULE(),
      "CreateObject");
  if (!f_CreateObject)
  {
    OutputDebugStringW(L"Cannot get CreateObject");
    return ERROR_PROC_NOT_FOUND;
  }
  // prepare password 
  UString password = wszPassword;
  bool passwordIsDefined = wszPassword[0] != L'\0';
  CMyComPtr<IInArchive> archive;
  if (f_CreateObject(&CLSID_Format, &IID_IInArchive, (void **)&archive) != S_OK)
  {
    OutputDebugStringW(L"Cannot get class object");
    return ERROR_OPERATION_ABORTED;
  }
  
  CInFileStream *fileSpec = new CInFileStream;
  CMyComPtr<IInStream> file = fileSpec;
  
  if (!fileSpec->Open(wszArchiveName))
  {
    OutputDebugStringFormat(L"Cannot open file as archive %s", wszArchiveName);
    return ERROR_FILE_NOT_FOUND;
  }
  {
    CArchiveOpenCallback *openCallbackSpec = new CArchiveOpenCallback;
    CMyComPtr<IArchiveOpenCallback> openCallback(openCallbackSpec);
    openCallbackSpec->PasswordIsDefined = passwordIsDefined;
    openCallbackSpec->Password = password;
    
    const UInt64 scanSize = 1 << 23;
    if (archive->Open(file, &scanSize, openCallback) != S_OK)
    {
      OutputDebugStringFormat(L"Cannot open file as archive %s", wszArchiveName);
      return ERROR_FILE_CORRUPT;
    }
    // Extract command
    CArchiveExtractCallback *extractCallbackSpec = new CArchiveExtractCallback;
    CMyComPtr<IArchiveExtractCallback> extractCallback(extractCallbackSpec);
    extractCallbackSpec->Init(archive, wszOutputFolder); // second parameter is output folder path
    extractCallbackSpec->PasswordIsDefined = passwordIsDefined;
    extractCallbackSpec->Password = password;

    DWORD dwLastError = GetLastError();
    if(funcSetTotal != NULL && funcSetCompleted != NULL)
    {
      extractCallbackSpec->InitCallback(funcSetTotal, funcSetCompleted);
      DWORD dwLastError2 = GetLastError();
      if(dwLastError != dwLastError2)
      {
        OutputDebugStringFormat(L"InitCallback failed %d", dwLastError2);
        return ERROR_INVALID_PARAMETER;
      }
    }

    const wchar_t *names[] =
    {
      L"mt"
    };
    const unsigned kNumProps = sizeof(names) / sizeof(names[0]);
    NCOM::CPropVariant values[kNumProps] =
    {
      (UInt32)nMultithread
    };
    CMyComPtr<ISetProperties> setProperties;
    archive->QueryInterface(IID_ISetProperties, (void **)&setProperties);
    if (setProperties)
    {
      if (setProperties->SetProperties(names, values, kNumProps) != S_OK)
      {
        OutputDebugStringW(L"SetProperties() error");
        return ERROR_BAD_ARGUMENTS;
      }
    }

    HRESULT result = archive->Extract(NULL, (UInt32)(Int32)(-1), false, extractCallback);

    if (result != S_OK)
    {
      OutputDebugStringW(L"Extract Error");
      return ERROR_FUNCTION_FAILED;
    }
  }
  return ERROR_SUCCESS;
}
