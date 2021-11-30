///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Exception Class Implementation
//
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Copyright 1987-2017 PARMAVEX SERVICES
//
// Licensed under the European Union Public Licence (EUPL), Version 1.1 or -
// as soon they will be approved by the European Commission - subsequent
// versions of the EUPL (the "Licence"). You may not use this work except in
// compliance with the Licence. You may obtain a copy of the Licence at:
//
// http://ec.europa.eu/idabc/eupl
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the Licence is distributed on an "AS IS" basis,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the Licence for the specific language governing permissions and
// limitations under the Licence. This source code is free software. It
// must not be sold, leased, rented, sub-licensed or used for any form of
// monetary recompense whatsoever. This notice must not be removed or altered
// from this source distribution.
//
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

// As this is an exception class, any errors raised in public or protected
// methods are suppressed.
//
// Where possible will use system calls for error messages but in many
// instances it is not clear where to get an error translation so
// will use method scope lookup arrays. These are in English.
//

///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/Exception.h"

// 2. C System Files
#include <Dbghelp.h>
#include <mapi.h>
#include <malloc.h>
#include <Wbemidl.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/FileVersion.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/SystemInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
Exception::Exception()
          :m_uErrorCode( 0 ),
           m_uErrorType( 0 ),
           m_CrashString(),
           m_Description(),
           m_Details(),
           m_Message ()
{
}


// Constructor of an exception object with error information
Exception::Exception( DWORD errorType,
                      DWORD errorCode, LPCWSTR pszDetails, const char* pszFunction )
          :m_uErrorCode( errorCode ),
           m_uErrorType( errorType ),
           m_CrashString(),
           m_Description(),
           m_Details(),
           m_Message()
{
    // Fill the description, details then make the message in that order
    try
    {
        FillInDescription();
        FillInDetails( pszDetails, pszFunction, 0 );
        FillInMessage();
    }
    catch ( const Exception& )
    { }     // Ignore
}

// Copy constructor
Exception::Exception( const Exception& oException )
          :m_uErrorCode( 0 ),
           m_uErrorType( 0 ),
           m_CrashString(),
           m_Description(),
           m_Details(),
           m_Message()
{
    Exception();
    *this = oException;
}

// Destructor
Exception::~Exception()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
Exception& Exception::operator= ( const Exception& oException)
{
    if ( this == &oException ) return *this;

    try
    {
        m_uErrorCode  = oException.m_uErrorCode;
        m_uErrorType  = oException.m_uErrorType;
        m_CrashString = oException.m_CrashString;
        m_Description = oException.m_Description;
        m_Details     = oException.m_Details;
        m_Message     = oException.m_Message;
    }
    catch ( const Exception& )
    { }     // Ignore

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the crash details for this exception.
//
//  Parameters:
//      None
//
//  Returns:
//      Constant reference to the crash details message
//===============================================================================================//
const String& Exception::GetCrashString() const
{
    return m_CrashString;
}

//===============================================================================================//
//  Description:
//      Return the error code for this exception
//
//  Parameters:
//      None
//
//  Returns:
//      DWORD representing the error code
//===============================================================================================//
DWORD Exception::GetErrorCode() const
{
    return m_uErrorCode;
}

//===============================================================================================//
//  Description:
//      Return the error type for this exception.
//
//  Parameters:
//      None
//
//  Returns:
//      DWORD representing the error type, e.g. COM, WMI, System etc.
//===============================================================================================//
DWORD Exception::GetErrorType() const
{
    return m_uErrorType;
}

//===============================================================================================//
//
//  Description:
//      Get the the error message, i.e. description, details and stack trace
//
//  Parameters:
//      None
//
//  Returns:
//      Constant reference to the error message
//
//===============================================================================================//
const String& Exception::GetMessage() const
{
    return m_Message;
}

//===============================================================================================//
//  Description:
//      Reset the class scope data
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Exception::Reset()
{
    try
    {
        m_uErrorCode  = 0;
        m_uErrorType  = 0;
        m_CrashString = PXS_STRING_EMPTY;
        m_Description = PXS_STRING_EMPTY;
        m_Details     = PXS_STRING_EMPTY;
        m_Message     = PXS_STRING_EMPTY;
    }
    catch ( const Exception& )
    { }     // Ignore
}

//===============================================================================================//
//  Description:
//        Translate an error code into a string
//
//  Parameters:
//      errorType   - named constant of the error type
//      errorCode   - the error code
//      pTranslation-  receive the translation
//
//  Remarks:
//      There is overlap of the error codes raised by different OS
//      sub-systems so its important to have the type of error
//
//  Returns:
//        true if code was translated otherwise false
//===============================================================================================//
bool Exception::Translate( DWORD errorType, DWORD errorCode, String* pTranslation )
{
    bool      success = false;
    Formatter Format;

    if ( pTranslation == nullptr )
    {
        return false;
    }

    try
    {
        switch( errorType )
        {
            default:
                PXSLogAppError1( L"Unrecognised error type = %%1.", Format.UInt32( errorType ) );
                break;

            case PXS_ERROR_TYPE_SYSTEM:
                success = TranslateSystemError( errorCode, pTranslation );
                break;

            case PXS_ERROR_TYPE_APPLICATION:
                success = TranslateAppError( errorCode, pTranslation );
                break;

            case PXS_ERROR_TYPE_NETWORK:

                // The network message table is in netmsg.dll however
                // network functions sometimes raise system errors
                success = TranslateModuleError( errorCode, L"netmsg.dll", pTranslation );
                if ( success == false )
                {
                    success = TranslateSystemError( errorCode, pTranslation );
                }
                break;

            case PXS_ERROR_TYPE_COMMON_DIALOG:
                success = TranslateCommonDlgError( errorCode, pTranslation );
                break;

            case PXS_ERROR_TYPE_CONFIG_MANAGER:
                success = TranslateCfgMgrError( errorCode, pTranslation );
                break;

            case PXS_ERROR_TYPE_COM:
                success = TranslateComError( static_cast<HRESULT>( errorCode ), pTranslation );
                break;

            case PXS_ERROR_TYPE_NTSTATUS:
                success = TranslateModuleError( errorCode, L"ntdll.dll", pTranslation );
                break;

            case PXS_ERROR_TYPE_MAPI:
                success = TranslateMAPIError( errorCode, pTranslation );
                break;

            case PXS_ERROR_TYPE_EXCEPTION:
                success = TranslateStructuredException( errorCode, pTranslation );
                break;

            case PXS_ERROR_TYPE_WMI:
                // From XP, WMI message table is in wbem\wmiutils.dll
                success = TranslateModuleError( errorCode, L"wbem\\wmiutils.dll", pTranslation );
                break;

            case PXS_ERROR_TYPE_ZLIB:
                success = TranslateZLibError( static_cast<int>( errorCode ), pTranslation );
                break;

            case PXS_ERROR_TYPE_DNS:
                success = TranslateDnsError( static_cast<DNS_STATUS>( errorCode ), pTranslation );
                break;
        }
    }
    catch ( const Exception& )
    { }     // Ignore

    return success;
}

//===============================================================================================//
//  Description:
//      Translate the name of the error type e.g. system, COM etc.
//
//  Parameters:
//      errorType    - error type code
//      pTranslation - receives the error name
//
//  Returns:
//      void
//===============================================================================================//
void Exception::TranslateType( DWORD errorType, String* pTranslation )
{
    size_t    i = 0;
    String    Insert2;
    Formatter Format;

    // Structure to hold error types
    struct _ERROR_TYPE
    {
        DWORD errorType;
        DWORD resourceID;
    } Types[] = {
        { PXS_ERROR_TYPE_UNKNOWN       , PXS_IDS_102_UNKNOWN        },
        { PXS_ERROR_TYPE_SYSTEM        , PXS_IDS_103_SYSTEM         },
        { PXS_ERROR_TYPE_APPLICATION   , PXS_IDS_104_APPLICATION    },
        { PXS_ERROR_TYPE_NETWORK       , PXS_IDS_105_NETWORK        },
        { PXS_ERROR_TYPE_COMMON_DIALOG , PXS_IDS_106_COMMON_DIALOG  },
        { PXS_ERROR_TYPE_CONFIG_MANAGER, PXS_IDS_107_DEVICE_MANAGER },
        { PXS_ERROR_TYPE_COM           , PXS_IDS_108_COM            },
        { PXS_ERROR_TYPE_NTSTATUS      , PXS_IDS_109_NTSTATUS       },
        { PXS_ERROR_TYPE_MAPI          , PXS_IDS_110_MAPI           },
        { PXS_ERROR_TYPE_EXCEPTION     , PXS_IDS_112_EXCEPTION      },
        { PXS_ERROR_TYPE_WMI           , PXS_IDS_111_WMI            },
        { PXS_ERROR_TYPE_ZLIB          , PXS_IDS_146_ZLIB           },
        { PXS_ERROR_TYPE_DNS           , PXS_IDS_147_DNS            } };

    if ( pTranslation == nullptr )
    {
        return;
    }

    // Find it
    try
    {
        for ( i = 0; i < ARRAYSIZE( Types ); i++ )
        {
            if ( errorType == Types[ i ].errorType )
            {
                PXSGetResourceString( Types[ i ].resourceID, pTranslation );
                break;
            }
        }

        // If not found, return the error code
        if ( pTranslation->IsEmpty() )
        {
            *pTranslation = Format.UInt32( errorType );
            Insert2.SetAnsi( __FUNCTION__ );
            PXSLogAppWarn2( L"Unrecognised error type %%1 in '%%2'.", *pTranslation, Insert2 );
        }
    }
    catch ( const Exception& )
    { }     // Ignore
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Format in a class scope string details about a crash
//
//  Parameters:
//      pException - pointer to an EXCEPTION_POINTERS structure, this
//                   provided by the SEH mechanism
//  Remarks:
//  Called by the constructor so must not throw
//
//  Returns:
//      void
//===============================================================================================//
void Exception::FillCrashString( const EXCEPTION_POINTERS* pException )
{
    int     heapStatus = 0;
    String  Translation, StackTrace, ErrorName;
    Formatter Format;

    if ( ( nullptr == pException ) ||
         ( nullptr == pException->ContextRecord   ) ||
         ( nullptr == pException->ExceptionRecord ) ||
         ( nullptr == pException->ExceptionRecord->ExceptionAddress ) )
    {
        return;
    }
    void* pExceptionAddress = pException->ExceptionRecord->ExceptionAddress;

    try
    {
        m_CrashString.Allocate( 2048 );
        m_CrashString = PXS_STRING_EMPTY;

        // Error type
        TranslateType( PXS_ERROR_TYPE_EXCEPTION, &ErrorName );
        m_CrashString = Format.String1( L"Error Type  : %%1\r\n", ErrorName );

        // Error code
        m_CrashString += L"Error Code  : ";
        if ( m_uErrorCode > 0xFFFF )
        {
            m_CrashString += Format.UInt32Hex( m_uErrorCode, true );
        }
        else
        {
            m_CrashString += Format.UInt32( m_uErrorCode );
        }
        m_CrashString += PXS_STRING_CRLF;
        m_CrashString += Format.StringPointer( L"Address     : %%1\r\n", pExceptionAddress );

        // Do a heap check
        heapStatus     = _heapchk();
        m_CrashString += Format.StringInt32( L"Heap Status : %%1\r\n", heapStatus );

        // Add the faulting thread id
        m_CrashString += Format.StringUInt32( L"In Thread ID: %%1\r\n", GetCurrentThreadId() );
        m_CrashString += PXS_STRING_CRLF;
        m_CrashString += L"Registers\r\n---------\r\n";

        // Add the cpu registers, according to WinNT.h
        #if defined ( _M_IX86 )   // 32 bit

            m_CrashString += Format.StringUInt32Hex( L"EBP Register: %%1\r\n",
                                                     pException->ContextRecord->Ebp, true );
            m_CrashString += Format.StringUInt32Hex( L"EDI Register: %%1\r\n",
                                                     pException->ContextRecord->Edi, true );
            m_CrashString += Format.StringUInt32Hex( L"EDP Register: %%1\r\n",
                                                     pException->ContextRecord->Ebp, true );
            m_CrashString += Format.StringUInt32Hex( L"EIP Register: %%1\r\n",
                                                     pException->ContextRecord->Eip, true );
            m_CrashString += Format.StringUInt32Hex( L"ESI Register: %%1\r\n",
                                                     pException->ContextRecord->Esi, true );
            m_CrashString += Format.StringUInt32Hex( L"EAX Register: %%1\r\n",
                                                     pException->ContextRecord->Eax, true );
            m_CrashString += Format.StringUInt32Hex( L"EBX Register: %%1\r\n",
                                                     pException->ContextRecord->Ebx, true );
            m_CrashString += Format.StringUInt32Hex( L"ECX Register: %%1\r\n",
                                                     pException->ContextRecord->Ecx, true );
            m_CrashString += Format.StringUInt32Hex( L"EDX Register: %%1\r\n",
                                                     pException->ContextRecord->Edx, true );

        #elif defined ( _M_X64 )  // AMD and Intel 64-bit

            m_CrashString += Format.StringUInt64Hex( L"RAX Register: %%1\r\n",
                                                     pException->ContextRecord->Rax);
            m_CrashString += Format.StringUInt64Hex( L"RCX Register: %%1\r\n",
                                                     pException->ContextRecord->Rcx);
            m_CrashString += Format.StringUInt64Hex( L"RDX Register: %%1\r\n",
                                                     pException->ContextRecord->Rdx);
            m_CrashString += Format.StringUInt64Hex( L"RBX Register: %%1\r\n",
                                                     pException->ContextRecord->Rbx);
            m_CrashString += Format.StringUInt64Hex( L"RSP Register: %%1\r\n",
                                                     pException->ContextRecord->Rsp);
            m_CrashString += Format.StringUInt64Hex( L"RBP Register: %%1\r\n",
                                                     pException->ContextRecord->Rbp);
            m_CrashString += Format.StringUInt64Hex( L"RSI Register: %%1\r\n",
                                                     pException->ContextRecord->Rsi);
            m_CrashString += Format.StringUInt64Hex( L"RDI Register: %%1\r\n",
                                                     pException->ContextRecord->Rdi);

        #else

            // e.g. _M_IA64
            #error Error, unsupported processor

        #endif

        MakeStackTrace( pException->ContextRecord, true, MAX_STACK_DEPTH, &StackTrace );
        m_CrashString += L"\r\nStack\r\n-----\r\n";
        m_CrashString += StackTrace;
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Store at class scope a descriptive error string from an error
//      code formed from the error type and code.
//
//  Parameters:
//      None
//
//  Remarks:
//      Called by constructor, must not throw any exceptions.
//
//  Returns:
//      void
//===============================================================================================//
void Exception::FillInDescription()
{
    String Description;

    try
    {
        // Translate the error code/type. Make sure translation is terminated
        // and append on a new line. Filter out success code (=0).
        if ( m_uErrorCode )
        {
            Translate( m_uErrorType, m_uErrorCode, &Description );
            Description.Trim();
            if ( ( Description.GetLength() ) &&
                 ( Description.EndsWithCharacter( PXS_CHAR_DOT ) == false ) )
            {
                Description += PXS_CHAR_DOT;
            }
        }
        m_Description = Description;
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//        Fill details about the error.
//
//  Parameters:
//      pszDetails  - an optional string with additional error details
//      pszFunction - an optional string of where the error occurred,
//                    typically this is __FUNCTION__
//      maxDepth    - indicates the maximum depth for a trace, zero for none
//
//  Remarks:
//      Called by constructor, must not throw any exceptions.
//
//  Returns:
//      void
//===============================================================================================//
void Exception::FillInDetails( LPCWSTR pszDetails, const char* pszFunction, DWORD maxDepth )
{
    const size_t EXTRA_TAB_POSITION = 8;
    wchar_t szCode[ 32 ] = { 0 };  // Big enough for an error code
    String  FunctionName, ApplicationName, StackTrace, Text, ErrorName;
    CONTEXT context;

    try
    {
        m_Details.Allocate( 1024 );
        Text.Allocate( 256 );

        // Application
        PXSGetResourceString( PXS_IDS_104_APPLICATION, &Text );
        Text += PXS_CHAR_COLON;
        Text += PXS_STRING_TAB;
        if ( Text.GetLength() <= EXTRA_TAB_POSITION )
        {
            Text += PXS_STRING_TAB;
        }
        m_Details  = Text;
        PXSGetApplicationName( &ApplicationName );
        m_Details += ApplicationName;
        m_Details += PXS_STRING_CRLF;

        // Error type
        PXSGetResourceString( PXS_IDS_101_ERROR_TYPE, &Text );
        Text += PXS_CHAR_COLON;
        if ( Text.GetLength() <= EXTRA_TAB_POSITION )
        {
            Text += PXS_STRING_TAB;
        }
        Text += PXS_STRING_TAB;
        m_Details += Text;
        TranslateType( m_uErrorType, &ErrorName );
        m_Details += ErrorName;
        m_Details += PXS_STRING_CRLF;

        // Error code
        PXSGetResourceString( PXS_IDS_100_ERROR_NUMBER, &Text );
        Text += PXS_CHAR_COLON;
        Text += PXS_STRING_TAB;
        if ( Text.GetLength() <= EXTRA_TAB_POSITION )
        {
            Text += PXS_STRING_TAB;
        }
        m_Details += Text;
        if ( m_uErrorCode > 0xFFFF )
        {
            StringCchPrintf( szCode, ARRAYSIZE( szCode ), L"0x%08lX", m_uErrorCode );
        }
        else
        {
            StringCchPrintf(szCode, ARRAYSIZE( szCode ), L"%lu", m_uErrorCode);
        }
        m_Details += szCode;
        m_Details += PXS_STRING_CRLF;

        // Method/function
        if ( pszFunction )
        {
            PXSGetResourceString( PXS_IDS_120_LOCATION, &Text );
            Text += PXS_CHAR_COLON;
            Text += PXS_STRING_TAB;
            if ( Text.GetLength() <= EXTRA_TAB_POSITION )
            {
                Text += PXS_STRING_TAB;
            }
            m_Details += Text;
            FunctionName.SetAnsi( pszFunction );
            m_Details += FunctionName;
            m_Details += PXS_STRING_CRLF;
        }

        // Details, put last as it may be a long string
        if ( pszDetails && *pszDetails )
        {
            PXSGetResourceString( PXS_IDS_119_DETAILS, &Text );
            Text += PXS_CHAR_COLON;
            Text += PXS_STRING_TAB;
            if ( Text.GetLength() <= EXTRA_TAB_POSITION )
            {
                Text += PXS_STRING_TAB;
            }
            m_Details += Text;
            m_Details += pszDetails;
            m_Details += PXS_STRING_CRLF;
        }

        // Make a small stack trace
        if ( maxDepth )
        {
            memset( &context, 0, sizeof ( context ) );
            context.ContextFlags = CONTEXT_ALL;
            RtlCaptureContext( &context );
            MakeStackTrace( &context, false, maxDepth, &StackTrace );
            if ( StackTrace.GetLength() )
            {
                PXSGetResourceString( PXS_IDS_145_STACK, &Text );
                Text += PXS_CHAR_COLON;
                Text += PXS_STRING_TAB;
                m_Details += Text;
                m_Details += PXS_STRING_CRLF;
                m_Details += StackTrace;
            }
        }
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Fill the class scope error message
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Exception::FillInMessage()
{
    try
    {
        // Start the description
        m_Message.Allocate( 1024 );
        m_Message = m_Description;
        if ( m_Message.GetLength() )
        {
            m_Message += PXS_STRING_CRLF;
        }

        // Add the error details
        if ( m_Details.GetLength() )
        {
            m_Message += PXS_STRING_CRLF;
            m_Message += L"________________________________________________";
            m_Message += PXS_STRING_CRLF;
            m_Message += PXS_STRING_CRLF;
            m_Message += m_Details;
            m_Message += PXS_STRING_CRLF;
            m_Message += PXS_STRING_CRLF;
        }

        // Add the crash details
        if ( m_CrashString.GetLength() )
        {
            m_Message += L"\r\nCrash\r\n-----\r\n";
            m_Message += m_CrashString;
        }
    }
    catch ( const Exception& )
    { }  // Ignore
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the name of a function from the specified address
//
//  Parameters:
//      pAddress     - the address
//      FunctionName - receives the function's name, or ordinal if none
//                     available
//
//  Remarks:
//      Note, symbol functions are single threaded so MSDN says SymInitialize
//      and SymCleanup should be called at system start up
//
//      The MSDN example use an array of ULONG64, but will use a char array
//      because the the name member is CHAR.
//
//  Returns:
//        true on success, else false
//===============================================================================================//
bool Exception::GetFunctionNameFromAddress( const void* pAddress, String* pFunctionName )
{
    const size_t BUF_LEN_CHARS = sizeof ( SYMBOL_INFO ) + MAX_SYM_NAME + 1;
    bool      success = false;
    char      szBuffer[ BUF_LEN_CHARS ] = { 0 };
    Formatter Format;
    HANDLE    hProcess     = nullptr;
    DWORD64   displacement = 0;
    DWORD_PTR addressPtr   = reinterpret_cast<DWORD_PTR>( pAddress );
    SYMBOL_INFO* pSymbol   = reinterpret_cast<SYMBOL_INFO*>( szBuffer );

    // Must have an address and a buffer
    if ( ( pAddress == nullptr ) || ( pFunctionName == nullptr ) )
    {
        return false;
    }
    *pFunctionName = PXS_CHAR_NULL;

    hProcess = GetCurrentProcess();
    if ( SymInitialize( hProcess, nullptr, TRUE ) == 0 )
    {
        return false;
    }

    // Need to clean up
    try
    {
        pSymbol->SizeOfStruct = sizeof ( SYMBOL_INFO );
        pSymbol->MaxNameLen   = MAX_SYM_NAME;
        if ( SymFromAddr( hProcess, static_cast<DWORD64>( addressPtr ), &displacement, pSymbol ) )
        {
            // The SYMBOL_INFO::Name member is at the end of the buffer
            szBuffer[ ARRAYSIZE( szBuffer ) - 1 ] = 0x00;
            pFunctionName->SetAnsi( pSymbol->Name );
            success = true;
        }
    }
    catch ( const Exception& )
    {  }    // Ignore
    SymCleanup( hProcess );

    return success;
}

//===============================================================================================//
//  Description:
//      Make a stack trace string
//
//  Parameters:
//      pContextRecord - CONTEXT structure
//      detailed       - indicates want detailed information
//      maxDepth       - maximum depth
//      pStackTrace    - receives the stack trace
//
//  Remarks:
//      See ACE Stack_Trace.cpp
//
//  Returns:
//      void
//===============================================================================================//
void Exception::MakeStackTrace( const CONTEXT* pContextRecord,
                                bool detailed, DWORD maxDepth, String* pStackTrace )
{
    void*   pAdrress = nullptr;
    DWORD   i = 0, machineType = 0;
    String  Translation;
    HANDLE  hCurrentProcess = GetCurrentProcess();
    HANDLE  hCurrentThread  = GetCurrentThread();
    CONTEXT* pContext       = const_cast<CONTEXT*>( pContextRecord );
    STACKFRAME64 stackFrame;

    if ( ( pContext == nullptr ) || ( pStackTrace == nullptr ) )
    {
        return;
    }
    pStackTrace->Allocate( 2048 );
    *pStackTrace = PXS_STRING_EMPTY;
    memset( &stackFrame, 0, sizeof ( stackFrame ) );

    #if defined ( _M_IX86 )   // 32 bit

        machineType = IMAGE_FILE_MACHINE_I386;
        stackFrame.AddrPC.Offset    = pContextRecord->Eip;
        stackFrame.AddrPC.Mode      = AddrModeFlat;
        stackFrame.AddrFrame.Offset = pContextRecord->Ebp;
        stackFrame.AddrFrame.Mode   = AddrModeFlat;
        stackFrame.AddrStack.Offset = pContextRecord->Esp;
        stackFrame.AddrStack.Mode   = AddrModeFlat;

    #elif defined ( _M_X64 )  // AMD and Intel 64-bit

        machineType = IMAGE_FILE_MACHINE_AMD64;
        stackFrame.AddrPC.Offset    = pContextRecord->Rip;
        stackFrame.AddrPC.Mode      = AddrModeFlat;
        stackFrame.AddrFrame.Offset = pContextRecord->Rbp;
        stackFrame.AddrFrame.Mode   = AddrModeFlat;
        stackFrame.AddrStack.Offset = pContextRecord->Rsp;
        stackFrame.AddrStack.Mode   = AddrModeFlat;

    #else

        #error Unsupported processor

    #endif

    // StackWalk64 with default callbacks
    maxDepth = PXSMinUInt32( maxDepth, MAX_STACK_DEPTH );
    while ( ( i < maxDepth ) && ( StackWalk64( machineType,
                                               hCurrentProcess,
                                               hCurrentThread,
                                               &stackFrame,
                                               reinterpret_cast<void*>( pContext ),
                                               nullptr,
                                               SymFunctionTableAccess64,
                                               SymGetModuleBase64, nullptr) ) )
    {
        pAdrress = reinterpret_cast<void*>( stackFrame.AddrPC.Offset );
        Translation = PXS_STRING_EMPTY;
        TranslateAddressToModuleInformation( pAdrress, detailed, &Translation );
        *pStackTrace += Translation;
        *pStackTrace += PXS_STRING_CRLF;

        i++;
    }
}

//===============================================================================================//
//  Description:
//      Get module information from the specified address
//
//  Parameters:
//      pAddress     - the address
//      detailed     - flag for base and version information
//      pTranslation - receives the translation
//
//  Returns:
//      void
//===============================================================================================//
void Exception::TranslateAddressToModuleInformation( const void* pAddress,
                                                     bool detailed, String* pTranslation )
{
    wchar_t szFilename[ MAX_PATH + 1 ] = { 0 };
    String  Filename, FunctionName, ManufacturerName, VersionString;
    String  Description;
    Formatter    Format;
    FileVersion  FileVer;
    MEMORY_BASIC_INFORMATION  mbi;

    if ( ( pAddress == nullptr ) || ( pTranslation == nullptr ) )
    {
        return;
    }

    *pTranslation = PXS_STRING_EMPTY;
    memset( &mbi, 0, sizeof ( mbi ) );
    if ( VirtualQuery( pAddress, &mbi, sizeof ( mbi ) ) == 0 )
    {
        return;     // Error
    }
    GetModuleFileName( reinterpret_cast<HMODULE>( mbi.AllocationBase ),
                       szFilename, ARRAYSIZE( szFilename ) );
    szFilename[ ARRAYSIZE( szFilename ) - 1 ] = PXS_CHAR_NULL;
    Filename = szFilename;
    GetFunctionNameFromAddress( pAddress, &FunctionName );

    // Version
    ManufacturerName = PXS_STRING_EMPTY;
    VersionString    = PXS_STRING_EMPTY;
    Description      = PXS_STRING_EMPTY;
    FileVer.GetVersion( Filename,
                        &ManufacturerName, &VersionString, &Description );

    // Make a line of text of the form
    // "Caller: ADDRESS Base: ADDRESS module!function [vVer]\r\n"
    if ( detailed )
    {
        *pTranslation += Format.StringPointer( L"Caller: %%1 ", pAddress );
        *pTranslation += Format.StringPointer( L"Base: %%1 ", mbi.AllocationBase );
        *pTranslation += Format.String3( L"%%1!%%2 [v%%3]", Filename, FunctionName, VersionString);
    }
    else
    {
        // Function name only
        *pTranslation += Format.Pointer( pAddress );
        *pTranslation += PXS_CHAR_SPACE;
        *pTranslation += FunctionName;
    }
}

//===============================================================================================//
//  Description:
//      Translate a application error code.
//
//  Parameters:
//      errorCode    - the error code
//      pTranslation - receives the translation
//
//  Returns:
//      true if the code was translated otherwise false
//===============================================================================================//
bool Exception::TranslateAppError( DWORD errorCode, String* pTranslation )
{
    bool   success = false;
    size_t i = 0;

    // Structure to hold error codes and resource identifiers.
    struct _ERROR_MSG
    {
        DWORD  errorCode;
        DWORD  resourceID;
    } Errors[] = {
        { PXS_ERROR_BAD_ARITHMETIC     , PXS_IDS_113_BAD_ARITHMETIC      },
        { PXS_ERROR_BAD_TYPE_CAST      , PXS_IDS_114_BAD_TYPE_CAST       },
        { PXS_ERROR_NULL_POINTER       , PXS_IDS_115_NULL_POINTER        },
        { PXS_ERROR_OUT_OF_BOUNDS      , PXS_IDS_116_OUT_OF_BOUNDS       },
        { PXS_ERROR_UNHANDLED_EXCEPTION, PXS_IDS_117_UNHANDLED_EXCEPTION },
        { PXS_ERROR_DB_NOT_CONNECTED   , PXS_IDS_125_DB_NOT_CONNECTED    },
        { PXS_ERROR_DB_OPERATION_FAILED, PXS_IDS_126_DB_OPERATION_FAILED }
    };

    if ( pTranslation == nullptr )
    {
        return false;
    }

    *pTranslation = PXS_CHAR_NULL;
    for ( i = 0; i < ARRAYSIZE( Errors ); i++ )
    {
        if ( errorCode == Errors[ i ].errorCode )
        {
            PXSGetResourceString( Errors[ i ].resourceID, pTranslation );
            if ( pTranslation->GetLength() )
            {
                success = true;
            }
            break;
        }
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Translate a COM error code (HRESULT).
//
//  Parameters:
//      hResult      - the error code
//      pTranslation - receives the translation
//
//  Returns:
//      true if the code was translates otherwise false
//===============================================================================================//
bool Exception::TranslateComError( HRESULT hResult, String* pTranslation )
{
    bool  success = false;
    DWORD errorCode  = 0;

    if ( pTranslation == nullptr )
    {
        return false;
    }
    *pTranslation = PXS_STRING_EMPTY;

    // Test for a Win32 error
    if ( HRESULT_FACILITY( hResult ) == FACILITY_WIN32 )
    {
        errorCode = static_cast<DWORD>( HRESULT_CODE( hResult ) );
        success = TranslateSystemError( errorCode, pTranslation );
    }

    // Regular error
    if ( success == false )
    {
        success = TranslateSystemError( static_cast<DWORD>( hResult ), pTranslation );
    }

    // If did not get Win32 error it may be a WMI error
    if ( success == false )
    {
        errorCode = static_cast<DWORD>( hResult );
        if ( ( errorCode >= 0x80041000 ) && ( errorCode < 0x80050000 ) )
        {
            // As of XP, the WMI message table is in wbem\wmiutils.dll
            success = TranslateModuleError( errorCode, L"wbem\\wmiutils.dll", pTranslation );
        }
    }

    // If still nothing, try a WBEM status code
    if ( success == false )
    {
        success = TranslateWBEMStatusError( hResult, pTranslation );
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Translate a common dialogue error code.
//
//  Parameters:
//      errorCode    - the error code
//      pTranslation - receives the translation
//
//  Returns:
//      true if the code was translates otherwise false
//===============================================================================================//
bool Exception::TranslateCommonDlgError( DWORD errorCode, String* pTranslation )
{
    bool   success = false;
    size_t i = 0;

    // Structure to hold error messages. There appears to be no documented
    // method to get these error strings so will use the descriptions given in
    // MSDN's CommDlgExtendedError's documentation.
    struct _ERROR_MSG
    {
        DWORD   errorCode;
        LPCWSTR pszError;
    } Errors[] = {
        { CDERR_DIALOGFAILURE   , L"The dialog box could not be created."                        },
        { CDERR_STRUCTSIZE      , L"Invalid value of common dialog box structure size."          },
        { CDERR_INITIALIZATION  , L"The common dialog box function failed during initialization."},
        { CDERR_NOTEMPLATE      , L"Caller failed to provide a corresponding template."          },
        { CDERR_NOHINSTANCE     , L"Caller failed to provided a corresponding instance handle."  },
        { CDERR_LOADSTRFAILURE  , L"The common dialog box failed to load a specified string."    },
        { CDERR_FINDRESFAILURE  , L"The common dialog box failed to find a specified resource."  },
        { CDERR_LOADRESFAILURE  , L"The common dialog box failed to load a specified resource."  },
        { CDERR_LOCKRESFAILURE  , L"The common dialog box failed to lock a specified resource."  },
        { CDERR_MEMALLOCFAILURE , L"The common dialog box was unable to allocate memory for "
                                  L"internal structures."                                        },
        { CDERR_MEMLOCKFAILURE  , L"The common dialog box was unable to lock the memory "
                                  L"associated with a handle."                                   },
        { CDERR_NOHOOK          , L"Caller failed to provide a pointer to a corresponding hook "
                                  L"procedure."                                                  },
        { CDERR_REGISTERMSGFAIL , L"Failed to register a window required by the common dialog "
                                  L"box."                                                        },
        { PDERR_SETUPFAILURE    , L"The print dialog box failed to load the required resources." },
        { PDERR_PARSEFAILURE    , L"The print dialog box failed to parse the strings in the the "
                                  L"WIN.INI file."                                               },
        { PDERR_RETDEFFAILURE   , L"The print dialog box failed because of inconsistent data "
                                  L"supplied by the caller."                                     },
        { PDERR_LOADDRVFAILURE  , L"The print dialog box failed to load the device driver for "
                                  L"the specified printer."                                      },
        { PDERR_GETDEVMODEFAIL  , L"The printer driver failed to initialize a structure."        },
        { PDERR_INITFAILURE     , L"The print dialog box failed during initialization."          },
        { PDERR_NODEVICES       , L"No printer drivers were found."                              },
        { PDERR_NODEFAULTPRN    , L"A default printer does not exist."                           },
        { PDERR_DNDMMISMATCH    , L"The data describes two different printers."                  },
        { PDERR_CREATEICFAILURE , L"The print dialog box failed when it attempted to create an "
                                  L"information context."                                        },
        { PDERR_PRINTERNOTFOUND , L"The WIN.INI file did not contain an entry for the requested "
                                  L"printer."                                                    },
        { PDERR_DEFAULTDIFFERENT, L"Printer described by the other structure members did not "
                                  L"match the current default printer."                          },
        { CFERR_NOFONTS         , L"No fonts exist."                                             },
        { CFERR_MAXLESSTHANMIN  , L"The maximum size specified is less than the minimum size."   },
        { FNERR_SUBCLASSFAILURE , L"List box failed because sufficient memory was not available."},
        { FNERR_INVALIDFILENAME , L"A filename is invalid."                                      },
        { FNERR_BUFFERTOOSMALL  , L"The buffer is too small for the filename."                   },
        { FRERR_BUFFERLENGTHZERO, L"Find-Replace data has an invalid buffer."                    }
    };

    if ( pTranslation == nullptr )
    {
        return false;
    }

    *pTranslation = PXS_CHAR_NULL;
    for ( i = 0; i < ARRAYSIZE( Errors ); i++ )
    {
        if ( errorCode == Errors[ i ].errorCode )
        {
            *pTranslation = Errors[ i ].pszError;
            success = true;
            break;
        }
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Translate a DNS status code.
//
//  Parameters:
//      status      - the status code
//      pTranslation - receives the translation
//
//  Returns:
//      true if the code was translates otherwise false
//===============================================================================================//
bool Exception::TranslateDnsError( DNS_STATUS status, String* pTranslation )
{
    // treat as system error
    return TranslateSystemError( static_cast<DWORD>( status ), pTranslation );
}

//===============================================================================================//
//  Description:
//      Translate a configuration manager error code.
//
//  Parameters:
//      errorCode    - the error code
//      pTranslation - receives the translation
//
//  Returns:
//      true if the code was translates otherwise false
//===============================================================================================//
bool Exception::TranslateCfgMgrError( CONFIGRET errorCode, String* pTranslation)
{
    bool   success = true;
    size_t i = 0;

    // There appears to be no documented method to get these error strings so
    // will use the descriptions given in the Configuration Manager's
    // documentation when given.
    struct _ERROR_MSG
    {
        CONFIGRET errorCode;
        LPCWSTR   pszError;
    } Errors[] = {
        { CR_SUCCESS                 , L"Operation succeeded."                                   },
        { CR_DEFAULT                 , L"Return the default result."                             },
        { CR_OUT_OF_MEMORY           , L"Out of memory."                                         },
        { CR_INVALID_POINTER         , L"Specified pointer is invalid."                          },
        { CR_INVALID_FLAG            , L"Flags parameter is invalid."                            },
        { CR_INVALID_DEVNODE         , L"Device node is invalid."                                },
        { CR_INVALID_RES_DES         , L"Resource descriptor is invalid."                        },
        { CR_INVALID_LOG_CONF        , L"Logical configuration is invalid."                      },
        { CR_INVALID_ARBITRATOR      , L"Arbitrator's registration identifier is invalid or there"
                                       L"is already such a global arbitrator."                   },
        { CR_INVALID_NODELIST        , L"Nodelist header is invalid. Device node already has "
                                       L"requirements."                                          },
        { CR_INVALID_RESOURCEID      , L"The resource identifier is ResType_All."                },
        { CR_DLVXD_NOT_FOUND         , L"Dynamically loadable VxD was not found."                },
        { CR_NO_SUCH_DEVNODE         , L"Device node could not be located."                      },
        { CR_NO_MORE_LOG_CONF        , L"No more logical configurations."                        },
        { CR_NO_MORE_RES_DES         , L"No more resource descriptors."                          },
        { CR_ALREADY_SUCH_DEVNODE    , L"Specified device node already exists."                  },
        { CR_INVALID_RANGE_LIST      , L"Range list not valid."                                  },
        { CR_INVALID_RANGE           , L"End range less than start range."                       },
        { CR_FAILURE                 , L"Failure (Arbitrator/Enumerator)."                       },
        { CR_NO_SUCH_LOGICAL_DEV     , L"Logical device not found in ISAPNP conversion."         },
        { CR_CREATE_BLOCKED          , L"Flags specifies that the devnode should not be created."},
        { CR_NOT_SYSTEM_VM           , L"The service must be coming from the system VM."         },
        { CR_REMOVE_VETOED           , L"The hardware tree should not be removed."               },
        { CR_APM_VETOED              , L"CR_APM_VETOED."                                         },
        { CR_INVALID_LOAD_TYPE       , L"Load type is greater than MAX_DLVXD_LOAD_TYPE."         },
        { CR_BUFFER_SMALL            , L"The operation succeeded, but the specified buffer was "
                                       L"too small."                                             },
        { CR_NO_ARBITRATOR           , L"Resource has no arbitrator."                            },
        { CR_NO_REGISTRY_HANDLE      , L"Operation will not produce registry entry."             },
        { CR_REGISTRY_ERROR          , L"Registry failed."                                       },
        { CR_INVALID_DEVICE_ID       , L"Length of device identifier is greater than "
                                       L"MAX_DEVICE_ID_LEN."                                     },
        { CR_INVALID_DATA            , L"ISA data is invalid."                                   },
        { CR_INVALID_API             , L"The function cannot be called from ring 3."             },
        { CR_DEVLOADER_NOT_READY     , L"Device loader was not ready."                           },
        { CR_NEED_RESTART            , L"The config handler could not dynamically start the "
                                       L"device, but will be able too the next boot."            },
        { CR_NO_MORE_HW_PROFILES     , L"No more hardware profiles."                             },
        { CR_DEVICE_NOT_THERE        , L"Driver could not find device."                          },
        { CR_NO_SUCH_VALUE           , L"Registry value does not exist."                         },
        { CR_WRONG_TYPE              , L"Registry value is of other type."                       },
        { CR_INVALID_PRIORITY        , L"Priority number>MAX_LCPRI."                             },
        { CR_NOT_DISABLEABLE         , L"That devnode cannot be disable right now."              },
        { CR_FREE_RESOURCES          , L"CR_FREE_RESOURCES."                                     },
        { CR_QUERY_VETOED            , L"CR_QUERY_VETOED."                                       },
        { CR_CANT_SHARE_IRQ          , L"CR_CANT_SHARE_IRQ."                                     },
        { CR_NO_DEPENDENT            , L"CR_NO_DEPENDENT."                                       },
        { CR_SAME_RESOURCES          , L"CR_SAME_RESOURCES."                                     },
        { CR_NO_SUCH_REGISTRY_KEY    , L"CR_NO_SUCH_REGISTRY_KEY."                               },
        { CR_INVALID_MACHINENAME     , L"CR_INVALID_MACHINENAME."                                },
        { CR_REMOTE_COMM_FAILURE     , L"CR_REMOTE_COMM_FAILURE."                                },
        { CR_MACHINE_UNAVAILABLE     , L"CR_MACHINE_UNAVAILABLE."                                },
        { CR_NO_CM_SERVICES          , L"CR_NO_CM_SERVICES."                                     },
        { CR_ACCESS_DENIED           , L"CR_ACCESS_DENIED."                                      },
        { CR_CALL_NOT_IMPLEMENTED    , L"CR_CALL_NOT_IMPLEMENTED."                               },
        { CR_INVALID_PROPERTY        , L"CR_INVALID_PROPERTY"                                    },
        { CR_DEVICE_INTERFACE_ACTIVE , L"CR_DEVICE_INTERFACE_ACTIVE."                            },
        { CR_NO_SUCH_DEVICE_INTERFACE, L"CR_NO_SUCH_DEVICE_INTERFACE."                           },
        { CR_INVALID_REFERENCE_STRING, L"CR_INVALID_REFERENCE_STRING."                           },
        { CR_INVALID_CONFLICT_LIST   , L"CR_INVALID_CONFLICT_LIST."                              },
        { CR_INVALID_INDEX           , L"CR_INVALID_INDEX."                                      },
        { CR_INVALID_STRUCTURE_SIZE  , L"CR_INVALID_STRUCTURE_SIZE."                             },
        { NUM_CR_RESULTS             , L"NUM_CR_RESULTS."                                        }
    };

    if ( pTranslation == nullptr )
    {
        return false;
    }

    *pTranslation = PXS_CHAR_NULL;
    for ( i = 0; i < ARRAYSIZE( Errors ); i++ )
    {
        if ( errorCode == Errors[ i ].errorCode )
        {
            *pTranslation = Errors[ i ].pszError;
            success = true;
            break;
        }
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Translate a MAPI error code.
//
//  Parameters:
//      errorCode    - The MAPI error code
//      pTranslation - receives the translation
//
//  Returns:
//      true if the code was translates otherwise false
//===============================================================================================//
bool Exception::TranslateMAPIError( ULONG errorCode, String* pTranslation )
{
    bool   success = false;
    size_t i = 0;

    // Structure to hold error messages. There appears to be no documented
    // method to get these error strings so will use the descriptions given in
    // MAPISendMail's documentation
    struct _ERROR_MSG
    {
        ULONG   errorCode;
        LPCWSTR pszError;
    } Errors[] = {
        { MAPI_E_AMBIGUOUS_RECIPIENT,
          L"A recipient matched more than one of the recipient "
          L"descriptor structures and MAPI_DIALOG was not set."            },
        { MAPI_E_ATTACHMENT_NOT_FOUND,
          L"The specified attachment was not found. "                      },
        { MAPI_E_ATTACHMENT_OPEN_FAILURE,
          L"The specified attachment could not be opened. "  },
        { MAPI_E_BAD_RECIPTYPE,
          L"The type of a recipient was not MAPI_TO, MAPI_CC, or MAPI_BCC."},
        { MAPI_E_FAILURE,
          L"One or more unspecified errors occurred."                      },
        { MAPI_E_INSUFFICIENT_MEMORY,
          L"There was insufficient memory to proceed."                     },
        { MAPI_E_INVALID_RECIPS,
          L"One or more recipients were invalid or did not resolve "
          L"to any address."                                               },
        { MAPI_E_LOGIN_FAILURE,
         L"There was no default logon, and the user failed to log on "
          L"successfully when the logon dialog box was displayed. "        },
        { MAPI_E_TEXT_TOO_LARGE,
          L"The text in the message was too large."                        },
        { MAPI_E_TOO_MANY_FILES,
          L"There were too many file attachments."                         },
        { MAPI_E_TOO_MANY_RECIPIENTS,
          L"case There were too many recipients."                          },
        { MAPI_E_UNKNOWN_RECIPIENT,
          L"A recipient did not appear in the address list."               },
        { MAPI_E_USER_ABORT,
          L"The user cancelled one of the dialog boxes."                   }
    };

    if ( pTranslation == nullptr )
    {
        return false;
    }

    *pTranslation = PXS_CHAR_NULL;
    for ( i = 0; i < ARRAYSIZE( Errors ); i++ )
    {
        if ( errorCode == Errors[ i ].errorCode )
        {
            *pTranslation = Errors[i].pszError;
            success = true;
            break;
        }
    }

    return success;
}

//===============================================================================================//
//  Description:
//        Translate error code using the specified system module
//
//  Parameters:
//      ntStatus         - the error code
//      pszModuleRelPath - the module's path relative to the system directory
//      pTranslation     - receives the translation
//
//  Remarks
//      NTSTATUS is defined as LONG but the error strings are in ntdll.dll
//      which requires using a DWORD with FormatMessage.
//
//  Returns:
//      true if the code was translates otherwise false
//===============================================================================================//
bool Exception::TranslateModuleError( DWORD  errorCode,
                                      LPCWSTR pszModuleRelPath, String* pTranslation )
{
    String      ModulePath;
    Formatter   Format;
    StringArray EmptyArray;
    SystemInformation  SystemInfo;

    if ( pTranslation == nullptr )
    {
        return false;
    }
    SystemInfo.GetSystemDirectoryPath( &ModulePath );
    ModulePath   += pszModuleRelPath;
    *pTranslation = Format.GetModuleString( ModulePath, errorCode, EmptyArray );   // = no inserts
    if ( pTranslation->GetLength() )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Translate a system error.
//
//  Parameters:
//      errorCode    - the error code
//      pTranslation - receives the translation
//
//  Returns:
//      true if the code was translates otherwise false
//===============================================================================================//
bool Exception::TranslateSystemError( DWORD errorCode, String* pTranslation )
{
    Formatter Format;

    if ( pTranslation == nullptr )
    {
        return false;
    }

    *pTranslation = Format.SystemError( errorCode );
    if ( pTranslation->GetLength() )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Translate a structured exception, a.k.a hardware exceptions
//      a.k.a. asynchronous exceptions.
//
//  Parameters:
//      exceptionCode - structured exception code
//      pTranslation  - receives the translation
//
//  Remarks
//      Description will be truncated if it does not entirely fit in the buffer.
//
//  Returns:
//      true if the code was translates otherwise false
//===============================================================================================//
bool Exception::TranslateStructuredException( DWORD exceptionCode, String* pTranslation )
{
    bool   success = false;
    size_t i = 0;
    Formatter Format;

    // Structure to hold error messages. There appears to be no documented
    // method to get these error strings so will use those in MSDN's
    // EXCEPTION_RECORD documentation.
    struct _ERROR_MSG
    {
        DWORD   exceptionCode;
        LPCWSTR pszError;
    } Errors[] = {
        { EXCEPTION_ACCESS_VIOLATION,
          L"The thread tried to read from or write to a virtual address for which it does not "
          L"have the appropriate access (EXCEPTION_ACCESS_VIOLATION)."                           },
        { EXCEPTION_ARRAY_BOUNDS_EXCEEDED,
          L"The thread tried to access an array element that is out of bounds and the underlying "
          L"hardware supports bounds checking (EXCEPTION_ARRAY_BOUNDS_EXCEEDED)."                },
        { EXCEPTION_BREAKPOINT,
          L"A breakpoint was encountered (EXCEPTION_BREAKPOINT)."                                },
        { EXCEPTION_DATATYPE_MISALIGNMENT,
          L"The thread tried to read or write data that is misaligned on hardware that does not "
          L"provide alignment. For example, 16-bit values must be aligned on 2-byte boundaries; "
          L"32-bit values on 4-byte boundaries, and so on (EXCEPTION_DATATYPE_MISALIGNMENT)."    },
        { EXCEPTION_FLT_DENORMAL_OPERAND,
          L"One of the operands in a floating-point operation is denormal. A denormal value is "
          L"one that is too small to represent as a standard floating-point value "
          L"(EXCEPTION_FLT_DENORMAL_OPERAND)."                                                   },
        { EXCEPTION_FLT_DIVIDE_BY_ZERO,
          L"The thread tried to divide a floating-point value by a floating-point divisor of "
          L"zero(EXCEPTION_FLT_DIVIDE_BY_ZERO)"                                                  },
        { EXCEPTION_FLT_INEXACT_RESULT,
          L"The result of a floating-point operation cannot be represented exactly as a decimal "
          L"fraction (EXCEPTION_FLT_INEXACT_RESULT)."                                            },
        { EXCEPTION_FLT_INVALID_OPERATION,
          L"This exception represents any floating-point exception not included in this list "
          L"(EXCEPTION_FLT_INVALID_OPERATION)."                                                  },
        { EXCEPTION_FLT_OVERFLOW,
          L"The exponent of a floating-point operation is greater than the magnitude allowed by "
          L"the corresponding type (EXCEPTION_FLT_OVERFLOW)."                                    },
        { EXCEPTION_FLT_STACK_CHECK,
          L"The stack overflowed or underflowed as the result of a floating-point operation "
          L"(EXCEPTION_FLT_STACK_CHECK)."                                                        },
        { EXCEPTION_FLT_UNDERFLOW,
          L"The exponent of a floating-point operation is less than the magnitude allowed by the "
          L"corresponding type (EXCEPTION_FLT_UNDERFLOW)."                                       },
        { EXCEPTION_ILLEGAL_INSTRUCTION,
          L"The thread tried to execute an invalid instruction (EXCEPTION_ILLEGAL_INSTRUCTION)." },
        { EXCEPTION_IN_PAGE_ERROR,
          L"The thread tried to access a page that was not present, and the system was unable to "
          L"load the page. For example, this exception might occur if a network connection is "
          L"lost while running a program over the network (EXCEPTION_IN_PAGE_ERROR)."            },
        { EXCEPTION_INT_DIVIDE_BY_ZERO,
          L"The thread tried to divide an integer value by an integer divisor of zero "
          L"(EXCEPTION_INT_DIVIDE_BY_ZERO)."                                                     },
        { EXCEPTION_INT_OVERFLOW,
          L"The result of an integer operation caused a carry out of the most significant bit of "
          L"the result (EXCEPTION_INT_OVERFLOW)."                                                },
        { EXCEPTION_INVALID_DISPOSITION,
          L"An exception handler returned an invalid disposition to the exception dispatcher. "
          L"Programmers using a high-level language such as C should never encounter this "
          L"exception (EXCEPTION_INVALID_DISPOSITION)."                                          },
        { EXCEPTION_NONCONTINUABLE_EXCEPTION,
          L"The thread tried to continue execution after a noncontinuable exception occurred "
          L"(EXCEPTION_NONCONTINUABLE_EXCEPTION)."                                               },
        { EXCEPTION_PRIV_INSTRUCTION,
          L"The thread tried to execute an instruction whose operation is not allowed in the "
          L"current machine mode (EXCEPTION_PRIV_INSTRUCTION)."                                  },
        { EXCEPTION_SINGLE_STEP,
          L"A trace trap or other single-instruction mechanism signaled that one instruction has "
          L"been executed (EXCEPTION_SINGLE_STEP)."                                              },
        { EXCEPTION_STACK_OVERFLOW,
          L"The thread used up its stack (EXCEPTION_STACK_OVERFLOW)."                            }
    };

    if ( pTranslation == nullptr )
    {
        return false;
    }

    *pTranslation = PXS_CHAR_NULL;
    for ( i = 0; i < ARRAYSIZE( Errors ); i++ )
    {
        if ( exceptionCode == Errors[ i ].exceptionCode )
        {
            *pTranslation = Errors[ i ].pszError;
            success = true;
            break;
        }
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Translate a WMEM status error
//
//  Parameters:
//      hResult      - the error code
//      pTranslation - receives the translation
//
//  Returns:
//      true on success, else false
//===============================================================================================//
bool Exception::TranslateWBEMStatusError(HRESULT hResult, String* pTranslation )
{
    bool    success = false;
    BSTR    MessageText = nullptr;
    HRESULT hr = 0;
    IWbemStatusCodeText* pIWbemStatusCodeText = nullptr;

    if ( pTranslation == nullptr )
    {
        return false;
    }

    try
    {
        *pTranslation = PXS_STRING_EMPTY;
        hr = CoCreateInstance( CLSID_WbemStatusCodeText,
                               nullptr,
                               CLSCTX_INPROC_SERVER,
                               IID_IWbemStatusCodeText,
                               reinterpret_cast<void**>( &pIWbemStatusCodeText ) );
        if ( FAILED( hr ) || ( pIWbemStatusCodeText == nullptr ) )
        {
            return false;
        }
        hr = pIWbemStatusCodeText->GetErrorCodeText( hResult, 0, 0, &MessageText );
        if ( SUCCEEDED( hr ) )
        {
            *pTranslation = MessageText;
            if ( pTranslation->GetLength() )
            {
                success = true;
            }
        }
    }
    catch ( const Exception& )
    { }     // Ignore
    SysFreeString( MessageText );
    pIWbemStatusCodeText->Release();

    return success;
}

//===============================================================================================//
//  Description:
//      Translate a zlib error.
//
//  Parameters:
//      errorCode    - the error code
//      pTranslation - receives the translation
//
//  Returns:
//      true if the code was translates otherwise false
//===============================================================================================//
bool Exception::TranslateZLibError( int errorCode, String* pTranslation )
{
    if ( pTranslation == nullptr )
    {
        return false;
    }
    pTranslation->Zero();

    switch ( errorCode )
    {
        // Fall through
        default:
        case 0:             // Z_OK = ERROR_SUCCESS
            break;

        case 1:             // Z_STREAM_END = ERROR_HANDLE_EOF
            TranslateSystemError( ERROR_HANDLE_EOF, pTranslation );
            break;

        case 2:             // Z_NEED_DICT
            *pTranslation = L"Need dictionary (Z_NEED_DICT).";
            break;

        case -1:            // Z_ERRNO
            TranslateSystemError( ERROR_IO_DEVICE, pTranslation );
            break;

        case -2:            // Z_STREAM_ERROR
            *pTranslation = L"Invalid compression level (Z_STREAM_ERROR).";
            break;

        case -3:            // Z_DATA_ERROR
            TranslateSystemError( ERROR_INVALID_DATA, pTranslation );
            break;

        case -4:            // Z_MEM_ERROR
            TranslateSystemError( ERROR_OUTOFMEMORY, pTranslation );
            break;

        case -5:            // Z_BUF_ERROR
            TranslateSystemError( ERROR_INSUFFICIENT_BUFFER, pTranslation );
            break;

        case -6:            // Z_VERSION_ERROR
            *pTranslation = L"zlib version mismatch (Z_VERSION_ERROR)";
            break;
    }

    if ( pTranslation->GetLength() )
    {
        return true;
    }

    return false;
}

