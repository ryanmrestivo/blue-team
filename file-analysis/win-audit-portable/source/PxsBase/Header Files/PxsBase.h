///////////////////////////////////////////////////////////////////////////////////////////////////
//
// PXS Base Header
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

#ifndef PXSBASE_PXS_BASE_H_
#define PXSBASE_PXS_BASE_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Compiler directives
///////////////////////////////////////////////////////////////////////////////////////////////////

#ifndef __cplusplus
    #define __cplusplus
#endif

#ifndef UNICODE
    #error UNICODE must be defined
#endif

#ifndef _UNICODE
    #error _UNICODE must be defined
#endif

#if defined ( _MSC_VER ) && ( _MSC_VER < 1600 )
    #error MSVC compiler v16.00 (2010) or newer is required

    #ifndef nullptr
        #define nullptr NULL
    #endif

#endif

// XP
#ifdef __GNUC__
    #define _WIN32_WINNT 0x0501
    #define WINVER       0x0501
#endif

// IE
#ifndef _WIN32_IE
    #define _WIN32_IE 0x0500
#endif

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface

// 2. C System Files
#include <WinSock2.h>   // Winsock2 before Windows
#include <Windows.h>
#include <cfgmgr32.h>
#include <stdint.h>
#include <LM.h>
#include <sql.h>
#include <stdlib.h>
#include <tchar.h>
#include <strsafe.h>    // strsafe.h after tchar.h
#include <WinDNS.h>
#include <Ws2tcpip.h>

#ifdef _MSC_VER
    #include <new.h>
#elif __GNUC__
    #include <new>
#else
    #error Unsupported compiler
#endif

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

// 6. Forwards
class Application;
class ByteArray;
class Exception;
class Font;
class NameValue;
class String;
class StringArray;
class TreeViewItem;
template< class T > class TArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Compilation Messages
///////////////////////////////////////////////////////////////////////////////////////////////////

// Turn on/off complier warnings.
#if defined ( _MSC_VER )

    // Disable code not in-lined for a release build
    #ifdef NDEBUG
        #pragma warning ( disable : 4710 )
    #endif

    // Enable
    #pragma warning ( 4 : 4061 )
    #pragma warning ( 4 : 4062 )
    #pragma warning ( 4 : 4191 )
    #pragma warning ( 4 : 4242 )
    #pragma warning ( 4 : 4254 )
    #pragma warning ( 4 : 4255 )
    #pragma warning ( 4 : 4263 )
    #pragma warning ( 4 : 4264 )
    #pragma warning ( 4 : 4265 )
    #pragma warning ( 4 : 4266 )
    #pragma warning ( 4 : 4287 )
    #pragma warning ( 4 : 4289 )
    #pragma warning ( 4 : 4296 )
    #pragma warning ( 4 : 4302 )
    #pragma warning ( 4 : 4350 )
    #pragma warning ( 4 : 4355 )
    #pragma warning ( 4 : 4365 )
    #pragma warning ( 4 : 4339 )
    #pragma warning ( 4 : 4342 )
    #pragma warning ( 4 : 4347 )
    #pragma warning ( 4 : 4350 )
    #pragma warning ( 4 : 4412 )
    #pragma warning ( 4 : 4431 )
    #pragma warning ( 4 : 4435 )
    #pragma warning ( 4 : 4437 )
    #pragma warning ( 4 : 4536 )
    #pragma warning ( 4 : 4545 )
    #pragma warning ( 4 : 4546 )
    #pragma warning ( 4 : 4547 )
    #pragma warning ( 4 : 4549 )
    #pragma warning ( 4 : 4555 )
    #pragma warning ( 4 : 4557 )
    #pragma warning ( 4 : 4571 )
    #pragma warning ( 4 : 4619 )
    #pragma warning ( 4 : 4623 )
    #pragma warning ( 4 : 4625 )
    #pragma warning ( 4 : 4626 )
    #pragma warning ( 4 : 4628 )
    #pragma warning ( 4 : 4640 )
    #pragma warning ( 4 : 4641 )
    #pragma warning ( 4 : 4682 )
    #pragma warning ( 4 : 4686 )
    #pragma warning ( 4 : 4692 )
    #pragma warning ( 4 : 4738 )
    #pragma warning ( 4 : 4826 )
    #pragma warning ( 4 : 4905 )
    #pragma warning ( 4 : 4906 )
    #pragma warning ( 4 : 4928 )
    #pragma warning ( 4 : 4931 )
    #pragma warning ( 4 : 4946 )
    #pragma warning ( 4 : 4962 )

#endif  //  _MSC_VER

///////////////////////////////////////////////////////////////////////////////////////////////////
// Numerical Constants
///////////////////////////////////////////////////////////////////////////////////////////////////

// -1
const size_t PXS_MINUS_ONE      = SIZE_MAX;

// Limits for 32/64-bit types
#if defined ( _WIN64 )

    #define PXS_SQLLEN_MAX      INT64_MAX
    #define PXS_SQLLEN_MIN      INT64_MIN
    #define PXS_SQLULEN_MAX     UINT64_MAX

#else

    #define PXS_SQLLEN_MAX      LONG_MAX
    #define PXS_SQLLEN_MIN      LONG_MIN
    #define PXS_SQLULEN_MAX     ULONG_MAX

#endif  // _WIN64

// time_t
#if defined ( _MSC_VER ) && ( _MSC_VER >= 1400 )    // MSVC 2005+

    #define PXS_TIME_MAX        INT64_MAX

#else

    #define PXS_TIME_MAX        LONG_MAX

#endif

// Error types
const DWORD PXS_ERROR_TYPE_UNKNOWN          =  0;
const DWORD PXS_ERROR_TYPE_SYSTEM           =  1;
const DWORD PXS_ERROR_TYPE_APPLICATION      =  2;
const DWORD PXS_ERROR_TYPE_NETWORK          =  3;
const DWORD PXS_ERROR_TYPE_COMMON_DIALOG    =  4;
const DWORD PXS_ERROR_TYPE_CONFIG_MANAGER   =  5;
const DWORD PXS_ERROR_TYPE_COM              =  6;
const DWORD PXS_ERROR_TYPE_NTSTATUS         =  7;
const DWORD PXS_ERROR_TYPE_MAPI             =  8;
const DWORD PXS_ERROR_TYPE_EXCEPTION        =  9;
const DWORD PXS_ERROR_TYPE_WMI              = 10;
const DWORD PXS_ERROR_TYPE_ZLIB             = 11;
const DWORD PXS_ERROR_TYPE_DNS              = 12;

// MSDN: Bit 29 is reserved for application-defined error codes
const DWORD PXS_ERROR_BASE                  = 0x20000000;
const DWORD PXS_ERROR_BAD_ARITHMETIC        = PXS_ERROR_BASE + 1;
const DWORD PXS_ERROR_BAD_TYPE_CAST         = PXS_ERROR_BASE + 2;
const DWORD PXS_ERROR_OUT_OF_BOUNDS         = PXS_ERROR_BASE + 3;
const DWORD PXS_ERROR_NULL_POINTER          = PXS_ERROR_BASE + 4;
const DWORD PXS_ERROR_UNHANDLED_EXCEPTION   = PXS_ERROR_BASE + 5;
const DWORD PXS_ERROR_DB_NOT_CONNECTED      = PXS_ERROR_BASE + 6;
const DWORD PXS_ERROR_DB_OPERATION_FAILED   = PXS_ERROR_BASE + 7;

// Severity codes
const DWORD PXS_SEVERITY_ERROR              = 1;
const DWORD PXS_SEVERITY_WARNING            = 2;
const DWORD PXS_SEVERITY_INFORMATION        = 3;

// Log file level
const DWORD PXS_LOG_LEVEL_NONE              = 0;
const DWORD PXS_LOG_LEVEL_ERROR             = 1;
const DWORD PXS_LOG_LEVEL_NORMAL            = 2;
const DWORD PXS_LOG_LEVEL_VERBOSE           = 3;

// State constants
const DWORD PXS_STATE_UNKNOWN               = 0;
const DWORD PXS_STATE_ERROR                 = 1;
const DWORD PXS_STATE_WAITING               = 2;
const DWORD PXS_STATE_EXECUTING             = 3;
const DWORD PXS_STATE_CANCELLED             = 4;
const DWORD PXS_STATE_COMPLETED             = 5;
const DWORD PXS_STATE_CLOSED                = 6;
const DWORD PXS_STATE_RESOLVE_WAIT          = 7;
const DWORD PXS_STATE_RESOLVING             = 8;
const DWORD PXS_STATE_RESOLVED              = 9;
const DWORD PXS_STATE_CONNECT_WAIT          = 10;
const DWORD PXS_STATE_CONNECTING            = 11;
const DWORD PXS_STATE_CONNECTED             = 12;

// Solid 16 colors
const COLORREF PXS_COLOUR_BLACK             = RGB(   0,   0,   0 );
const COLORREF PXS_COLOUR_BLUE              = RGB(   0,   0, 255 );
const COLORREF PXS_COLOUR_CYAN              = RGB(   0, 255, 255 );
const COLORREF PXS_COLOUR_GREY              = RGB( 128, 128, 128 );
const COLORREF PXS_COLOUR_GREEN             = RGB(   0, 255,   0 );
const COLORREF PXS_COLOUR_LITEGREY          = RGB( 192, 192, 192 );
const COLORREF PXS_COLOUR_DARKGREEN         = RGB(   0, 128,   0 );
const COLORREF PXS_COLOUR_MAGENTA           = RGB( 255,   0, 255 );
const COLORREF PXS_COLOUR_MAROON            = RGB( 128,   0,   0 );
const COLORREF PXS_COLOUR_NAVY              = RGB(   0,   0, 128 );
const COLORREF PXS_COLOUR_OLIVE             = RGB( 128, 128,   0 );
const COLORREF PXS_COLOUR_PURPLE            = RGB( 128,   0, 128 );
const COLORREF PXS_COLOUR_RED               = RGB( 255,   0,   0 );
const COLORREF PXS_COLOUR_TEAL              = RGB(   0, 128, 128 );
const COLORREF PXS_COLOUR_WHITE             = RGB( 255, 255, 255 );
const COLORREF PXS_COLOUR_YELLOW            = RGB( 255, 255,   0 );

// Shape constants
const DWORD PXS_SHAPE_NONE                  = 0;
const DWORD PXS_SHAPE_LINE                  = 1;
const DWORD PXS_SHAPE_TRIANGLE_LEFT         = 2;
const DWORD PXS_SHAPE_TRIANGLE_RIGHT        = 3;
const DWORD PXS_SHAPE_RECTANGLE             = 4;
const DWORD PXS_SHAPE_RECTANGLE_DOTTED      = 5;
const DWORD PXS_SHAPE_RECTANGLE_HOLLOW      = 6;
const DWORD PXS_SHAPE_DIAMOND               = 7;
const DWORD PXS_SHAPE_CROSS                 = 8;
const DWORD PXS_SHAPE_FIVE_SIDES            = 9;
const DWORD PXS_SHAPE_CIRCLE                = 10;
const DWORD PXS_SHAPE_ARROW_DOWN            = 11;
const DWORD PXS_SHAPE_ARROW_UP              = 12;
const DWORD PXS_SHAPE_RAISED                = 13;
const DWORD PXS_SHAPE_SUNK                  = 14;
const DWORD PXS_SHAPE_FRAME                 = 15;
const DWORD PXS_SHAPE_FRAME_RTL             = 16;
const DWORD PXS_SHAPE_3D_SUNK               = 17;
const DWORD PXS_SHAPE_3D_RAISED             = 18;
const DWORD PXS_SHAPE_SHADOW                = 19;
const DWORD PXS_SHAPE_TAB                   = 20;
const DWORD PXS_SHAPE_TAB_RTL               = 21;
const DWORD PXS_SHAPE_TAB_DOTTED            = 22;
const DWORD PXS_SHAPE_TAB_DOTTED_RTL        = 23;
const DWORD PXS_SHAPE_MENU_BARS             = 24;
const DWORD PXS_SHAPE_HOME                  = 25;
const DWORD PXS_SHAPE_REFRESH               = 26;

// Layout style constants
const DWORD PXS_LAYOUT_STYLE_NONE           = 0;
const DWORD PXS_LAYOUT_STYLE_ROW_LEFT       = 1;
const DWORD PXS_LAYOUT_STYLE_ROW_MIDDLE     = 2;

// Alignment constants
const DWORD PXS_BOTTOM_ALIGNMENT            = 1;
const DWORD PXS_CENTER_ALIGNMENT            = 2;
const DWORD PXS_LEFT_ALIGNMENT              = 3;
const DWORD PXS_RIGHT_ALIGNMENT             = 4;
const DWORD PXS_TOP_ALIGNMENT               = 5;

// Scroll port line height in pixels
const int PXS_DEFAULT_SCREEN_LINE_HEIGHT    = 16;

// Constants for GetProductInfo. These are the ones not in the SDK used to
// build the project. So if update the SDK or change _WIN32_WINNT will need
// to remove those constants that are already defined.

const DWORD PRODUCT_EDUCATION                           = 0x00000079;
const DWORD PRODUCT_EDUCATION_N                         = 0x0000007A;
const DWORD PRODUCT_ENTERPRISE_S                        = 0x0000007D;
const DWORD PRODUCT_ENTERPRISE_S_EVALUATION             = 0x00000081;
const DWORD PRODUCT_ENTERPRISE_S_N                      = 0x0000007E;
const DWORD PRODUCT_ENTERPRISE_S_N_EVALUATION           = 0x00000082;
const DWORD PRODUCT_IOTUAP                              = 0x0000007B;
const DWORD PRODUCT_MOBILE_ENTERPRISE                   = 0x00000085;


///////////////////////////////////////////////////////////////////////////////////////////////////
// Application messages: Reserve 1-99 for this library
///////////////////////////////////////////////////////////////////////////////////////////////////

const WORD  PXS_APP_MSG_NONE                = WM_APP;
const WORD  PXS_APP_MSG_RTL_READING         = WM_APP + 1;
const WORD  PXS_APP_MSG_SELECT_ALL          = WM_APP + 2;
const WORD  PXS_APP_MSG_COPY_SELECTION      = WM_APP + 3;
const WORD  PXS_APP_MSG_BUTTON_CLICK        = WM_APP + 4;
const WORD  PXS_APP_MSG_LOGGER_UPDATE       = WM_APP + 5;
const WORD  PXS_APP_MSG_SHOW_MENU           = WM_APP + 6;
const WORD  PXS_APP_MSG_COLLAPSE_ALL        = WM_APP + 7;
const WORD  PXS_APP_MSG_EXPAND_ALL          = WM_APP + 8;
const WORD  PXS_APP_MSG_DESELECT_ALL        = WM_APP + 9;
const WORD  PXS_APP_MSG_REMOVE_ITEM         = WM_APP + 10;
const WORD  PXS_APP_MSG_ITEM_SELECTED       = WM_APP + 11;
const WORD  PXS_APP_MSG_SPLITTER            = WM_APP + 12;
const WORD  PXS_APP_MSG_HIDE_WINDOW         = WM_APP + 13;

///////////////////////////////////////////////////////////////////////////////////////////////////
// String Identifiers: Reserve 1-999 for this library
///////////////////////////////////////////////////////////////////////////////////////////////////

// Common Strings                                   // ENGLISH
const DWORD PXS_IDS_100_ERROR_NUMBER        = 100;  // Error Number
const DWORD PXS_IDS_101_ERROR_TYPE          = 101;  // Error Type
const DWORD PXS_IDS_102_UNKNOWN             = 102;  // Unknown
const DWORD PXS_IDS_103_SYSTEM              = 103;  // System
const DWORD PXS_IDS_104_APPLICATION         = 104;  // Application
const DWORD PXS_IDS_105_NETWORK             = 105;  // Network
const DWORD PXS_IDS_106_COMMON_DIALOG       = 106;  // Common Dialog
const DWORD PXS_IDS_107_DEVICE_MANAGER      = 107;  // Device Manager
const DWORD PXS_IDS_108_COM                 = 108;  // COM
const DWORD PXS_IDS_109_NTSTATUS            = 109;  // NTSTATUS
const DWORD PXS_IDS_110_MAPI                = 110;  // MAPI
const DWORD PXS_IDS_111_WMI                 = 111;  // WMI
const DWORD PXS_IDS_112_EXCEPTION           = 112;  // Exception
const DWORD PXS_IDS_113_BAD_ARITHMETIC      = 113;  // An arithmetic overflow or underflow occurred.    [whitespace/line_length] NOLINT
const DWORD PXS_IDS_114_BAD_TYPE_CAST       = 114;  // Failed to convert to a different data type.      [whitespace/line_length] NOLINT
const DWORD PXS_IDS_115_NULL_POINTER        = 115;  // Null or invalid pointer.                         [whitespace/line_length] NOLINT
const DWORD PXS_IDS_116_OUT_OF_BOUNDS       = 116;  // A value is out of bounds.
const DWORD PXS_IDS_117_UNHANDLED_EXCEPTION = 117;  // A fatal error occurred.
const DWORD PXS_IDS_118_MULTI_FATAL_ERRORS  = 118;  // Multiple fatal error, the programme will exit.   [whitespace/line_length] NOLINT
const DWORD PXS_IDS_119_DETAILS             = 119;  // Details
const DWORD PXS_IDS_120_LOCATION            = 120;  // Location
const DWORD PXS_IDS_121_DESCRIPTION         = 121;  // Description
const DWORD PXS_IDS_122_SQL_STATE           = 122;  // SQL State
const DWORD PXS_IDS_123_NATIVE_ERROR        = 123;  // Native Error
const DWORD PXS_IDS_124_ERROR_RECORD_NUMBER = 124;  // Error Record #
const DWORD PXS_IDS_125_DB_NOT_CONNECTED    = 125;  // Not connected to the database.   [whitespace/line_length] NOLINT
const DWORD PXS_IDS_126_DB_OPERATION_FAILED = 126;  // A database operation failed.     [whitespace/line_length] NOLINT
const DWORD PXS_IDS_127_TOO_MANY_ROWS_AFFECT= 127;  // Too many rows affected.
const DWORD PXS_IDS_128_CANCEL              = 128;  // Cancel
const DWORD PXS_IDS_129_COPY                = 129;  // Copy
const DWORD PXS_IDS_130_CLOSE               = 130;  // Close
const DWORD PXS_IDS_131_SELECT_ALL          = 131;  // Select All
const DWORD PXS_IDS_132_EMAIL               = 132;  // E-Mail
const DWORD PXS_IDS_133_EXIT                = 133;  // Exit
const DWORD PXS_IDS_134_BYTES               = 134;  // Bytes
const DWORD PXS_IDS_135_KB                  = 135;  // KB
const DWORD PXS_IDS_136_MB                  = 136;  // MB
const DWORD PXS_IDS_137_GB                  = 137;  // GB
const DWORD PXS_IDS_138_TB                  = 138;  // TB
const DWORD PXS_IDS_139_MINUTE              = 139;  // Minute
const DWORD PXS_IDS_140_MINUTES             = 140;  // Minutes
const DWORD PXS_IDS_141_HOUR                = 141;  // Hour
const DWORD PXS_IDS_142_HOURS               = 142;  // Hours
const DWORD PXS_IDS_143_DAY                 = 143;  // Day
const DWORD PXS_IDS_144_DAYS                = 144;  // Days
const DWORD PXS_IDS_145_STACK               = 145;  // Stack
const DWORD PXS_IDS_146_ZLIB                = 146;  // ZLib
const DWORD PXS_IDS_147_DNS                 = 147;  // DNS

// Find Text Bar
const DWORD PXS_IDS_500_FIND_COLON          = 500;  // Find:                    [whitespace/labels] NOLINT
const DWORD PXS_IDS_501_FIND_BACKWARDS      = 501;  // Find backwards
const DWORD PXS_IDS_502_FIND_FORWARDS       = 502;  // Find forwards
const DWORD PXS_IDS_503_MATCH_CASE          = 503;  // Match case
const DWORD PXS_IDS_504_TEXT_NOT_FOUND      = 504;  // Text not found

// About Dialog
const DWORD PXS_IDS_510_REGISTERED_OWNER    = 510;  // Registered Owner
const DWORD PXS_IDS_511_REGISTERED_ORG      = 511;  // Registered Organization  [whitespace/line_length] NOLINT
const DWORD PXS_IDS_512_PRODUCT_ID          = 512;  // Product ID
const DWORD PXS_IDS_513_VERSION             = 513;  // Version
const DWORD PXS_IDS_514_SIZE                = 514;  // Size
const DWORD PXS_IDS_515_TIMESTAMP           = 515;  // Timestamp
const DWORD PXS_IDS_516_COPYRIGHT_STRING    = 516;  // This computer programme is protected by copyright law. Unauthorized reproduction or distribution may result in civil and ciminal prosecution.    [whitespace/line_length] NOLINT

// Application Error Dialog
const DWORD PXS_IDS_520_APPLICATION_ERROR   = 520;  // Application Error
const DWORD PXS_IDS_521_ERROR_WILL_EXIT     = 521;  // The programme detected an internal error and will exit.  [whitespace/line_length] NOLINT
const DWORD PXS_IDS_522_PRIVACY             = 522;  // Privacy
const DWORD PXS_IDS_523_INFORMATION_HEREIN  = 523;  // No data herein is specific to your computer. This information is used only for correcting the programme error and is held in the strictest possible confidence.  [whitespace/line_length] NOLINT
const DWORD PXS_IDS_524_EMAIL_ERROR_MESSAGE = 524;  // If you have an Internet connection and an e-mail programme, you can help us to improve the quality of our software by informing us of this error. Select 'E-Mail' to send us the error details or 'Exit' to terminate the programme immediately. [whitespace/line_length] NOLINT

// Resource identifiers reserved fro 600 to 601, defined Resources.h

///////////////////////////////////////////////////////////////////////////////////////////////////
// Strings
///////////////////////////////////////////////////////////////////////////////////////////////////

// Useful characters and strings
const wchar_t PXS_CHAR_NULL                 = L'\0';
const wchar_t PXS_CHAR_COLON                = L':';
const wchar_t PXS_CHAR_COMMA                = L',';
const wchar_t PXS_CHAR_CR                   = L'\r';
const wchar_t PXS_CHAR_DOT                  = L'.';
const wchar_t PXS_CHAR_LF                   = L'\n';
const wchar_t PXS_CHAR_TAB                  = L'\t';
const wchar_t PXS_CHAR_QUOTE                = L'"';
const wchar_t PXS_PATH_SEPARATOR            = L'\\';
const wchar_t PXS_CHAR_SPACE                = L' ';
const wchar_t PXS_STRING_CRLF[]             = L"\r\n";
const wchar_t PXS_STRING_DOT[]              = L".";
const wchar_t PXS_STRING_EMPTY[]            = L"";
const wchar_t PXS_STRING_NO[]               = L"No";     // English
const wchar_t PXS_STRING_ONE[]              = L"1";
const wchar_t PXS_STRING_SPACE[]            = L" ";
const wchar_t PXS_STRING_TAB[]              = L"\t";
const wchar_t PXS_STRING_YES[]              = L"Yes";    // English
const wchar_t PXS_STRING_ZERO[]             = L"0";

// Database names, result of SQLGetInfo( SQL_DBMS_NAME )
const wchar_t PXS_DBMS_NAME_ACCESS[]        = L"Access";
const wchar_t PXS_DBMS_NAME_MYSQL[]         = L"MySQL";
const wchar_t PXS_DBMS_NAME_ORACLE[]        = L"Oracle";
const wchar_t PXS_DBMS_NAME_SQL_SERVER[]    = L"Microsoft SQL Server";
const wchar_t PXS_DBMS_NAME_POSTGRE_SQL[]   = L"PostgreSQL";

// Database Keywords, UTC_TIMESTAMP is non-standard
const wchar_t PXS_KEYWORD_CURRENT_TIMESTAMP[]   = L"CURRENT_TIMESTAMP";
const wchar_t PXS_KEYWORD_UPPER[]               = L"UPPER";
const wchar_t PXS_KEYWORD_TIMESTAMP[]           = L"UTC_TIMESTAMP";

// Characters allowed to be used in a computer name see SetComputerName
const wchar_t PXS_STRING_COMPUTER_CHARS[] = L"abcdefghijklmnopqrstuvwxyz" \
                                            L"ABCDEFGHIJGLMNOPQRSTUVWXYZ" \
                                            L"!@#$%^&')(.-_{}~";


///////////////////////////////////////////////////////////////////////////////////////////////////
// Internet
///////////////////////////////////////////////////////////////////////////////////////////////////

// IRI Schemes
const wchar_t PXS_STRING_SCHEME_FILE[]          = L"file";
const wchar_t PXS_STRING_SCHEME_HTTP[]          = L"http";
const wchar_t PXS_STRING_SCHEME_HTTPS[]         = L"https";
const wchar_t PXS_STRING_SCHEME_RESOURCE[]      = L"resource";

// Http
const DWORD   PXS_HTTP_DEFAULT_PORT             = 80;
const DWORD   PXS_HTTPS_DEFAULT_PORT            = 443;
const wchar_t PXS_STRING_HTTP_DEFAULT_PORT[]    = L"80";
const wchar_t PXS_STRING_HTTPS_DEFAULT_PORT[]   = L"443";

// RFC3987
const wchar_t PXS_STRING_RFC3987_GEN_DELIMS[]   = L":/?#[]@";
const wchar_t PXS_STRING_RFC3987_SUB_DELIMS[]   = L"!$&'()*+,;=";
const wchar_t PXS_STRING_RFC3987_UNRESERVED[]   = L"abcdefghijklmnopqrstuvwxyz" \
                                                  L"ABCDEFGHIJKLMNOPQRSTUVWXYZ" \
                                                  L"0123456789-._~";



///////////////////////////////////////////////////////////////////////////////////////////////////
// Structures
///////////////////////////////////////////////////////////////////////////////////////////////////

typedef struct _PXS_TYPE_NUMBER_STRING
{
    size_t  number;
    wchar_t szValue[ 32 ];
} PXS_TYPE_NUMBER_STRING;


typedef struct _PXS_TYPE_STRING_STRING
{
    LPCWSTR  pszString1;
    LPCWSTR  pszString2;
} PXS_TYPE_STRING_STRING;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Function prototypes
///////////////////////////////////////////////////////////////////////////////////////////////////

// EnumSystemFirmwareTables requires XP 64bit and newer
typedef UINT ( WINAPI* LPFN_ENUM_SYSTEM_FIRMWARE_TABLES )
(
    DWORD   FirmwareTableProviderSignature,
    PVOID   pFirmwareTableBuffer,
    DWORD   BufferSize
);

// GetSystemFirmwareTable requires XP 64bit and newer
typedef UINT ( WINAPI* LPFN_GET_SYSTEM_FIRMWARE_TABLE )
(
    DWORD   FirmwareTableProviderSignature,
    DWORD   FirmwareTableID,
    PVOID   pFirmwareTableBuffer,
    DWORD   BufferSize
);

// GetProductInfo requires Vista and newer
typedef BOOL ( WINAPI* LPFN_GET_PRODUCT_INFO )
(
    DWORD   dwOSMajorVersion,
    DWORD   dwOSMinorVersion,
    DWORD   dwSpMajorVersion,
    DWORD   dwSpMinorVersion,
    PDWORD  pdwReturnedProductType
);

///////////////////////////////////////////////////////////////////////////////////////////////////
// Global POD Variables
///////////////////////////////////////////////////////////////////////////////////////////////////

extern Application* g_pApplication;          // Pointer to application object

///////////////////////////////////////////////////////////////////////////////////////////////////
// Global Functions
///////////////////////////////////////////////////////////////////////////////////////////////////

// Handle out of memory for the new operator.
#ifdef _MSC_VER

    int PXSNewHandler( size_t size );

#elif __GNUC__

    void PXSNewHandler() __attribute__ ( ( noreturn ) );

#else

    #error Unsupported compiler

#endif

// Add values
double   PXSAddDouble( double x, double y );
short    PXSAddInt16( short x, short y );
int      PXSAddInt32( int x, int y );
__int64  PXSAddInt64( __int64 x, __int64 y );
long     PXSAddLong( long x, long y );
size_t   PXSAddSizeT( size_t x, size_t y );
size_t   PXSAddSizeT3( size_t x, size_t y, size_t z );
time_t   PXSAddTimeT( time_t x, time_t y );
SQLULEN  PXSAddSqlULen( SQLULEN x, SQLULEN y );
BYTE     PXSAddUInt8( BYTE x, BYTE y );
WORD     PXSAddUInt16( WORD x, WORD y );
DWORD    PXSAddUInt32( DWORD x, DWORD y );
UINT64   PXSAddUInt64( UINT64 x, UINT64 y );
UINT_PTR PXSAddUIntPtr( UINT_PTR x, UINT_PTR y );
WPARAM   PXSAddWParam( WPARAM x, WPARAM y );

// Cast from one type to another
wchar_t PXSCastCharToWChar( char ch );
BYTE    PXSCastInt8ToUInt8( char value );
WORD    PXSCastInt8ToUInt16( char value );
DWORD   PXSCastDoubleToUInt32( double value );
time_t  PXSCastFileTimeToTimeT( const FILETIME& ft );
char    PXSCastInt16ToInt8( short value );
size_t  PXSCastInt16ToSizeT( short value );
USHORT  PXSCastInt16ToUInt16( short value );
size_t  PXSCastInt32ToSizeT( int value );
BYTE    PXSCastInt32ToUInt8( int value );
WORD    PXSCastInt32ToUInt16( int value );
DWORD   PXSCastInt32ToUInt32( int value );
int     PXSCastInt64ToInt32( __int64 value );
BYTE    PXSCastInt64ToUInt8( __int64 value );
DWORD   PXSCastInt64ToUInt32( __int64 value );
size_t  PXSCastInt64ToSizeT( __int64 value );
time_t  PXSCastInt64ToTimeT( __int64 value );
UINT64  PXSCastInt64ToUInt64( __int64 value );
size_t  PXSCastLongToSizeT( long value );
BYTE    PXSCastLongToUInt8( long value );
DWORD   PXSCastLongToUInt32( long value );
size_t  PXSCastLongPtrToSizeT( LONG_PTR longPtr );
WPARAM  PXSCastLResultToWParam( LRESULT value );
size_t  PXSCastPtrDiffToSizeT( ptrdiff_t value );
short   PXSCastSizeTToInt16( size_t value );
int     PXSCastSizeTToInt32( size_t value );
long    PXSCastSizeTToLong( size_t value );
SQLLEN  PXSCastSizeTToSqlLen( size_t value );
SQLULEN PXSCastSizeTToSqlULen( size_t value );
WORD    PXSCastSizeTToUInt16( size_t value );
DWORD   PXSCastSizeTToUInt32( size_t value );
ULONG   PXSCastSizeTToULong( size_t value );
BYTE    PXSCastSqlLenToUInt8( SQLLEN value );
size_t  PXSCastSqlLenToSizeT( SQLLEN value );
size_t  PXSCastSqlULenToSizeT( SQLULEN value );
void    PXSCastTimeTToFileTime( time_t value, FILETIME* pFileTime );
DWORD   PXSCastTimeTToUInt32( time_t value );
int     PXSCastUInt32ToInt32( unsigned int value );
SQLINTEGER PXSCastUInt32ToSqlInteger( DWORD value );
time_t  PXSCastUInt32ToTimeT( unsigned int value );
BYTE    PXSCastUInt32ToUInt8( unsigned int value );
WORD    PXSCastUInt32ToUInt16( unsigned int value );
__int64 PXSCastUInt64ToInt64( UINT64 value );
size_t  PXSCastUInt64ToSizeT( UINT64 value );
DWORD   PXSCastUInt64ToUInt32( UINT64 value );

// Extract Low/High order part of a value
DWORD   PXSWParamToLowUInt32( WPARAM value );

// Limit to a range
void    PXSLimitInt( int lower, int upper, int* pValue );
void    PXSLimitSizeT( size_t lower, size_t upper, size_t* pValue );
void    PXSLimitUInt32( DWORD lower, DWORD upper, DWORD* pValue );

// Logging
bool    PXSIsAppLogging();
void    PXSLogAppError( LPCWSTR pszMessage );
void    PXSLogAppError1( LPCWSTR pszMessage, const String& Insert1 );
void    PXSLogAppError2( LPCWSTR pszMessage, const String& Insert1, const String& Insert2 );
void    PXSLogAppError3( LPCWSTR pszMessage,
                         const String& Insert1, const String& Insert2, const String& Insert3 );
void    PXSLogAppErrorStringUInt32_2( LPCWSTR pszMessage,
                                      const String& Insert1, DWORD value2, DWORD value3 );
void    PXSLogAppErrorUInt32( LPCWSTR pszMessage, DWORD value );
void    PXSLogAppErrorUInt32_2( LPCWSTR pszMessage, DWORD value1, DWORD value2 );
void    PXSLogAppErrorUInt32_3( LPCWSTR pszMessage, DWORD value1, DWORD value2, DWORD value3 );
void    PXSLogAppInfo( LPCWSTR pszMessage );
void    PXSLogAppInfo1( LPCWSTR pszMessage, const String& Insert1 );
void    PXSLogAppInfo2( LPCWSTR pszMessage, const String& Insert1, const String& Insert2 );
void    PXSLogAppInfoStringUInt32( LPCWSTR pszMessage, const String& Insert1, DWORD value2 );
void    PXSLogAppInfoStringUInt32_2( LPCWSTR pszMessage,
                                     const String& Insert1, DWORD value2, DWORD value3 );
void    PXSLogAppInfoUInt32( LPCWSTR pszMessage, DWORD value );
void    PXSLogAppInfoUInt32_2( LPCWSTR pszMessage, DWORD value1, DWORD value2 );
void    PXSLogAppInfoUInt32_3( LPCWSTR pszMessage, DWORD value1, DWORD value2, DWORD value3 );
void    PXSLogAppWarn( LPCWSTR pszMessage );
void    PXSLogAppWarn1( LPCWSTR pszMessage, const String& Insert1 );
void    PXSLogAppWarn2( LPCWSTR pszMessage, const String& Insert1, const String& Insert2 );
void    PXSLogComWarn( HRESULT hResult, LPCWSTR pszMessage );
void    PXSLogComWarn1( HRESULT hResult, LPCWSTR pszMessage, const String& Insert1 );
void    PXSLogConfigManError1( CONFIGRET crResult, LPCWSTR pszMessage, const String& Insert1 );
void    PXSLogDnsError1( DNS_STATUS status, LPCWSTR pszMessage, const String& Insert1 );
void    PXSLogException( const Exception& e, const char* function );
void    PXSLogException( LPCWSTR message, const Exception& e, const char* function );
void    PXSLogNetError1( NET_API_STATUS status, LPCWSTR pszMessage, const String& Insert1 );
void    PXSLogNetError2( NET_API_STATUS status,
                         LPCWSTR pszMessage, const String& Insert1, const String& Insert2 );
void    PXSLogNtStatusError1( NTSTATUS status, LPCWSTR pszMessage, const String& Insert1 );
void    PXSLogNtStatusWarn1( NTSTATUS status, LPCWSTR pszMessage, const String& Insert1 );
void    PXSLogSocketError( int wsaError, LPCWSTR pszMessage );
void    PXSLogSysError( DWORD errorCode, LPCWSTR pszMessage );
void    PXSLogSysError1( DWORD errorCode, LPCWSTR pszMessage, const String& Insert1 );
void    PXSLogSysError2( DWORD errorCode,
                         LPCWSTR pszMessage, const String& Insert1, const String& Insert2 );
void    PXSLogSysInfo( DWORD errorCode, LPCWSTR pszMessage );
void    PXSLogSysInfo1( DWORD errorCode, LPCWSTR pszMessage, const String& Insert1 );
void    PXSLogSysInfo2( DWORD errorCode,
                        LPCWSTR pszMessage, const String& Insert1, const String& Insert2 );
void    PXSLogSysWarn( DWORD errorCode, LPCWSTR pszMessage );
void    PXSLogSysWarn1( DWORD errorCode, LPCWSTR pszMessage, const String& Insert1 );
void    PXSLogSysWarn2( DWORD errorCode,
                        LPCWSTR pszMessage, const String& Insert1, const String& Insert2 );
void    PXSLogWriteEntry( DWORD severity,
                          DWORD errorType,
                          DWORD errorCode,
                          bool translateError,
                          LPCWSTR pszMessage, const String& Insert1, const String& Insert2 );

// Maximum/Minimum
int     PXSMaxInt( int x, int y );
DWORD   PXSMaxUInt32( DWORD x, DWORD y );
size_t  PXSMaxSizeT( size_t x, size_t y );
int     PXSMinInt( int x, int y );
size_t  PXSMinSizeT( size_t x, size_t y );
time_t  PXSMaxTimeT( time_t x, time_t y );
DWORD   PXSMinUInt32( DWORD x, DWORD y );

// Multiplication
double  PXSMultiplyDouble( double x, double y );
int     PXSMultiplyInt32( int x, int y );
__int64 PXSMultiplyInt64( __int64 x, __int64 y );
long    PXSMultiplyLong( long x, long y );
size_t  PXSMultiplySizeT( size_t x, size_t y );
SQLULEN PXSMultiplySqlULen( SQLULEN x, SQLULEN y );
SQLLEN  PXSMultiplySqlLen( SQLLEN x, SQLLEN y );
time_t  PXSMultiplyTimeT( time_t x, time_t y );
BYTE    PXSMultiplyUInt8( BYTE x, BYTE y );
WORD    PXSMultiplyUInt16( WORD x, WORD y );
DWORD   PXSMultiplyUInt32( DWORD x, DWORD y );
UINT64  PXSMultiplyUInt64( UINT64 x, UINT64 y );

// Searching/Sorting
int     PXSQSortStringAscending( const void* pArg1, const void* pArg2 );
int     PXSQSortStringDescending( const void* pArg1, const void* pArg2 );
int     PXSQSortStringStringAscending( const void* pArg1, const void* pArg2 );
void    PXSSortNameValueArray( TArray< NameValue >* pNameValues );

// Abort
void        PXSShowExceptionDialog( const Exception& e, HWND hWndOwner );
LONG WINAPI PXSShowUnhandledExceptionDialog( EXCEPTION_POINTERS* pException );
void        PXSTerminateHandler( void );
LONG WINAPI PXSWriteUnhandledExceptionToLog( EXCEPTION_POINTERS* pException );

// Other
void     PXSBitmapToGreyScale( HBITMAP hBitmap );
void     PXSBitmapRgbToGrb( HBITMAP hBitmap );
BYTE     PXSCharToDigit( wchar_t wch );
void     PXSCoInitializeSecurity();
void     PXSColourRefToHtmlColour( COLORREF colour, String* pHtmlColour );
int      PXSCompareString( LPCWSTR pszString1, LPCWSTR pszString2, bool caseSensitive );
int      PXSCompareStringN( LPCWSTR pszString1,
                            LPCWSTR pszString2, size_t numChars, bool caseSensitive );
int      PXSCompareSystemTime( const SYSTEMTIME* pTime1, const SYSTEMTIME* pTime2 );
size_t   PXSConvertWCharToUtf8( wchar_t wch, char* pBuffer, size_t numChars );
wchar_t* PXSCopyString( const wchar_t* pwzString );
HBITMAP  PXSCreateFilledBitmap( HDC hdc,
                                int width,
                                int height, COLORREF fill, COLORREF border, DWORD shape );
bool     PXSDigitChars4ToInt32( wchar_t wch1,
                                wchar_t wch2, wchar_t wch3, wchar_t wch4, int* pValue );
void     PXSDrawBitmap( HDC hdc, HBITMAP hBitmap, int x, int y );
void     PXSDrawTransparentBitmap( HDC hdc, HBITMAP hBitmap, int x, int y, COLORREF transparent );
char*    PXSDuplicateStringA( const char* pszString );
size_t   PXSEscapeRichTextChar( wchar_t ch, wchar_t* pszBuffer, size_t bufChars );
void     PXSExpandEnvironmentStrings( LPCWSTR pszSource, String* pExpanded );
WORD     PXSFletcherChecksum16( const BYTE* pData , size_t numBytes );
void     PXSGetApplicationName( String* pApplicationName );
void     PXSGetExeDirectory( String* pExeDirectory );
void     PXSGetExePath( String* pExePath );
size_t   PXSGetHtmlCharacterEntity( wchar_t ch, wchar_t* pszBuffer, size_t bufChars );
void     PXSGetMd5Hash( BYTE* pData, DWORD numBytes, String* pMd5Hash );
void     PXSGetResourceString( DWORD resouceID, String* pResourceString );
void     PXSGetResourceString1( DWORD resouceID, const String& Insert1, String* pResourceString );
HFONT    PXSGetStockFont( int stockFont );
void     PXSGetStockFontTextMetrics( HWND hWindow, int stockFont, TEXTMETRIC* lptm );
BYTE     PXSHexDigitsToByte( wchar_t hexdig1, wchar_t hexdig2 );
char     PXSHexDigitsToChar( char hexdig1, char hexdig2 );
void     PXSInitializeComOnThread();
bool     PXSIsAlphaCharacter( wchar_t ch );
bool     PXSIsDigit( wchar_t ch );
bool     PXSIsDigitsOnly( const wchar_t* pString, size_t numChars );
bool     PXSIsHexitA( char ch );
bool     PXSIsHexitW( wchar_t ch );
bool     PXSIsHighColour( HDC hDC );
bool     PXSIsHighContrast();
bool     PXSIsOnlyChar( const wchar_t* pString, size_t numChars, wchar_t wch );
bool     PXSIsInUInt32Array( DWORD value, const DWORD* pArray, size_t numElements );
bool     PXSIsLogicalDriveLetter( wchar_t ch );
bool     PXSIsNanDouble( double value );
bool     PXSIsValidComputerNameChar( wchar_t wch );
bool     PXSIsValidStateCode( DWORD state );
bool     PXSIsValidStructTM( const struct tm& t );
bool     PXSIsValidUtf16( wchar_t wch );
bool     PXSIsUtf8( char* pString, size_t numChars );
bool     PXSIsWhiteSpace( wchar_t wch );
bool     PXSIsWhiteSpaceA( char ch );
void     PXSLoadDataResource( const String& Name, ByteArray* pBuffer );
void     PXSLoadHtmlResource( WORD resourceID, ByteArray* pBuffer );
void     PXSLoadResource( LPCWSTR lpName, LPCWSTR lpType, ByteArray* pBuffer );
int      PXSMakeInt32( BYTE b1, BYTE b2, BYTE b3, BYTE b4 );
WORD     PXSMakeUInt16( BYTE b1, BYTE b2 );
DWORD    PXSMakeUInt32( BYTE b1, BYTE b2, BYTE b3, BYTE b4 );
void     PXSOutputDebugTickCount( LPCWSTR pszString );
bool     PXSPromptForFilePath( HWND hWndOwner,
                               bool open,
                               bool promptOverwrite,
                               LPCWSTR pszFilter,
                               DWORD*  puFilterIndex,
                               LPCWSTR pszFileName, LPCWSTR psExtensions, String* pFilePath );
void     PXSReplaceBitmapColour( HBITMAP hBitmap, COLORREF oldColour, COLORREF newColour );
bool     PXSShowChooseFontDialog( HWND hWndOwner, Font* pFontObject );
DWORD    PXSSwapWords( DWORD value );
void     PXSTranslateDriveType( UINT driveType, String* pTranslation );
void     PXSUnQuoteString( String* pString );
void     PXSVariantToString( const VARIANT* pVariant, String* pValue );
void     PXSVariantValueToString( const VARIANT* pVariant, String* pValue );

// Rich Edit
DWORD CALLBACK PXSRichEditStreamInCallback( DWORD_PTR dwCookie, LPBYTE pbBuff, LONG cb, LONG* pcb );
void  PXSGetRichTextDocumentStart( String* pRichText );

// Open File Name Dialog Hook
UINT_PTR CALLBACK PXSOFNHookProc( HWND hdlg, UINT uiMsg, WPARAM wParam, LPARAM lParam );

#endif  // PXSBASE_PXS_BASE_H_
