///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Global Function Definitions
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

///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/PxsBase.h"

// 2. C System Files
#include <float.h>
#include <stdint.h>
#include <math.h>
#include <ShlObj.h>
#include <WinCrypt.h>

// 3. C++ System Files
#include <cmath>

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/AllocateWChars.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/ArithmeticException.h"
#include "PxsBase/Header Files/BadCastException.h"
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/ByteArray.h"
#include "PxsBase/Header Files/ComException.h"
#include "PxsBase/Header Files/Directory.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/MessageDialog.h"
#include "PxsBase/Header Files/NameValue.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/SystemInformation.h"
#include "PxsBase/Header Files/UnhandledException.h"
#include "PxsBase/Header Files/UnhandledExceptionDialog.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Module POD Variables
///////////////////////////////////////////////////////////////////////////////////////////////////

static LONG m_lExceptionCount = 0;  // Count of fatal exceptions

///////////////////////////////////////////////////////////////////////////////////////////////////
// Global Functions
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Add two double values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      The sum
//===============================================================================================//
double PXSAddDouble( double x, double y )
{
    if ( PXSIsNanDouble( x ) || PXSIsNanDouble( y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    // Adding positive numbers
    if ( ( x > 0 ) &&
         ( y > 0 ) &&
         ( x > ( DBL_MAX - y ) ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

     // Adding negative numbers
    if ( ( x < 0 ) &&
         ( y < 0 ) &&
         ( x < ( -DBL_MAX - y ) ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return ( x + y );
}

//===============================================================================================//
//  Description:
//      Add two signed 16-bit values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      16-bit signed integer
//===============================================================================================//
short PXSAddInt16( short x, short y )
{
    // Adding positive numbers
    if ( ( x > 0 ) &&
         ( y > 0 ) &&
         ( x > ( INT16_MAX - y ) ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

     // Adding negative numbers
    if ( ( x < 0 ) &&
         ( y < 0 ) &&
         ( x < ( INT16_MIN - y ) ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return static_cast<short>( x + y );  // Cast away integer arithmetic
}


//===============================================================================================//
//  Description:
//      Add two signed 32-bit values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      32-bit signed integer
//===============================================================================================//
int PXSAddInt32( int x, int y )
{
    // Adding positive numbers
    if ( ( x > 0 ) &&
         ( y > 0 ) &&
         ( x > ( INT32_MAX - y ) ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

     // Adding negative numbers
    if ( ( x < 0 ) &&
         ( y < 0 ) &&
         ( x < ( INT32_MIN - y ) ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return ( x + y );
}

//===============================================================================================//
//  Description:
//      Add two signed 64-bit values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      64-bit signed integer
//===============================================================================================//
__int64 PXSAddInt64( __int64 x, __int64 y )
{
    // Adding positive numbers
    if ( ( x > 0 ) &&
         ( y > 0 ) &&
         ( x > ( INT64_MAX - y ) ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

     // Adding negative numbers
    if ( ( x < 0 ) &&
         ( y < 0 ) &&
         ( x < ( INT64_MIN - y ) ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return ( x + y );
}

//===============================================================================================//
//  Description:
//      Add two long values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      the long integer
//===============================================================================================//
long PXSAddLong( long x, long y )
{
    // Adding positive numbers
    if ( ( x > 0 ) &&
         ( y > 0 ) &&
         ( x > ( LONG_MAX - y ) ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

     // Adding negative numbers
    if ( ( x < 0 ) &&
         ( y < 0 ) &&
         ( x < ( LONG_MIN - y ) ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return ( x + y );
}

//===============================================================================================//
//  Description:
//      Add two size_t values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      size_t
//===============================================================================================//
size_t PXSAddSizeT( size_t x, size_t y )
{
    if ( x > ( SIZE_MAX - y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return ( x + y );
}

//===============================================================================================//
//  Description:
//      Add three size_t values
//
//  Parameters:
//      x - the first value
//      y - the second value
//      z - the third value
//
//  Returns:
//      size_t
//===============================================================================================//
size_t PXSAddSizeT3( size_t x, size_t y, size_t z )
{
    size_t temp;

    temp = PXSAddSizeT( x   , y );
    return PXSAddSizeT( temp, z );
}


//===============================================================================================//
//  Description:
//      Add two time_t values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      time_t
//===============================================================================================//
time_t PXSAddTimeT( time_t x, time_t y )
{
    if ( x > ( PXS_TIME_MAX - y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return ( x + y );
}

//===============================================================================================//
//  Description:
//      Add two SQLULEN values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      SQLULEN
//===============================================================================================//
SQLULEN PXSAddSqlULen( SQLULEN x, SQLULEN y )
{
    if ( x > ( PXS_SQLULEN_MAX - y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return ( x + y );
}

//===============================================================================================//
//  Description:
//      Add two unsigned 8-bit values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      unsigned 8-bit value
//===============================================================================================//
BYTE PXSAddUInt8( BYTE x, BYTE y )
{
    if ( x > ( UINT8_MAX - y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return static_cast<BYTE>( x + y );  // Cast away integer arithmetic
}

//===============================================================================================//
//  Description:
//      Add two unsigned 16-bit values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      unsigned 16-bit value
//===============================================================================================//
WORD PXSAddUInt16( WORD x, WORD y )
{
    if ( x > ( UINT16_MAX - y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return static_cast<WORD>( x + y );  // Cast away integer arithmetic
}

//===============================================================================================//
//  Description:
//      Add two unsigned 32-bit values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      unsigned 32-bit value
//===============================================================================================//
DWORD PXSAddUInt32( DWORD x, DWORD y )
{
    if ( x > ( DWORD_MAX - y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return ( x + y );
}

//===============================================================================================//
//  Description:
//      Add two unsigned 64-bit values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      UINT64
//===============================================================================================//
UINT64 PXSAddUInt64( UINT64 x, UINT64 y )
{
    if ( x > ( UINT64_MAX - y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return ( x + y );
}

//===============================================================================================//
//  Description:
//      Add two UINT_PTR numbers, WPARAM = UINT_PTR = unsigned 32 or 64-bit
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      UINT_PTR
//===============================================================================================//
UINT_PTR PXSAddUIntPtr( UINT_PTR x, UINT_PTR y )
{
    return PXSAddWParam( x, y );
}

//===============================================================================================//
//  Description:
//      Add two WPARAM numbers, WPARAM = UINT_PTR = unsigned 32 or 64-bit
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      WPARAM
//===============================================================================================//
WPARAM PXSAddWParam( WPARAM x, WPARAM y )
{
    #if defined( _WIN64 )

        return PXSAddUInt64( x, y );

    #else

        return PXSAddUInt32( x, y );

    #endif
}

//===============================================================================================//
//  Description:
//      Convert the specified bitmap to grey scale
//
//  Parameters:
//      hBitmap - handle to the bitmap
//
//  Remarks:
//      Uses GetPixel and SetPixel so use sparingly
//      Grey scale is effected by averaging the RGB value for each pixel
//
//  Returns:
//      void
//===============================================================================================//
void PXSBitmapToGreyScale( HBITMAP hBitmap )
{
    long     i = 0, j = 0;
    HDC      hCompatibleDC;
    BITMAP   bitmap;
    HGDIOBJ  hOldBitmap;
    COLORREF oldPixel  = CLR_INVALID, newPixel = CLR_INVALID;
    COLORREF greyPixel = CLR_INVALID;

    if ( hBitmap == nullptr )
    {
        return;     // Nothing to do
    }

    memset( &bitmap, 0, sizeof ( bitmap ) );
    if ( GetObject( hBitmap, sizeof ( bitmap ), &bitmap ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetObject", __FUNCTION__ );
    }

    hCompatibleDC = CreateCompatibleDC( nullptr );
    if ( hCompatibleDC == nullptr )
    {
        throw SystemException( GetLastError(), L"CreateCompatibleDC", __FUNCTION__ );
    }
    hOldBitmap = SelectObject( hCompatibleDC, hBitmap );

    // Replace each pixel
    for ( i = 0; i < bitmap.bmWidth; i++ )
    {
        for ( j = 0; j < bitmap.bmHeight; j++ )
        {
            oldPixel  = GetPixel( hCompatibleDC, i, j );
            greyPixel = static_cast< COLORREF > ( GetRValue( oldPixel )
                                                + GetGValue( oldPixel )
                                                + GetBValue( oldPixel ) ) / 3;
            newPixel  = static_cast< COLORREF > (  greyPixel |
                                                  (greyPixel <<  8) |
                                                  (greyPixel << 16) );
            SetPixelV( hCompatibleDC, i, j, newPixel );
        }
    }
    SelectObject( hCompatibleDC, hOldBitmap );
    DeleteDC( hCompatibleDC );
}

//===============================================================================================//
//  Description:
//      Cast a double to an unsigned 32-bit integer
//
//  Parameters:
//      value - the value
//
//  Returns:
//      DWORD
//===============================================================================================//
DWORD PXSCastDoubleToUInt32( double value )
{
    if ( ( PXSIsNanDouble( value ) ) ||
         ( value < 0.0 ) ||
         ( value > UINT32_MAX ) )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<DWORD>( value );
}

//===============================================================================================//
//  Description:
//      Cast an signed char to an unsigned wide character
//
//  Parameters:
//      ch - the character
//
//  Returns:
//      unsigned char
//===============================================================================================//
wchar_t PXSCastCharToWChar( char ch )
{
    if ( ch < 0 )
    {
        return static_cast<wchar_t>( 0x100 + ch );
    }

    return static_cast<wchar_t>( ch );
}

//===============================================================================================//
//  Description:
//      Cast the specified FILETIME to a time_t
//
//  Parameters:
//      ft - the file time value
//
//  Remarks:
//      Throws on out of bounds error
//
//  Returns:
//      time_t
//===============================================================================================//
time_t PXSCastFileTimeToTimeT( const FILETIME& ft )
{
    UINT64 time64;

    // see Q167296
    time64 = ft.dwHighDateTime;
    time64 = ( time64 << 32 );
    time64 = PXSAddUInt64( time64, ft.dwLowDateTime );
    time64 = time64 - 116444736000000000LL;
    time64 = time64 / 10000000;

    if ( time64 > PXS_TIME_MAX )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<time_t>( time64 );
}

//===============================================================================================//
//  Description:
//      Cast an 8-bit signed integer to a 8-bit unsigned integer
//
//  Parameters:
//      value - the integer
//
//  Returns:
//      unsigned char
//===============================================================================================//
unsigned char PXSCastInt8ToUInt8( char value )
{
    if ( value < 0 )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<unsigned char>( value );
}

//===============================================================================================//
//  Description:
//      Cast an 8-bit signed integer to an 16-bit unsigned integer
//
//  Parameters:
//      value - the integer
//
//  Returns:
//      WORD
//===============================================================================================//
WORD PXSCastInt8ToUInt16( char value )
{
    if ( value < 0 )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<WORD>( value );
}

//===============================================================================================//
//  Description:
//      Cast a 16-bit signed integer to an 8-bit signed integer
//
//  Parameters:
//      value - the integer
//
//  Returns:
//      char
//===============================================================================================//
char PXSCastInt16ToInt8( short value )
{
    if ( ( value < INT8_MIN ) || ( value > INT8_MAX ) )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<char>( 0xff & value );
}

//===============================================================================================//
//  Description:
//      Cast a 16-bit signed integer to a size_t
//
//  Parameters:
//      value - the integer
//
//  Returns:
//      size_t
//===============================================================================================//
size_t PXSCastInt16ToSizeT( short value )
{
    if ( value < 0 )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<size_t>( value );
}

//===============================================================================================//
//  Description:
//      Cast a 16-bit signed integer to a 16-bit unsigned integer
//
//  Parameters:
//      value - the integer
//
//  Returns:
//      unsigned short
//===============================================================================================//
unsigned short PXSCastInt16ToUInt16( short value )
{
    if ( value < 0 )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<unsigned short>( value );
}

//===============================================================================================//
//  Description:
//      Cast a 32-bit signed integer to a size_t
//
//  Parameters:
//      value - the integer
//
//  Returns:
//      size_t
//===============================================================================================//
size_t PXSCastInt32ToSizeT( int value )
{
    if ( value < 0 )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<size_t>( value );
}

//===============================================================================================//
//  Description:
//      Cast a signed 32-bit integer to an unsigned 8-bit integer
//
//  Parameters:
//      value - the signed 32-bit integer
//
//  Returns:
//      unsigned 8-bit integer
//===============================================================================================//
BYTE PXSCastInt32ToUInt8( int value )
{
    if ( ( value < 0 ) || ( value > UINT8_MAX ) )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<BYTE>( 0xff & value );
}

//===============================================================================================//
//  Description:
//      Cast a signed 32-bit integer to an unsigned 16-bit integer
//
//  Parameters:
//      value - the signed 32-bit integer
//
//  Returns:
//      unsigned 16-bit integer
//===============================================================================================//
WORD PXSCastInt32ToUInt16( int value )
{
    if ( ( value < 0 ) || ( value > UINT16_MAX ) )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<WORD>( 0xffff & value );
}

//===============================================================================================//
//  Description:
//      Cast a signed 32-bit integer to an unsigned 32-bit integer
//
//  Parameters:
//      value - the signed 32-bit integer
//
//  Returns:
//      unsigned 32-bit integer
//===============================================================================================//
DWORD PXSCastInt32ToUInt32( int value )
{
    if ( value < 0 )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<DWORD>( value );
}

//===============================================================================================//
//  Description:
//      Cast a 64-bit signed integer to a 32-bit signed integer
//
//  Parameters:
//      value - the integer
//
//  Returns:
//      BYTE
//===============================================================================================//
int PXSCastInt64ToInt32( __int64 value )
{
    if ( ( value > INT32_MAX ) || ( value < INT32_MIN ) )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<int>( 0xffffffff & value );
}

//===============================================================================================//
//  Description:
//      Cast a 64-bit signed integer to an 8-bit unsigned integer
//
//  Parameters:
//      value - the integer
//
//  Returns:
//      BYTE
//===============================================================================================//
BYTE PXSCastInt64ToUInt8( __int64 value )
{
    if ( ( value < 0 ) || ( value > UINT8_MAX ) )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<BYTE>( 0xff & value );
}

//===============================================================================================//
//  Description:
//      Cast a 64-bit signed integer to a 32-bit unsigned integer
//
//  Parameters:
//      value - the integer
//
//  Returns:
//      DWORD
//===============================================================================================//
DWORD  PXSCastInt64ToUInt32( __int64 value )
{
    if ( ( value < 0 ) || ( value > UINT32_MAX ) )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<DWORD>( 0xffffffff & value );
}

//===============================================================================================//
//  Description:
//      Cast a 64-bit signed integer to a size_t
//
//  Parameters:
//      value - the integer
//
//  Returns:
//      size_t
//===============================================================================================//
size_t PXSCastInt64ToSizeT( __int64 value )
{
    size_t sizeT = 0;

    if ( value < 0 )
    {
        throw BadCastException( __FUNCTION__ );
    }

    #if defined ( _WIN64 )

        sizeT = static_cast<size_t>( value );

    #else

        if ( value > SIZE_MAX )
        {
            throw BadCastException( __FUNCTION__ );
        }
        sizeT = static_cast<size_t>( 0xffffffff & value );

    #endif

    return sizeT;
}

//===============================================================================================//
//  Description:
//      Cast a 64-bit signed integer to a time_t
//
//  Parameters:
//      value - the integer
//
//  Returns:
//      time_t
//===============================================================================================//
time_t PXSCastInt64ToTimeT( __int64 value )
{
    if ( value > PXS_TIME_MAX )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<time_t>( value );
}

//===============================================================================================//
//  Description:
//      Cast a signed 64-bit integer to an unsigned 64-bit integer
//
//  Parameters:
//      value - the integer
//
//  Returns:
//      UINT64
//===============================================================================================//
UINT64 PXSCastInt64ToUInt64( __int64 value )
{
    if ( value < 0 )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<UINT64>( value );
}

//===============================================================================================//
//  Description:
//      Cast a long integer to a size_t integer
//
//  Parameters:
//      value - the long integer
//
//  Remarks:
//      Throws on out of bounds error
//
//  Returns:
//      size_t
//===============================================================================================//
size_t PXSCastLongToSizeT( long value )
{
    if ( value < 0 )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<size_t>( value );
}

//===============================================================================================//
//  Description:
//      Cast a long integer to an unsigned 8-bit integer
//
//  Parameters:
//      value - the long integer
//
//  Remarks:
//      Throws on out of bounds error
//
//  Returns:
//      BYTE
//===============================================================================================//
BYTE PXSCastLongToUInt8( long value )
{
    if ( ( value < 0 ) || ( value > CHAR_MAX ) )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<BYTE>( 0xff & value );
}

//===============================================================================================//
//  Description:
//      Cast a long integer to an unsigned 32-bit integer
//
//  Parameters:
//      value - the long integer
//
//  Returns:
//      DWORD
//===============================================================================================//
DWORD PXSCastLongToUInt32( long value )
{
    if ( value < 0 )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<DWORD>( value );
}

//===============================================================================================//
//  Description:
//      Cast a LONG_PTR to a size_t
//
//  Parameters:
//      value - the LONG_PTR
//
//  Returns:
//      size_t
//===============================================================================================//
size_t  PXSCastLongPtrToSizeT( LONG_PTR value )
{
    #if defined ( _WIN64 )

        return PXSCastInt64ToSizeT( value );

    #else

        return PXSCastLongToSizeT( value );

    #endif
}

//===============================================================================================//
//  Description:
//      Cast an LRESULT to a WPARAM
//
//  Parameters:
//      value - the LRESULT
//
//  Returns:
//      WPARAM
//===============================================================================================//
WPARAM  PXSCastLResultToWParam( LRESULT value )
{
    #if defined ( _WIN64 )

        // __int64 to unsigned __int64
        return PXSCastInt64ToUInt64( value );

    #else

        // long to unsigned int
        return PXSCastLongToUInt32( value );

    #endif
}

//===============================================================================================//
//  Description:
//      Cast a ptrdiff_t to size_t
//
//  Parameters:
//      value - the pointer difference
//
//  Returns:
//      size_t
//===============================================================================================//
size_t PXSCastPtrDiffToSizeT( ptrdiff_t value )
{
    if ( value < 0  )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<size_t>( value );
}

//===============================================================================================//
//  Description:
//      Cast a size_t value to a long integer
//
//  Parameters:
//      value - the size_t integer
//
//  Returns:
//      long
//===============================================================================================//
long PXSCastSizeTToLong( size_t value )
{
    if ( value > LONG_MAX )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<long>( 0xffffffff & value );
}

//===============================================================================================//
//  Description:
//      Cast a size_t value to a signed 32-bit integer
//
//  Parameters:
//      value - the size_t integer
//
//  Returns:
//      int
//===============================================================================================//
int PXSCastSizeTToInt32( size_t value )
{
    if ( value > INT32_MAX )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<int>( 0xffffffff & value );
}

//===============================================================================================//
//  Description:
//      Cast a size_t value to an SQLLEN value
//
//  Parameters:
//      value - the size_t integer
//
//  Returns:
//      SQLLEN value
//===============================================================================================//
SQLLEN PXSCastSizeTToSqlLen( size_t value )
{
    if ( value > PXS_SQLLEN_MAX )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<SQLLEN>( value );
}

//===============================================================================================//
//  Description:
//      Cast a size_t value to an SQLULEN
//
//  Parameters:
//      value - the size_t integer
//
//  Returns:
//      SQLULEN value
//===============================================================================================//
SQLULEN PXSCastSizeTToSqlULen( size_t value )
{
    return static_cast<SQLULEN>( value );
}

//===============================================================================================//
//  Description:
//      Cast a SQLLEN value to a size_t
//
//  Parameters:
//      value - the SQLLEN integer
//
//  Returns:
//      size_t value
//===============================================================================================//
size_t PXSCastSqlLenToSizeT( SQLLEN value )
{
    if ( ( value < 0 ) || ( value > PXS_SQLLEN_MAX ) )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<size_t>( value );
}

//===============================================================================================//
//  Description:
//      Cast a size_t value to an 16-bit unsigned integer
//
//  Parameters:
//      value - the size_t integer
//
//  Returns:
//      16-bit unsigned integer
//===============================================================================================//
WORD PXSCastSizeTToUInt16( size_t value )
{
    if ( value > UINT16_MAX )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<WORD>( 0xffff & value );
}

//===============================================================================================//
//  Description:
//      Cast a size_t value to an 32-bit unsigned integer
//
//  Parameters:
//      value - the size_t integer
//
//  Returns:
//      32-bit unsigned integer
//===============================================================================================//
DWORD PXSCastSizeTToUInt32( size_t value )
{
    if ( value > UINT32_MAX )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<DWORD>( 0xffffffff & value );
}

//===============================================================================================//
//  Description:
//      Cast a size_t value to an unsigned long integer
//
//  Parameters:
//      value - the size_t integer
//
//  Returns:
//      ULONG
//===============================================================================================//
ULONG PXSCastSizeTToULong( size_t value )
{
    if ( value > ULONG_MAX )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<ULONG>( ULONG_MAX & value );
}

//===============================================================================================//
//  Description:
//      Cast a size_t value to 16-bit signed integer
//
//  Parameters:
//      value - the value
//
//  Returns:
//      short
//===============================================================================================//
short PXSCastSizeTToInt16( size_t value )
{
    if ( value > INT16_MAX  )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<short>( 0xffff & value );
}

//===============================================================================================//
//  Description:
//      Cast a SQLLEN value to 8-bit unsigned integer
//
//  Parameters:
//      value - the SQLLEN integer
//
//  Returns:
//      unsigned 8-bit
//===============================================================================================//
BYTE PXSCastSqlLenToUInt8( SQLLEN value )
{
    if ( ( value < 0 ) || ( value > UCHAR_MAX ) )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<BYTE>( 0xff & value );
}

//===============================================================================================//
//  Description:
//      Cast a SQLULEN value to size_t
//
//  Parameters:
//      value - the SQLULEN integer
//
//  Returns:
//      size_t value
//===============================================================================================//
size_t PXSCastSqlULenToSizeT( SQLULEN value )
{
    return static_cast<size_t>( value );
}

//===============================================================================================//
//  Description:
//      Cast the specified time_t to a FILETIME
//
//  Parameters:
//      value     - the value
//      pFileTime - receives the FILETIME
//
//  Returns:
//      void
//===============================================================================================//
void PXSCastTimeTToFileTime( time_t value, FILETIME* pFileTime )
{
     __int64 epoc;

     if ( pFileTime == nullptr )
     {
         throw ParameterException( L"pFileTime", __FUNCTION__ );
     }

    // Convert to a file time, see Q167296
    epoc = PXSMultiplyInt64( value, 10000000 );
    epoc = PXSAddInt64( epoc, 116444736000000000LL );
    pFileTime->dwLowDateTime  = PXSCastInt64ToUInt32( 0xFFFFFFFF & epoc );
    pFileTime->dwHighDateTime = PXSCastInt64ToUInt32( epoc >> 32 );
}

//===============================================================================================//
//  Description:
//      Cast a time_t value to and unsigned 32-bit integer
//
//  Parameters:
//      value - the value
//
//  Remarks:
//      Throws on out of bounds error
//
//  Returns:
//      DWORD
//===============================================================================================//
DWORD PXSCastTimeTToUInt32( time_t value )
{
    if ( value < 0 )
    {
        throw BadCastException( __FUNCTION__ );
    }

    #if defined ( _WIN64 )

        if ( value > DWORD_MAX )
        {
            throw BadCastException( __FUNCTION__ );
        }

    #endif

    return static_cast<DWORD>( value );
}

//===============================================================================================//
//  Description:
//      Cast an unsigned 32-bit integer to a signed 32-bit integer
//
//  Parameters:
//      value - the unsigned 32-bit integer
//
//  Returns:
//      int
//===============================================================================================//
int PXSCastUInt32ToInt32( unsigned int value )
{
    if ( value > INT32_MAX )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<int>( value );
}

//===============================================================================================//
//  Description:
//      Cast an unsigned 32-bit integer to an SQLINTEGER
//
//  Parameters:
//      value - the unsigned 32-bit integer
//
//  Returns:
//      time_t
//===============================================================================================//
SQLINTEGER PXSCastUInt32ToSqlInteger( DWORD value )
{
    // SQLINTEGER = long
    if ( value > LONG_MAX )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<SQLINTEGER>( value );
}

//===============================================================================================//
//  Description:
//      Cast an unsigned 32-bit integer to a time_t integer
//
//  Parameters:
//      value - the unsigned 32-bit integer
//
//  Returns:
//      time_t
//===============================================================================================//
time_t PXSCastUInt32ToTimeT( unsigned int value )
{
    if ( value > PXS_TIME_MAX )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<time_t>( value );
}

//===============================================================================================//
//  Description:
//      Cast an unsigned 32-bit integer to an unsigned 8-bit integer
//
//  Parameters:
//      value - the unsigned 32-bit integer
//
//  Returns:
//      BYTE
//===============================================================================================//
BYTE PXSCastUInt32ToUInt8( unsigned int value )
{
    if ( value > UINT8_MAX )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<BYTE>( 0xff & value );
}

//===============================================================================================//
//  Description:
//      Cast an unsigned 32-bit integer to an unsigned 16-bit integer
//
//  Parameters:
//      value - the unsigned 32-bit integer
//
//  Returns:
//      WORD
//===============================================================================================//
WORD PXSCastUInt32ToUInt16( unsigned int value )
{
    if ( value > UINT16_MAX )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<WORD>( 0xffff & value );
}

//===============================================================================================//
//  Description:
//      Cast an unsigned 64-bit integer to a signed 64-bit integer
//
//  Parameters:
//      value - the unsigned 64-bit integer
//
//  Returns:
//      signed 64-bit integer
//===============================================================================================//
__int64 PXSCastUInt64ToInt64( unsigned __int64 value )
{
    if ( value > INT64_MAX )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<__int64>( value );
}

//===============================================================================================//
//  Description:
//      Cast an unsigned 64-bit integer to size_t integer
//
//  Parameters:
//      value - the unsigned 64-bit integer
//
//  Returns:
//      size_t
//===============================================================================================//
size_t PXSCastUInt64ToSizeT( UINT64 value )
{
    #if defined ( _WIN64 )

        return static_cast<size_t>( value );

    #else

        if ( value > SIZE_MAX )
        {
            throw BadCastException( __FUNCTION__ );
        }
        return static_cast<size_t>( value );

    #endif
}

//===============================================================================================//
//  Description:
//      Cast an unsigned 64-bit integer to an unsigned 32-bit integer
//
//  Parameters:
//      value - the unsigned 64-bit integer
//
//  Returns:
//      unsigned 32-bit integer
//===============================================================================================//
DWORD PXSCastUInt64ToUInt32( unsigned __int64 value )
{
    if ( value > UINT32_MAX )
    {
        throw BadCastException( __FUNCTION__ );
    }

    return static_cast<DWORD>( 0xffffffff & value );
}

//===============================================================================================//
//  Description:
//      Convert the input character to a digit, i.e. "0" is zero
//
//  Parameters:
//      wch
//
//  Returns:
//      The numerical value of the digit, between 0 and 9
//===============================================================================================//
BYTE PXSCharToDigit( wchar_t wch )
{
    if ( PXSIsDigit( wch ) == false )
    {
        throw BoundsException( L"wch", __FUNCTION__ );
    }

    return PXSCastInt32ToUInt8( wch - '0' );
}

//===============================================================================================//
//  Description:
//      Convert the four input characters to a number,e.g. "1994" = 1994 or "0094" = 94
//
//  Parameters:
//      wch1   - the most significant digit
//      wch2   - the most significant digit
//      wch3   - the most significant digit
//      wch4   - the least significant digit
//      pValue - receives the numberic value
//
//  Returns:
//      true if able to convert otherwise false
//===============================================================================================//
bool PXSDigitChars4ToInt32( wchar_t wch1, wchar_t wch2, wchar_t wch3, wchar_t wch4, int* pValue )
{
    if ( pValue == nullptr )
    {
        throw NullException( L"pValue", __FUNCTION__  );
    }

    if ( PXSIsDigit( wch1 ) && PXSIsDigit( wch2 ) &&
         PXSIsDigit( wch3 ) && PXSIsDigit( wch4 )  )
    {
        *pValue = ( 1000 * PXSCharToDigit( wch1 ) ) +
                  (  100 * PXSCharToDigit( wch2 ) ) +
                  (   10 * PXSCharToDigit( wch3 ) ) + PXSCharToDigit( wch4 );
    }
    else
    {
        return false;
    }

    return true;
}


//===============================================================================================//
//  Description:
//      Initialise COM security at the process level
//
//  Parameters:
//      None
//
//  Remarks:
//      Will use RPC_C_AUTHN_LEVEL_DEFAULT  and RPC_C_IMP_LEVEL_IMPERSONATE
//      authentication and impersonation levels as subsequent calls to
//      CoSetProxyBlanket may fail if use different values.
//
//  Returns:
//      void
//===============================================================================================//
void PXSCoInitializeSecurity()
{
     HRESULT hResult = CoInitializeSecurity( nullptr,
                                             -1,
                                             nullptr,
                                             nullptr,
                                             RPC_C_AUTHN_LEVEL_DEFAULT,
                                             RPC_C_IMP_LEVEL_IMPERSONATE,
                                             nullptr, EOAC_NONE, nullptr );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CoInitializeSecurity", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Convert a COLORREF to a html colour, e.g. #ffcc99
//
//  Parameters:
//      colour      - colour to convert
//      pHtmlColour - string to receive the html colour
//
//  Returns:
//      void
//===============================================================================================//
void PXSColourRefToHtmlColour( COLORREF colour, String* pHtmlColour )
{
    wchar_t szHtmlColour[ 8 ] = { 0 };    // Max is 7 characters

    if ( pHtmlColour == nullptr )
    {
        throw ParameterException( L"pHtmlColour", __FUNCTION__ );
    }
    *pHtmlColour = PXS_STRING_EMPTY;

    if ( colour == CLR_INVALID )
    {
        return;
    }

    szHtmlColour[ 0 ] = '#';
    StringCchPrintf( szHtmlColour + 1,
                     ARRAYSIZE( szHtmlColour ) - 1,
                     L"%02x%02x%02x",
                     GetRValue( colour ),
                     GetGValue( colour ),
                     GetBValue( colour ) );
    *pHtmlColour = szHtmlColour;
}

//===============================================================================================//
//  Description:
//      Compare two strings
//
//  Parameters:
//      pszString1    - string one
//      pszString2    - string two
//      caseSensitive - indicates if its case sensitive
//
//  Returns:
//      -ve if pszString1 < pszString2
//        0 if pszString1 = pszString2 ( or NULL == NULL )
//      +ve if pszString2 > pszString2
//===============================================================================================//
int PXSCompareString( LPCWSTR pszString1, LPCWSTR pszString2, bool caseSensitive )
{
    return PXSCompareStringN( pszString1, pszString2, PXS_MINUS_ONE, caseSensitive );
}

//===============================================================================================//
//  Description:
//      Compare two strings up to the specified number of characters
//
//  Parameters:
//      pszString1    - string one
//      pszString2    - string two
//      numChars      - number of characters to compare
//      caseSensitive - indicates if its case sensitive
//
//  Returns:
//      -ve if pszString1 < pszString2
//        0 if pszString1 = pszString2 ( or NULL == NULL )
//      +ve if pszString2 > pszString2
//===============================================================================================//
int PXSCompareStringN( LPCWSTR pszString1, 
                       LPCWSTR pszString2, size_t numChars, bool caseSensitive )
{
    int   result     = 0, cchCount = 0;
    DWORD dwCmpFlags = 0;  // = Case sensitive

    if ( pszString1 && pszString2 )
    {
        if ( caseSensitive == false )
        {
            dwCmpFlags = NORM_IGNORECASE;
        }

        // CompareString wants an int
        if ( numChars == PXS_MINUS_ONE )
        {
            cchCount = -1;
        }
        else
        {
            cchCount = PXSCastSizeTToInt32( numChars );
        }
        result = CompareString( LOCALE_INVARIANT,
                                dwCmpFlags, pszString1, cchCount, pszString2, cchCount );
        if ( result )
        {
            result -= 2;
        }
    }
    else if ( pszString1 && !pszString2 )
    {
        result = 1;
    }
    else if ( !pszString1 && pszString2 )
    {
        result = -1;
    }

    return result;
}

//===============================================================================================//
//  Description:
//      Compare two SYSTEMTIME structures
//
//  Parameters:
//      pTime1 - system time one
//      pTime2 - system time two
//
//  Remarks:
//      Does not check the structures for valid values
//
//  Returns:
//      -1 if pTime1 < pTime2
//       0 if pTime1 = pTime2
//      +1 if pTime1 > pTime2
//===============================================================================================//
int PXSCompareSystemTime( const SYSTEMTIME* pTime1, const SYSTEMTIME* pTime2 )
{
    if ( ( pTime1 == nullptr ) || ( pTime2 == nullptr ) )
    {
        throw ParameterException( L"pTime1/pTime2", __FUNCTION__ );
    }

    // Year
    if ( pTime1->wYear > pTime2->wYear )
    {
        return 1;
    }

    if ( pTime1->wYear < pTime2->wYear )
    {
        return -1;
    }

    // Month
    if ( pTime1->wMonth > pTime2->wMonth )
    {
        return 1;
    }

    if ( pTime1->wMonth < pTime2->wMonth )
    {
        return -1;
    }

    // Day
    if ( pTime1->wDay > pTime2->wDay )
    {
        return 1;
    }

    if ( pTime1->wDay < pTime2->wDay )
    {
        return -1;
    }

    // Hour
    if ( pTime1->wHour > pTime2->wHour )
    {
        return 1;
    }

    if ( pTime1->wHour < pTime2->wHour )
    {
        return -1;
    }

    // wMinute
    if ( pTime1->wMinute > pTime2->wMinute )
    {
        return 1;
    }

    if ( pTime1->wMinute < pTime2->wMinute )
    {
        return -1;
    }

    // wSecond
    if ( pTime1->wSecond > pTime2->wSecond )
    {
        return 1;
    }

    if ( pTime1->wSecond < pTime2->wSecond )
    {
        return -1;
    }

    // wMilliseconds
    if ( pTime1->wMilliseconds > pTime2->wMilliseconds )
    {
        return 1;
    }

    if ( pTime1->wMilliseconds < pTime2->wMilliseconds )
    {
        return -1;
    }

    return 0;
}

//===============================================================================================//
//  Description:
//      Covert the specified character to UTF-8
//
//  Parameters:
//      wch      - teh character to convert
//      pBuffer  - recieves the UTF-8 bytes
//      numChars - the size of the buffer
//
//  Returns:
//      number of bytes put into pBuffer
//===============================================================================================//
size_t PXSConvertWCharToUtf8( wchar_t wch, char* pBuffer, size_t numChars )
{
    int result, cchMultiByte;

    if ( ( pBuffer == nullptr ) || ( numChars == 0 ) )
    {
        throw ParameterException( L"pBuffer/numChars", __FUNCTION__ );
    }
    memset( pBuffer, 0, numChars );

    if ( wch < 128 )
    {
        pBuffer[ 0 ] = static_cast< char >( wch );
        return 1;
    }

    // For the code page 65001 (UTF-8), dwFlags must be set to either 0 or WC_ERR_INVALID_CHARS.
    // Otherwise, the function fails with ERROR_INVALID_FLAGS.
    // Because the length of the input is specified, i.e. 1, result does not include NULL terminator
    // so the returned value is the actual number of UTF-8 bytes to encode the wide char.
    cchMultiByte = PXSCastSizeTToInt32( numChars );
    result       = WideCharToMultiByte( CP_UTF8,
                                        0, &wch, 1, pBuffer, cchMultiByte, nullptr, nullptr );
    if ( result == 0 )
    {
        throw SystemException( GetLastError(), L"WideCharToMultiByte", __FUNCTION__ );
    }

    return PXSCastInt32ToSizeT( result );
}

//===============================================================================================//
//  Description:
//      Make a copy the input string. The caller is responsible for calling delete[] the free
//      the allocated memory 
//
//  Parameters:
//      pwzString - the string to copy
//
//  Returns:
//      non-const pointer to the new string. NULL if the input is NULL
//===============================================================================================//
wchar_t* PXSCopyString( const wchar_t* pwzString )
{
    size_t   length;
    wchar_t* pwzCopy;

    if ( pwzString == nullptr )
    {
        return nullptr;
    }
    length  = wcslen( pwzString );
    length  = PXSAddSizeT( length, 1 );  // Terminator
    pwzCopy = new wchar_t[ length ];
    if ( pwzCopy == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    StringCchCopy( pwzCopy, length, pwzString );

    return pwzCopy;
}

//===============================================================================================//
//  Description:
//      Create a bitmap with the specified border and fill colours
//
//  Parameters:
//      hdc    - handle to device context for the bitmap
//      width  - width of bitmap in pixels
//      height - height  of bitmap in pixels
//      fill   - the colour to fill the interior of the border
//      border - the border colour, 1 pixel wide
//      shape  - defined shape constant
//
//  Returns:
//      void
//===============================================================================================//
HBITMAP PXSCreateFilledBitmap( HDC hdc,
                               int width,
                               int height, COLORREF fill, COLORREF border, DWORD shape )
{
    HDC     compatibleDC;
    RECT    bounds    = { 0, 0, width, height };
    HBITMAP hBitmap   = nullptr;
    HGDIOBJ oldBitmap = nullptr;
    StaticControl Static;

    if ( hdc == nullptr )
    {
        throw ParameterException( L"hdc", __FUNCTION__ );
    }

    hBitmap = CreateCompatibleBitmap( hdc, width, height );
    if ( hBitmap == nullptr )
    {
        throw SystemException( GetLastError(), L"CreateCompatibleBitmap", __FUNCTION__ );
    }

    compatibleDC = CreateCompatibleDC( hdc );
    if ( compatibleDC == nullptr )
    {
        DeleteObject( hBitmap );
        throw SystemException( GetLastError(), L"CreateCompatibleDC", __FUNCTION__ );
    }
    oldBitmap = SelectObject( compatibleDC, hBitmap );

    try
    {
        Static.SetBounds( bounds );
        Static.SetBackground( PXS_COLOUR_WHITE );
        Static.Draw( compatibleDC );

        Static.SetBackground( fill );
        Static.SetBounds( bounds );
        Static.SetShape( shape );
        Static.SetShapeColour( border );
        Static.Draw( compatibleDC );
    }
    catch ( const Exception& )
    {
        if ( oldBitmap )
        {
            SelectObject( compatibleDC, oldBitmap );
        }
        DeleteObject( hBitmap );
        DeleteDC( compatibleDC );
        throw;
    }

    // Reset
    if ( oldBitmap )
    {
        SelectObject( compatibleDC, oldBitmap );
    }
    DeleteDC( compatibleDC );

    return hBitmap;
}

//===============================================================================================//
//  Description:
//      Draw a bitmap on the specified device context
//
//  Parameters:
//      hdc     - handle to device context to draw on
//      hBitmap - handle to bitmap
//      xPos    - x position of bitmap
//      yPos    - y position of bitmap
//
//  Returns:
//      void
//===============================================================================================//
void PXSDrawBitmap( HDC hdc, HBITMAP hBitmap, int x, int y )
{
    HDC     compatibleDC;
    BITMAP  bitmap;
    HGDIOBJ oldBitmap;

    if ( ( hdc == nullptr ) || ( hBitmap == nullptr ) )
    {
        return;     // Nothing to do
    }

    // Get the bitmap
    memset( &bitmap, 0, sizeof ( bitmap ) );
    if ( GetObject( hBitmap, sizeof ( bitmap ), &bitmap ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetObject", __FUNCTION__ );
    }

    // Select the bitmap into a compatible device context
    compatibleDC = CreateCompatibleDC( hdc );
    if ( compatibleDC == nullptr )
    {
        throw SystemException( GetLastError(), L"CreateCompatibleDC", __FUNCTION__ );
    }
    oldBitmap = SelectObject( compatibleDC, hBitmap );
    if ( BitBlt( hdc, x, y, bitmap.bmWidth, bitmap.bmHeight, compatibleDC, 0, 0, SRCCOPY ) == 0 )
    {
        throw SystemException( GetLastError(), L"BitBlt", __FUNCTION__ );
    }

    // Reset
    if ( oldBitmap )
    {
        SelectObject( compatibleDC, oldBitmap );
    }
    DeleteDC( compatibleDC );
}

//===============================================================================================//
//  Description:
//      Duplicate the specified string
//
//  Parameters:
//      pszString - the string to duplicate
//
//  Remarks:
//      Caller must use free to release the allocated memory
//
//  Returns:
//      pointer to duplicated string, NULL if input is NULL
//===============================================================================================//
char* PXSDuplicateStringA( const char* pszString )
{
    char*   psz;
    size_t  cch = 0;
    HRESULT hResult;

    if ( pszString == nullptr )
    {
        return nullptr;
    }

    hResult = StringCchLengthA( pszString , STRSAFE_MAX_CCH, &cch );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"StringCchLength", __FUNCTION__ );
    }

    cch = PXSAddSizeT( cch, 1 );        // Terminator
    psz = new char[ cch ];
    if ( psz == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }

    hResult = StringCchCopyA( psz, cch, pszString );
    if ( FAILED( hResult ) )
    {
        delete [] psz;
        throw ComException( hResult, L"StringCchCopyA", __FUNCTION__ );
    }
    psz[ cch - 1 ] = 0x00;

    return psz;
}
//===============================================================================================//
//  Description:
//      Draw a transparent bitmap
//
//  Parameters:
//      hdc         - handle to device context to draw on
//      hBitmap     - handle to bitmap
//      xPos        - x position of bitmap
//      yPos        - y position of bitmap
//      transparent - the transparent colour
//
//  Returns:
//      void
//===============================================================================================//
void PXSDrawTransparentBitmap( HDC hdc, HBITMAP hBitmap, int x, int y, COLORREF transparent )
{
    HDC     compatibleDC;
    BITMAP  bitmap;
    HGDIOBJ oldBitmap;

    // Check inputs
    if ( ( hdc     == nullptr ) ||
         ( hBitmap == nullptr ) ||
         ( transparent == CLR_INVALID ) )
    {
        return;     // Nothing to do
    }

    // Get the bitmap
    memset( &bitmap, 0, sizeof ( bitmap ) );
    if ( GetObject( hBitmap, sizeof ( bitmap ), &bitmap ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetObject", __FUNCTION__ );
    }

    // Select the bitmap into a compatible device context
    compatibleDC = CreateCompatibleDC( hdc );
    if ( compatibleDC == nullptr )
    {
        throw SystemException( GetLastError(), L"CreateCompatibleDC", __FUNCTION__ );
    }
    oldBitmap = SelectObject( compatibleDC, hBitmap );
    if ( TransparentBlt( hdc,
                         x,
                         y,
                         bitmap.bmWidth,
                         bitmap.bmHeight,
                         compatibleDC, 0, 0, bitmap.bmWidth, bitmap.bmHeight, transparent ) == 0 )
    {
        throw SystemException( GetLastError(), L"TransparentBlt", __FUNCTION__ );
    }

    // Reset
    if ( oldBitmap )
    {
        SelectObject( compatibleDC, oldBitmap );
    }
    DeleteDC( compatibleDC );
}

//===============================================================================================//
//  Description:
//      Convert the specified character to a rich text character sequence
//
//  Parameters:
//      ch        - the character
//      pszBuffer - buffer to receive the entity
//      bufChars  - size of the buffer, Longest is \u-nnnnn? so the buffer
//                  must be at least 10 chars long to include the terminator
//
//  Remarks:
//      Will convert &<>" and those above 0xFF
//
//  Returns:
//      Number of characters written to the buffer NOT including the NULL
//      Zero if the character does not need converting
//===============================================================================================//
size_t PXSEscapeRichTextChar( wchar_t ch, wchar_t* pszBuffer, size_t bufChars )
{
    size_t numCopied = 0;

    if ( pszBuffer == nullptr )
    {
        throw ParameterException( L"pszBuffer", __FUNCTION__ );
    }

    // Need to hold up to \u-nnnnn? + terminator
    if ( bufChars < 10 )
    {
        throw SystemException( ERROR_INSUFFICIENT_BUFFER, L"bufChars", __FUNCTION__ );
    }
    *pszBuffer = PXS_CHAR_NULL;

    // These need escaping with a backslash
    if ( ( ch == '\\' ) ||
         ( ch == '{'  ) ||
         ( ch == '}'  )  )
    {
        pszBuffer[ 0 ] = '\\';
        pszBuffer[ 1 ] = ch;
        numCopied = 2;
    }
    else if ( ( ch > 0x80 ) && ( ch <= 0xFF ) )
    {
        // Hi-ASCII, escape using \'hh
        StringCchPrintf( pszBuffer, bufChars, L"\\'%02x", ch );
        numCopied = 4;
    }
    else if ( ch >= 0x100 )
    {
        // 2-bytes, escape using decimal \unnnnn? or \u-nnnnn? RTF uses
        // signed values so above 32767 must be expressed as negative.
        if ( ch >= 0x8000 )
        {
            StringCchPrintf( pszBuffer, bufChars, L"\\u-%05d?", 0x10000 - ch );
            numCopied = 9;
        }
        else
        {
            StringCchPrintf( pszBuffer, bufChars, L"\\u%05d?", ch );
            numCopied = 8;
        }
    }

    return numCopied;
}

//===============================================================================================//
//  Description:
//      Wrapper for ExpandEnvironmentStrings
//
//  Parameters:
//      pszSource - the environment string to expand
//      pExpanded - receives the expanded string
//
//  Returns:
//      void
//===============================================================================================//
void PXSExpandEnvironmentStrings( LPCWSTR pszSource, String* pExpanded )
{
    DWORD    numChars;
    wchar_t* pszBuffer;
    AllocateWChars WChars;

    // Filter out "" so can test for zero as error condition below
    if ( ( pszSource == nullptr ) || ( *pszSource == PXS_CHAR_NULL ) )
    {
        return;     // Nothing to do
    }

    if ( pExpanded == nullptr )
    {
        throw ParameterException( L"pExpanded", __FUNCTION__ );
    }
    *pExpanded = PXS_STRING_EMPTY;

    // Get the buffer size required, overestimate as things change
    numChars  = ExpandEnvironmentStrings( pszSource, nullptr, 0 );
    numChars  = PXSMultiplyUInt32( numChars, 2 );
    numChars  = PXSAddUInt32( numChars, 1 );      // Terminator
    pszBuffer = WChars.New( numChars );
    if ( ExpandEnvironmentStrings( pszSource, pszBuffer, numChars ) == 0 )
    {
        throw SystemException( GetLastError(), pszSource, "ExpandEnvironmentStrings" );
    }
    pszBuffer[ numChars - 1 ] = PXS_CHAR_NULL;
    *pExpanded = pszBuffer;
}

//===============================================================================================//
//
//  Description:
//      Compute the Fletcher checksum for a block of data
//
//  Parameters:
//      pData - pointer to the data
//      bytes - number of bytes of data
//
//  Remarks:
//      As per RFC 1146 and Wikipedia
//      This checksum can detect byte reversal making it good for
//      endian error detection. This is of course less robust than a
//      CRC-32 but only requires 2 bytes of storage for the checksum.
//
//  Returns:
//      unsigned 16-bit integer which is the checksum
//===============================================================================================//
WORD PXSFletcherChecksum16( const BYTE* pData , size_t bytes )
{
    size_t len = 0, sum1 = 0xff, sum2 = 0xff;

    if ( pData == nullptr || bytes == 0 )
    {
        return 0;   // Nothing to do
    }

    while ( bytes )
    {
        // Set a maximum of 20 loops at which sum2 may exceed 0xFFFF.
        len    = PXSMinSizeT( bytes, 20 );
        bytes -= len;

        do
        {
            sum1 += *pData++;
            sum2 += sum1;
            len--;
        } while ( len );

        // Reduce
        sum1 = ( sum1 & 0xff ) + ( sum1 >> 8 );
        sum2 = ( sum2 & 0xff ) + ( sum2 >> 8 );
    }

    /// Reduce
    sum1 = ( sum1 & 0xff ) + ( sum1 >> 8 );
    sum2 = ( sum2 & 0xff ) + ( sum2 >> 8 );

    return ( (WORD)( ( sum2 << 8 ) | sum1 ) );
}

//===============================================================================================//
//  Description:
//      Get the application's name
//
//  Parameters:
//      pApplicationName - optional, receives the application name
//
//  Returns:
//===============================================================================================//
void PXSGetApplicationName( String* pApplicationName )
{
    // Must have allocated the application object
    if ( g_pApplication )
    {
        g_pApplication->GetApplicationName( pApplicationName );
    }
}

//===============================================================================================//
//  Description:
//      Get the directory of the currently running programme
//
//  Parameters:
//      pDirectoryPath - receives the programme directory string
//
//  Remarks:
//      Always ensure there is a trailing backslash '\' if function succeeds
//
//  Returns:
//      void
//===============================================================================================//
void PXSGetExeDirectory( String* pExeDirectory )
{
    size_t index;

    if ( pExeDirectory == nullptr )
    {
        throw ParameterException( L"pExeDirectory", __FUNCTION__ );
    }
    PXSGetExePath( pExeDirectory );
    index = pExeDirectory->LastIndexOf( PXS_PATH_SEPARATOR );
    if ( index == PXS_MINUS_ONE )
    {
        throw SystemException( ERROR_BAD_PATHNAME, pExeDirectory->c_str(), __FUNCTION__ );
    }
    pExeDirectory->Truncate( index + 1 );
    pExeDirectory->Trim();
}

//===============================================================================================//
//  Description:
//      Get the executable name for this process.
//
//  Parameters:
//      pExePath - String to receive the executable's path
//
//  Returns:
//      void
//===============================================================================================//
void PXSGetExePath( String* pExePath )
{
    wchar_t  szPath[ MAX_PATH + 1 ] = { 0 };

    if ( pExePath == nullptr )
    {
        throw ParameterException( L"pExePath", __FUNCTION__ );
    }
    *pExePath = PXS_STRING_EMPTY;

    if ( GetModuleFileName( nullptr, szPath, ARRAYSIZE( szPath ) ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetModuleFileName", __FUNCTION__ );
    }
    szPath[ ARRAYSIZE( szPath ) - 1 ] = PXS_CHAR_NULL;
    *pExePath = szPath;
}

//===============================================================================================//
//  Description:
//      Convert the specified character to a HTML character entity reference
//
//  Parameters:
//      ch        - the character
//      pszBuffer - buffer to receive the entity
//      bufChars  - size of the buffer. Longest is &#ddddd; so the buffer
//                  must be at least 9 chars long  to include the terminator
//
//  Remarks:
//      Will convert &<>" and those above 0xFF
//
//  Returns:
//      Number of characters written to the buffer NOT including the NULL
//      Zero if the character does not need converting
//===============================================================================================//
size_t PXSGetHtmlCharacterEntity( wchar_t ch, wchar_t* pszBuffer, size_t bufChars )
{
    size_t numCopied = 0;

    if ( pszBuffer == nullptr )
    {
        throw ParameterException( L"pszBuffer", __FUNCTION__ );
    }

    // Need to hold up to &#ddddd; + terminator
    if ( bufChars < 9 )
    {
        throw SystemException( ERROR_INSUFFICIENT_BUFFER, L"bufChars", __FUNCTION__ );
    }
    *pszBuffer = PXS_CHAR_NULL;

    if ( ch == '&' )
    {
        StringCchCopy( pszBuffer, bufChars, L"&amp;" );
        numCopied = 5;
    }
    else if ( ch == '<' )
    {
        StringCchCopy( pszBuffer, bufChars, L"&lt;" );
        numCopied = 4;
    }
    else if ( ch == '>' )
    {
        StringCchCopy( pszBuffer, bufChars, L"&gt;" );
        numCopied = 4;
    }
    else if ( ch == '"' )
    {
        StringCchCopy( pszBuffer, bufChars, L"&quot;" );
        numCopied = 6;
    }
    else if ( ch >= 0x100 )
    {
        // Use form &#ddddd;
        StringCchPrintf( pszBuffer, bufChars, L"&#%d;", ch );
        numCopied = wcslen( pszBuffer );
    }

    return numCopied;
}

//===============================================================================================//
//  Description:
//      Get a string based on its ID
//
//  Parameters:
//      resourceID     - unique ID for the string
//      pResourceString- optional, receives the resource string
//
//  Returns:
//      void
//===============================================================================================//
void PXSGetResourceString( DWORD resourceID, String* pResourceString )
{
    // Must have allocated the application object
    if ( g_pApplication )
    {
        g_pApplication->GetResourceString( resourceID, pResourceString );
    }
}

//===============================================================================================//
//  Description:
//      Get a string based on its ID
//
//  Parameters:
//      resourceID     - unique ID for the string
//      Insert1        - insert string for placeholder %%1
//      pResourceString- receives the resource string
//
//  Returns:
//      void
//===============================================================================================//
void PXSGetResourceString1( DWORD resourceID, const String& Insert1, String* pResourceString )
{
    // Must have allocated the application object
    if ( g_pApplication )
    {
        g_pApplication->GetResourceString1( resourceID, Insert1, pResourceString );
    }
}

//===============================================================================================//
//  Description:
//      Get the specified stock font
//
//  Parameters:
//      stockFont - defined stock font constant
//
//  Returns:
//      void
//===============================================================================================//
HFONT PXSGetStockFont( int stockFont )
{
    HFONT hFont;

    // These are currently documented
    if ( ( stockFont != ANSI_FIXED_FONT     ) &&
         ( stockFont != ANSI_VAR_FONT       ) &&
         ( stockFont != DEVICE_DEFAULT_FONT ) &&
         ( stockFont != DEFAULT_GUI_FONT    ) &&
         ( stockFont != OEM_FIXED_FONT      ) &&
         ( stockFont != SYSTEM_FONT         ) &&
         ( stockFont != SYSTEM_FIXED_FONT   )  )
    {
        throw ParameterException( L"stockFont", __FUNCTION__ );
    }

    hFont = (HFONT)GetStockObject( stockFont );
    if ( hFont == nullptr )
    {
        throw SystemException( GetLastError(), L"GetStockObject", __FUNCTION__ );
    }

    return hFont;
}

//===============================================================================================//
//  Description:
//      Get the TEXTMETRIC for the specified stock font in the specifed window
//
//  Parameters:
//      hWindow   - window, if NULL will use the desktop
//      stockFont - defined stock font constant
//      lptm      - receives the metrics
//
//  Returns:
//      void
//===============================================================================================//
void PXSGetStockFontTextMetrics( HWND hWindow, int stockFont, TEXTMETRIC* lptm )
{
    HDC     hDC;
    HFONT   hFont;
    DWORD   lastError = ERROR_SUCCESS;
    HGDIOBJ oldFont;

    if ( lptm == nullptr )
    {
        throw ParameterException( L"lptm", __FUNCTION__ );
    }
    memset( lptm, 0, sizeof ( *lptm ) );

    hFont = PXSGetStockFont( stockFont );
    hDC   = GetDC( hWindow );
    if ( hDC == nullptr )
    {
        throw SystemException( GetLastError(), L"GetDC", __FUNCTION__ );
    }
    oldFont = SelectObject( hDC, hFont );
    if ( GetTextMetrics( hDC, lptm ) == 0 )
    {
        lastError = GetLastError();
    }

    // Reset
    if ( oldFont )
    {
        SelectObject( hDC, oldFont );
    }
    ReleaseDC( hWindow, hDC );

    // Now test for GetTextMetrics error
    if ( lastError != ERROR_SUCCESS )
    {
        throw SystemException( GetLastError(), L"GetTextMetrics", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Convert the specified hexadecimal to an unsigned char character e.g. 0x20 = 32
//
//  Parameters:
//      hexdig1 - first character
//      hexdig2 - second character
//
//  Returns:
//      unsigned char
//===============================================================================================//
BYTE PXSHexDigitsToByte( wchar_t hexdig1, wchar_t hexdig2 )
{
    if ( ( hexdig1 >= 'a' ) && ( hexdig1 <= 'f' ) )
    {
        hexdig1 -= ( 'a' - 10 );
    }
    else if ( ( hexdig1 >= 'A' ) && ( hexdig1 <= 'F' ) )
    {
        hexdig1 -= ( 'A' - 10 );
    }
    else if ( ( hexdig1 >= '0' ) && ( hexdig1 <= '9' ) )
    {
        hexdig1 -= '0';
    }
    else
    {
        throw ParameterException( L"hexdig1", __FUNCTION__ );
    }

    if ( ( hexdig2 >= 'a' ) && ( hexdig2 <= 'f' ) )
    {
        hexdig2 -= ( 'a' - 10 );
    }
    else if ( ( hexdig2 >= 'A' ) && ( hexdig2 <= 'F' ) )
    {
        hexdig2 -= ( 'A' - 10 );
    }
    else if ( ( hexdig2 >= '0' ) && ( hexdig2 <= '9' ) )
    {
        hexdig2 -= '0';
    }
    else
    {
        throw ParameterException( L"hexdig2", __FUNCTION__ );
    }

    return static_cast<BYTE>( ( hexdig1 * 16 ) + hexdig2 );
}

//===============================================================================================//
//  Description:
//      Convert the specified hexadecimal to a character e.g. 20 = 32 = space
//
//  Parameters:
//      hexdig1 - first character
//      hexdig2 - second character
//
//  Returns:
//      character
//===============================================================================================//
char PXSHexDigitsToChar( char hexdig1, char hexdig2 )
{
    char    ch = 0;
    BYTE    b;
    wchar_t wch1, wch2;

    if ( ( hexdig1 < 0 ) || ( hexdig2 < 0 ) )
    {
        throw ParameterException( L"hexdig1/hexdig2", __FUNCTION__ );
    }
    wch1 = PXSCastCharToWChar( hexdig1 );
    wch2 = PXSCastCharToWChar( hexdig2 );
    b    = PXSHexDigitsToByte( wch1, wch2 );
    if ( b > 127 )
    {
        ch = static_cast< char >( -(256 - b) );
    }
    else
    {
        ch = static_cast< char >( b );
    }

    return ch;
}

//===============================================================================================//
//  Description:
//      Initialize COM on the calling thread
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void PXSInitializeComOnThread()
{
    HRESULT hResult;

    // Second parameter COINIT_MULTITHREADED = 0x0
    hResult = CoInitializeEx(nullptr,
                             COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if ( FAILED( hResult ) )
    {
         if ( hResult == S_FALSE )
         {
             // The COM library is already initialized so balance calls
             CoUninitialize();
         }
         else if ( RPC_E_CHANGED_MODE != hResult )
         {
             throw ComException( hResult, L"CoUninitialize", __FUNCTION__ );
         }
    }
}

//===============================================================================================//
//  Description:
//      Test if specified character is alpha character, i.e. a-z or A-Z
//
//  Parameters:
//      ch - the character to test
//
//  Returns:
//      true if the input char is [a-zA-Z]
//===============================================================================================//
bool PXSIsAlphaCharacter( wchar_t ch )
{
    if ( ( ch >= 'A' ) && ( ch <= 'Z' ) )
    {
        return true;
    }

    if ( ( ch >= 'a' ) && ( ch <= 'z' ) )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Determine if the aplication level logger is logging
//
//  Parameters:
//      none
//
//  Returns:
//      true if logging otherwise false
//===============================================================================================//
bool PXSIsAppLogging()
{
    if ( g_pApplication == nullptr )
    {
        return false;
    }

    return g_pApplication->IsLogging();
}

//===============================================================================================//
//  Description:
//      Determine if the specified character is a digit, base 10
//
//  Parameters:
//      ch - the character to test
//
//  Returns:
//      true if the character is a base 10 digit
//===============================================================================================//
bool PXSIsDigit( wchar_t ch )
{
    if ( ( ch < '0' ) || ( ch > '9' ) )
    {
        return false;
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Test if this string comprises of only digits 0..9
//
//  Parameters:
//      none
//
//  Returns:
//      true if only digits, otherwise false
//===============================================================================================//
bool PXSIsDigitsOnly( const wchar_t* pString, size_t numChars )
{
    bool   isDigits = true;
    size_t i  = 0;

    if ( ( pString == nullptr ) || ( numChars == 0 ) )
    {
        return false;
    }

    while ( ( i < numChars ) && isDigits )
    {
        isDigits = PXSIsDigit( pString[ i ] );
        i++;
    }

    return isDigits;
}

//===============================================================================================//
//  Description:
//      Determine if the specified character is a hexit
//
//  Parameters:
//      ch - the character to test
//
//  Remarks
//      Will not change the case of the input character as may result in
//      non-hexadecimal characters on non-English locales
//
//  Returns:
//      true if the character is a hexit, otherwise false
//===============================================================================================//
bool PXSIsHexitA( char ch )
{
    wchar_t wch;

    if ( ch < 0 )
    {
        return false;
    }
    wch = PXSCastCharToWChar( ch );

    return PXSIsHexitW( wch );
}

//===============================================================================================//
//  Description:
//      Determine if the specified character is a hexit
//
//  Parameters:
//      ch - the character to test
//
//  Remarks
//      Will not change the case of the input character as may result in
//      non-hexadecimal characters on non-English locales
//
//  Returns:
//      true if the character is a hexit, otherwise false
//===============================================================================================//
bool PXSIsHexitW( wchar_t ch )
{
    if ( ( ( ch >= '0' ) && ( ch <= '9' ) ) ||
         ( ( ch >= 'A' ) && ( ch <= 'F' ) ) ||
         ( ( ch >= 'a' ) && ( ch <= 'f' ) )  )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Get if the device context is hi-colour
//
//  Parameters:
//      hDC - handle to device context, if NULL will use the desktop
//
//  Remarks:
//      Various test to see if the device is using a large number of colours
//
//  Returns:
//      true is high colour, else false
//===============================================================================================//
bool PXSIsHighColour( HDC hDC )
{
    const int HIGH_COLOR_BITS = 16;
    bool  hiColour = false;
    HDC   hdcTop   = nullptr;
    SystemInformation SystemInfo;

    // Use bit count as indication of number of colours, i.e. ignore
    // the number of colour planes supported by the device
    if ( hDC )
    {
        if ( GetDeviceCaps( hDC, BITSPIXEL ) >= HIGH_COLOR_BITS )
        {
            hiColour = true;
        }
    }
    else
    {
        // Get the device context of the desktop
        hdcTop = GetDC( nullptr );
        if ( hdcTop )
        {
            if ( GetDeviceCaps( hdcTop, BITSPIXEL ) >= HIGH_COLOR_BITS )
            {
                hiColour = true;
            }

            ReleaseDC( nullptr, hdcTop );
        }
    }

    // When in high contrast say the display is low colour resolution
    if ( PXSIsHighContrast() )
    {
        hiColour = false;
    }

    // Remote session, if so will say running as low colour to reduce bandwidth
    if ( SystemInfo.IsRemoteSession() )
    {
        hiColour = false;
    }

    return hiColour;
}

//===============================================================================================//
//  Description:
//      Get if the desktop is in high contrast mode
//
//  Parameters:
//      None
//
//  Remarks:
//      Test various settings to see if the desktop appears to have a high
//      contrast scheme
//
//  Returns:
//      true if desktop is in hi-contrast mode
//===============================================================================================//
bool PXSIsHighContrast()
{
    bool hiContrast = false;
    HIGHCONTRAST hc = { 0, 0, nullptr };

    // High contrast setting
    hc.cbSize = sizeof ( HIGHCONTRAST );
    if ( SystemParametersInfo( SPI_GETHIGHCONTRAST, 0, &hc, 0 ) )
    {
        // Documentation says HCF_INDICATOR is not used. The flag is not set
        // by choosing a high contrast colour scheme. It is set using the
        // Accessibility Control Panel
        if ( HCF_HIGHCONTRASTON & hc.dwFlags )
        {
            hiContrast = true;
        }
    }

    // Test for System colours
    // Window Background
    if ( hiContrast == false )
    {
        if ( PXS_COLOUR_BLACK == GetSysColor( COLOR_WINDOW ) )
        {
            hiContrast = true;
        }
    }

    // Menu Background
    if ( hiContrast == false )
    {
        if ( PXS_COLOUR_BLACK == GetSysColor( COLOR_MENU ) )
        {
            hiContrast = true;
        }
    }

    // Window Title Bar Background
    if ( hiContrast == false )
    {
        if ( PXS_COLOUR_BLACK == GetSysColor( COLOR_ACTIVECAPTION ) )
        {
            hiContrast = true;
        }
    }

    return hiContrast;
}

//===============================================================================================//
//  Description:
//      Determine if the specified unsigned 32-bit integer is in the
//      specified array. The array does not have to be sorted.
//
//  Parameters:
//      value       - the value to search for
//      pArray      - pointer to a DWORD array
//      numElements - the number of elements in the array
//
//  Returns:
//      true if found otherwise false
//===============================================================================================//
bool PXSIsInUInt32Array( DWORD value, const DWORD* pArray, size_t numElements )
{
    bool   found = false;
    size_t i = 0;

    if ( pArray == nullptr || numElements == 0 )
    {
        return false;
    }

    while ( ( i < numElements ) && ( found == false )  )
    {
        if ( value == pArray[ i ] )
        {
            found = true;
        }
        i++;
    }

    return found;
}

//===============================================================================================//
//  Description:
//      Test if specified character is valid for a logical drive
//
//  Parameters:
//      ch - the character to test
//
//  Remarks:
//      Test if for A to Z but does not verify the drive actually exists
//  Returns:
//      true if the input char is [a-zA-Z]
//===============================================================================================//
bool PXSIsLogicalDriveLetter( wchar_t ch )
{
    if ( ( ch >= 'A' ) && ( ch <= 'Z' ) )
    {
        return true;
    }

    if ( ( ch >= 'a' ) && ( ch <= 'z' ) )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Test if specified double value is not-a-number
//
//  Parameters:
//      pszString - the string
//      ch        - the repeated character to test for
//
//  Returns:
//      true if only is a nan otherwise false
//===============================================================================================//
bool PXSIsNanDouble( double value )
{
    bool bNan = false;

    #ifdef _MSC_VER

        if ( _isnan( value ) )
        {
            bNan = true;
        }

    #elif __GNUC__

        // cppcheck "bNan reassigned value" but leave as is for readability
        bNan = std::isnan( value );

    #else

        #error Unsupported compiler

    #endif

    return bNan;
}

//===============================================================================================//
//  Description:
//      Determine if the specified character is valid for a computer name
//
//  Parameters:
//      wch - the character to test
//
//  Returns:
//      true if valid otherwise false
//===============================================================================================//
bool PXSIsValidComputerNameChar( wchar_t wch )
{
    if ( wcschr( PXS_STRING_COMPUTER_CHARS, wch ) )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Determine if the specified state code is a valid
//
//  Parameters:
//      state - a defined value representing the state
//
//  Returns:
//      true if valid otherwise false
//===============================================================================================//
bool PXSIsValidStateCode( DWORD state )
{
    if ( ( state == PXS_STATE_UNKNOWN      ) ||
         ( state == PXS_STATE_ERROR        ) ||
         ( state == PXS_STATE_WAITING      ) ||
         ( state == PXS_STATE_EXECUTING    ) ||
         ( state == PXS_STATE_CANCELLED    ) ||
         ( state == PXS_STATE_COMPLETED    ) ||
         ( state == PXS_STATE_RESOLVE_WAIT ) ||
         ( state == PXS_STATE_RESOLVING    ) ||
         ( state == PXS_STATE_CONNECT_WAIT ) ||
         ( state == PXS_STATE_CONNECTED    )  )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Determine if the input time structure has valid values
//
//  Parameters:
//      t - the time structure
//
//  Remarks:
//      Does not examine the day of the week or leap seconds for validity
//
//  Returns:
//      true if valid otherwise false
//===============================================================================================//
bool PXSIsValidStructTM( const struct tm& t )
{
    // Test values
    if ( ( t.tm_yday <  0 ) ||          // Before 1900
         ( t.tm_mday <  1 ) ||
         ( t.tm_mday > 31 ) ||
         ( t.tm_hour <  0 ) ||
         ( t.tm_hour > 23 ) ||
         ( t.tm_min  <  0 ) ||
         ( t.tm_min  > 59 ) ||
         ( t.tm_sec  <  0 ) ||
         ( t.tm_sec  > 59 )  )
    {
        return false;
    }

    // Test days of specific months
    if ( t.tm_mday > 30 )
    {
        if ( ( t.tm_mon ==  3 ) ||     // Apr
             ( t.tm_mon ==  5 ) ||     // Jun
             ( t.tm_mon ==  8 ) ||     // Sep
             ( t.tm_mon == 10 )  )     // Nov
        {
            return false;
        }
    }

    // February
    if ( t.tm_mon == 1 )
    {
        if ( t.tm_year % 4 )        // Non-leap year
        {
            if ( t.tm_mday > 29 )
            {
                return false;
            }
        }
        else                        // Leap year
        {
            if ( t.tm_mday > 30 )
            {
                return false;
            }
        }
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Determine if the specified character a valid for UTF-16
//
//  Parameters:
//      wch - the character
//
//  Remarks:
//      Taken from MSDN "Surrogates and Supplementary Characters"
//
//  Returns:
//      true if valid otherwise false
//===============================================================================================//
bool PXSIsValidUtf16( wchar_t wch )
{
    if ( wch == 0xFFFE )
    {
        return false;
    }

    if ( wch == 0xFFFF )
    {
        return false;
    }

    if ( ( wch >= 0xFDD0 ) && ( wch <= 0xFDEF ) )
    {
        return false;
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Test if the specified string is comprised of only the specified character
//
//  Parameters:
//      pString  - the string to test
//      numChars - the character length of the string
//      wch      - the repeated character to test for
//
//  Returns:
//      true if all characters are the same as the input
//===============================================================================================//
bool PXSIsOnlyChar( const wchar_t* pString, size_t numChars, wchar_t wch )
{
    bool   isOnly = true;
    size_t i = 0;

    if ( ( pString == nullptr ) || ( numChars == 0 ) )
    {
        return false;
    }

    while ( ( i < numChars ) && isOnly )
    {
        if ( pString[ i ] != wch )
        {
            isOnly = false;
        }
        i++;
    }

    return isOnly;
}

//===============================================================================================//
//  Description:
//      Simple heuristic test if the string seems to have UTF-8 characters
//
//  Parameters:
//      pString  - the string to test
//      numChars - length of string in chars
//
//  Remarks:
//      This a simple test that scans for the presence of what is a valid UTF-8
//      character. Since UTF-16 is two bytes will only check up to code point
//      U+FFFF which is up to three bytes, i.e. the first multilingual plane.
//      Two Bytes  : 110xxxxx 10xxxxxx
//      Three Bytes: 1110xxxx 10xxxxxx 10xxxxxx
//
//      Wikipedia: The chance of a random string of bytes being valid UTF-8 and not pure ASCII
//      is 3.9% for a two-byte sequence, 0.41% for a three-byte sequence and 0.026% for a
//      four-byte sequence.
//
//  Returns:
//    true if only contains a UTF-8 character
//===============================================================================================//
bool PXSIsUtf8( char* pString, size_t numChars )
{
    bool isUtf8 = false;
    BYTE ch1    = 0, ch2 = 0, ch3 = 0;
    size_t  i   = 0;

    if ( pString == nullptr )
    {
        return false;
    }

    // First pass, test for char 1 = 110xxxxx and char 2 = 10xxxxxx
    if ( numChars < 2 )
    {
        return false;
    }
    while ( ( i < ( numChars - 1 ) ) && ( isUtf8 == false ) )
    {
        ch1 = static_cast< BYTE >( pString[ i ] );
        ch2 = static_cast< BYTE >( pString[ i + 1 ] );
        if ( ( ch1 >= 192 ) && ( ch1 <= 223 ) &&
             ( ch2 >= 128 ) && ( ch2 <= 191 )  )
        {
            isUtf8 = true;
        }
        i++;
    }

    if ( isUtf8 )
    {
        return true;
    }

    // Second pass, test for char 1 = 1110xxxx, char 2 = 10xxxxxx and char 3 = 10xxxxxx
    if ( numChars < 3 )
    {
        return false;
    }
    i = 0;
    while ( ( i < ( numChars - 2 ) ) && ( isUtf8 == false ) )
    {
        ch1 = static_cast< BYTE >( pString[ i ] );
        ch2 = static_cast< BYTE >( pString[ i + 1 ] );
        ch3 = static_cast< BYTE >( pString[ i + 2 ] );
        if ( ( ch1 >= 224 ) && ( ch1 <= 239 ) &&
             ( ch2 >= 128 ) && ( ch2 <= 191 ) &&
             ( ch3 >= 128 ) && ( ch3 <= 191 )  )
        {
            isUtf8 = true;
        }
        i++;
    }

    return isUtf8;
}

//===============================================================================================//
//  Description:
//      Determine if the specified character is a white space
//
//  Parameters:
//      wch - the character to test
//
//  Remarks
//      Will ignore <control-0085>
//
//  Returns:
//      true if the character is a white space, otherwise false
//===============================================================================================//
bool PXSIsWhiteSpace( wchar_t wch )
{
    // Most characters are not spaces
    if ( ( wch > 0x20 ) && ( wch < 0xA0 ) )
    {
        return false;
    }

    if ( ( wch == 0x20 ) ||                         // SPACE
         ( wch == 0xA0 ) ||                         // NO-BREAK SPACE
         ( ( wch >= 0x09 ) && ( wch <= 0x0D ) ) )   // <control>
    {
        return true;
    }

    // Unicode white spaces, see UAX44, these were taken from PropList-6.2.0.txt
    if ( ( wch == 0x1680 ) ||                       // OGHAM SPACE MARK
         ( wch == 0x180E ) ||                       // MONGOLIAN VOWEL SEPARATOR
         ( wch == 0x2028 ) ||                       // LINE SEPARATOR
         ( wch == 0x2029 ) ||                       // PARAGRAPH SEPARATOR
         ( wch == 0x202F ) ||                       // NARROW NO-BREAK SPACE
         ( wch == 0x205F ) ||                       // MEDIUM MATHEMATICAL SPACE
         ( wch == 0x3000 ) ||                       // IDEOGRAPHIC SPACE
         ( wch == 0xFEFF )  )                       // BOM
    {
        return true;
    }

    // These too
    if ( ( wch >= 0x2000 ) && ( wch <= 0x200A ) )   // EN QUAD..HAIR SPACE
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Determine if the specified character is a white space
//
//  Parameters:
//      ch - the character to test
//
//  Returns:
//      true if the character is a white space, otherwise false
//===============================================================================================//
bool PXSIsWhiteSpaceA( char ch )
{
    // Most characters are not spaces
    if ( ( ch > 0x20 ) && ( ch < 0xA0 ) )
    {
        return false;
    }

    if ( ( ch == 0x20 ) ||                       // SPACE
         ( ch == 0xA0 ) ||                       // NO-BREAK SPACE
         ( ( ch >= 0x09 ) && ( ch <= 0x0D ) ) )  // <control>
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Make an signed 32-bit integer from the input bytes
//
//  Parameters:
//      b1 - the most significant byte
//      b2 - the next most significant byte
//      b3 - the next most significant byte
//      b4 - the least significant byte
//
//  Returns:
//      int value
//===============================================================================================//
int PXSMakeInt32( BYTE b1, BYTE b2, BYTE b3, BYTE b4 )
{
    return ( b1 + ( b2 << 8 ) + ( b3 << 16 ) + ( b4 << 24 ) );
}

//===============================================================================================//
//  Description:
//      Make an unsigned 16-bit integer from the input bytes
//
//  Parameters:
//      b1 - the high byte
//      b2 - the next most significant byte
//
//  Returns:
//      WORD value
//===============================================================================================//
WORD PXSMakeUInt16( BYTE b1, BYTE b2 )
{
    return PXSAddUInt16( b1, (WORD)( b2 << 8 ) );
}

//===============================================================================================//
//  Description:
//      Make an unsigned 32-bit integer from the input bytes
//
//  Parameters:
//      b1 - the most significant byte
//      b2 - the next most significant byte
//      b3 - the next most significant byte
//      b4 - the least significant byte
//
//  Returns:
//      DWORD value
//===============================================================================================//
DWORD PXSMakeUInt32( BYTE b1, BYTE b2, BYTE b3, BYTE b4 )
{
    return ( b1 + ( (DWORD)b2 << 8 ) + ( (DWORD)b3 << 16 ) + ( (DWORD)b4 << 24 ) );
}

//===============================================================================================//
//  Description:
//      Use OutputDebugString to write out the specified string and tick count
//
//  Parameters:
//      pszString - the message
//
//  Returns:
//      void
//===============================================================================================//
void PXSOutputDebugTickCount( LPCWSTR pszString )
{
    char buffer[ 128 ] = {0 };

    if ( pszString )
    {
        OutputDebugStringW( pszString );
    }

    StringCchPrintfA( buffer, ARRAYSIZE( buffer ), ": tick = %lums\r\n", GetTickCount() );
    OutputDebugStringA( buffer );
}

//===============================================================================================//
//  Description:
//      Get the specified RT_RCDATA data resource
//
//  Parameters:
//      Name    - the resource's name
//      pBuffer - buffer to receive the bytes
//
//  Returns:
//      void
//===============================================================================================//
void PXSLoadDataResource( const String& Name, ByteArray* pBuffer )
{
    PXSLoadResource( Name.c_str(), RT_RCDATA, pBuffer );
}

//===============================================================================================//
//  Description:
//      Get the specified html resource
//
//  Parameters:
//      resourceID - the id of the resource
//      pBuffer    - buffer to receive the bytes
//
//  Returns:
//      void
//===============================================================================================//
void PXSLoadHtmlResource( WORD resourceID, ByteArray* pBuffer )
{
    PXSLoadResource( MAKEINTRESOURCE( resourceID ), RT_HTML, pBuffer );
}

//===============================================================================================//
//  Description:
//      Get the specified executable resource
//
//  Parameters:
//      lpName  - the name of the resource, could be an integer id
//      lpType  - the resource type, e.g. RT_HTML
//      pBuffer  - buffer to receive the bytes
//
//
//  Return:
//      It's messy creating an error message as lpName may either be a string or a resource
//      identifier cast to a string pointer. Since resources are compiled into the exe, not
//      expecting any errors so will just report the failing system function.
//
//  Returns:
//      void
//===============================================================================================//
void PXSLoadResource( LPCWSTR lpName, LPCWSTR lpType, ByteArray* pBuffer )
{
    DWORD     resSize;
    void*     pResource;
    HGLOBAL   hResource;
    HRSRC     hrsrcFile;

    if ( ( lpName == 0 ) || ( lpType == nullptr ) || ( pBuffer == nullptr ) )
    {
        throw ParameterException( L"lpName/lpType/pBuffer", __FUNCTION__ );
    }
    pBuffer->Free();

    hrsrcFile = FindResource( nullptr, lpName, lpType );
    if ( hrsrcFile == nullptr )
    {
        throw SystemException( GetLastError(), L"FindResource", __FUNCTION__ );
    }
    resSize = SizeofResource( nullptr, hrsrcFile );

    // It is not necessary to destroy the resource
    hResource = LoadResource( nullptr, hrsrcFile );
    if ( hResource == nullptr )
    {
        throw SystemException( GetLastError(), L"LoadResource", __FUNCTION__ );
    }

    // It is not necessary to unlock the resource
    pResource = LockResource( hResource );
    if ( pResource == nullptr )
    {
        throw SystemException( GetLastError(), L"LockResource", __FUNCTION__ );
    }
    pBuffer->Append( static_cast<BYTE*>(pResource), resSize );
}

//===============================================================================================//
//  Description:
//      Write an application error message to the application's log
//
//  Parameters:
//      pszMessage - the message
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppError( LPCWSTR pszMessage )
{
    String Insert1, Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                          PXS_ERROR_TYPE_APPLICATION, 0, false, pszMessage, Insert1, Insert2 );
    }
    catch( Exception& )
    { }
}

//===============================================================================================//
//  Description:
//      Write an application error message to the application's log
//
//  Parameters:
//      pszMessage - the message
//      Insert1    - string insert parameter, any %%1 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppError1( LPCWSTR pszMessage, const String& Insert1 )
{
    String Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                          PXS_ERROR_TYPE_APPLICATION, 0, false, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application error message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      Insert1    - string insert parameter, any %%1 are replaced
//      Insert2    - string insert parameter, any %%2 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppError2( LPCWSTR pszMessage,
                      const String& Insert1, const String& Insert2 )
{
    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                          PXS_ERROR_TYPE_APPLICATION, 0, false, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application error message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      Insert1    - string insert parameter, any %%1 are replaced
//      Insert2    - string insert parameter, any %%2 are replaced
//      Insert3    - string insert parameter, any %%3 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppError3( LPCWSTR pszMessage,
                      const String& Insert1, const String& Insert2, const String& Insert3 )
{
    String    Message, Empty1, Empty2;
    Formatter Format;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        Message = Format.String3( pszMessage, Insert1, Insert2, Insert3 );
        PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                          PXS_ERROR_TYPE_APPLICATION, 0, false, Message.c_str(), Empty1, Empty2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application error message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      Insert1    - DWORD parameter, any %%1 are replaced
//      value2     - DWORD parameter, any %%2 are replaced
//      value3     - DWORD parameter, any %%3 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppErrorStringUInt32_2( LPCWSTR pszMessage,
                                   const String& Insert1, DWORD value2, DWORD value3 )
{
    String    Message, Insert2, Insert3;
    Formatter Format;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        Insert2 = Format.UInt32( value2 );
        Insert3 = Format.UInt32( value3 );
        Message = Format.String3( pszMessage, Insert1, Insert2, Insert3 );
        Insert2.Zero();
        Insert3.Zero();
        PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                          PXS_ERROR_TYPE_APPLICATION,
                          0, false, Message.c_str(), Insert2, Insert3 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application error message to the application's log
//
//  Parameters:
//      pszMessage - the message with a substitutable parameter
//      value      - DWORD parameter, any %%1 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppErrorUInt32( LPCWSTR pszMessage, DWORD value )
{
    String Insert1, Insert2;
    Formatter Format;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        Insert1 = Format.UInt32( value );
        PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                          PXS_ERROR_TYPE_APPLICATION, 0, false, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application error message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      value1     - DWORD parameter, any %%1 are replaced
//      value2     - DWORD parameter, any %%2 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppErrorUInt32_2( LPCWSTR pszMessage, DWORD value1, DWORD value2 )
{
    String    Message, Insert1, Insert2;
    Formatter Format;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        Message = Format.StringUInt32_2( pszMessage, value1, value2 );
        PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                          PXS_ERROR_TYPE_APPLICATION,
                          0, false, Message.c_str(), Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application error message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      value1     - DWORD parameter, any %%1 are replaced
//      value2     - DWORD parameter, any %%2 are replaced
//      value3     - DWORD parameter, any %%3 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppErrorUInt32_3( LPCWSTR pszMessage, DWORD value1, DWORD value2, DWORD value3 )
{
    String    Message, Insert1, Insert2;
    Formatter Format;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        Message = Format.StringUInt32_3( pszMessage, value1, value2, value3 );
        PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                          PXS_ERROR_TYPE_APPLICATION,
                          0, false, Message.c_str(), Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application informational message to the application's log
//
//  Parameters:
//      pszMessage - the message
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppInfo( LPCWSTR pszMessage )
{
    String Insert1, Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_INFORMATION,
                          PXS_ERROR_TYPE_APPLICATION, 0, false, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application informational message to the application's log
//
//  Parameters:
//      pszMessage - the message with a substitutable parameter
//      Insert1    - string insert parameter, any %%1 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppInfo1( LPCWSTR pszMessage, const String& Insert1 )
{
    String Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_INFORMATION,
                          PXS_ERROR_TYPE_APPLICATION, 0, false, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application informational message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      Insert1    - string insert parameter, any %%1 are replaced
//      Insert2    - string insert parameter, any %%2 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppInfo2( LPCWSTR pszMessage,
                     const String& Insert1, const String& Insert2 )
{
    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_INFORMATION,
                          PXS_ERROR_TYPE_APPLICATION, 0, false, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application informational message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      Insert1    - DWORD parameter, any %%1 are replaced
//      value2     - DWORD parameter, any %%2 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppInfoStringUInt32( LPCWSTR pszMessage, const String& Insert1, DWORD value2 )
{
    String    Message, Insert2, Insert3;
    Formatter Format;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        Insert2 = Format.UInt32( value2 );
        Message = Format.String2( pszMessage, Insert1, Insert2 );
        Insert2.Zero();
        PXSLogWriteEntry( PXS_SEVERITY_INFORMATION,
                          PXS_ERROR_TYPE_APPLICATION,
                          0, false, Message.c_str(), Insert2, Insert3 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application informational message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      Insert1    - DWORD parameter, any %%1 are replaced
//      value2     - DWORD parameter, any %%2 are replaced
//      value3     - DWORD parameter, any %%3 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppInfoStringUInt32_2( LPCWSTR pszMessage,
                                  const String& Insert1, DWORD value2, DWORD value3 )
{
    String    Message, Insert2, Insert3;
    Formatter Format;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        Insert2 = Format.UInt32( value2 );
        Insert3 = Format.UInt32( value3 );
        Message = Format.String3( pszMessage, Insert1, Insert2, Insert3 );
        Insert2.Zero();
        Insert3.Zero();
        PXSLogWriteEntry( PXS_SEVERITY_INFORMATION,
                          PXS_ERROR_TYPE_APPLICATION,
                          0, false, Message.c_str(), Insert2, Insert3 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application informational message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      value      - DWORD parameter, any %%1 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppInfoUInt32( LPCWSTR pszMessage, DWORD value )
{
    String Insert1, Insert2;
    Formatter Format;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        Insert1 = Format.UInt32( value );
        PXSLogWriteEntry( PXS_SEVERITY_INFORMATION,
                          PXS_ERROR_TYPE_APPLICATION,
                          0, false, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application informational message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      value1     - DWORD parameter, any %%1 are replaced
//      value2     - DWORD parameter, any %%2 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppInfoUInt32_2( LPCWSTR pszMessage, DWORD value1, DWORD value2 )
{
    String    Message, Insert1, Insert2;
    Formatter Format;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        Message = Format.StringUInt32_2( pszMessage, value1, value2 );
        PXSLogWriteEntry( PXS_SEVERITY_INFORMATION,
                          PXS_ERROR_TYPE_APPLICATION,
                          0, false, Message.c_str(), Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application informational message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      value1     - DWORD parameter, any %%1 are replaced
//      value2     - DWORD parameter, any %%2 are replaced
//      value3     - DWORD parameter, any %%3 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppInfoUInt32_3( LPCWSTR pszMessage, DWORD value1, DWORD value2, DWORD value3 )
{
    String    Message, Insert1, Insert2;
    Formatter Format;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        Message = Format.StringUInt32_3( pszMessage, value1, value2, value3 );
        PXSLogWriteEntry( PXS_SEVERITY_INFORMATION,
                          PXS_ERROR_TYPE_APPLICATION,
                          0, false, Message.c_str(), Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application warning message to the application's log
//
//  Parameters:
//      pszMessage - the message
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppWarn( LPCWSTR pszMessage )
{
    String Insert1, Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_WARNING,
                          PXS_ERROR_TYPE_APPLICATION, 0, false, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application warning message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      Insert1    - string insert parameter, any %%1 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppWarn1( LPCWSTR pszMessage, const String& Insert1 )
{
    String Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_WARNING,
                          PXS_ERROR_TYPE_APPLICATION, 0, false, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an application warning message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      Insert1    - string insert parameter, any %%1 are replaced
//      Insert2    - string insert parameter, any %%2 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogAppWarn2( LPCWSTR pszMessage, const String& Insert1, const String& Insert2 )
{
    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_WARNING,
                          PXS_ERROR_TYPE_APPLICATION, 0, false, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a C.O.M. warning message to the log
//
//  Parameters:
//      hResult    - the C.O.M. error code
//      pszMessage - the message
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogComWarn( HRESULT hResult, LPCWSTR pszMessage )
{
    String Insert1, Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_WARNING,
                          PXS_ERROR_TYPE_COM,
                          (DWORD)hResult, true, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a C.O.M. warning message to the application's log
//
//  Parameters:
//      hResult    - the C.O.M. error code
//      pszMessage - the message with substitutable parameters
//      Insert1    - string insert parameter, any %%1 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogComWarn1( HRESULT hResult, LPCWSTR pszMessage, const String& Insert1 )
{
    String Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_WARNING,
                          PXS_ERROR_TYPE_COM,
                          (DWORD)hResult, true, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a configuration manager error message to the application's log
//
//  Parameters:
//      crResult    - the configuration manager error code
//      pszMessage - the message with substitutable parameters
//      Insert1    - string insert parameter, any %%1 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogConfigManError1( CONFIGRET crResult, LPCWSTR pszMessage, const String& Insert1 )
{
    String Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }

        // Filter success
        if ( crResult == CR_SUCCESS )
        {
            PXSLogWriteEntry( PXS_SEVERITY_INFORMATION,
                              PXS_ERROR_TYPE_APPLICATION,
                              0, false, pszMessage, Insert1, Insert2 );
        }
        else
        {
            PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                              PXS_ERROR_TYPE_CONFIG_MANAGER,
                              crResult, true, pszMessage, Insert1, Insert2 );
        }
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a DNS error message to the application's log
//
//  Parameters:
//      status    - the status code
//      pszMessage - the message with substitutable parameters
//      Insert1    - string insert parameter, any %%1 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogDnsError1( DNS_STATUS status, LPCWSTR pszMessage, const String& Insert1 )
{
    String Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                          PXS_ERROR_TYPE_DNS, (DWORD)status, true, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an exception to the application's log
//
//  Parameters:
//      e        - the exception
//      function - the logging function
//
//   Remarks:
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogException( const Exception& e, const char* function )
{
    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogException( nullptr, e, function );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write an exception to the application's log
//
//  Parameters:
//      pszPrePend - a message to add before the exception's message
//      e          - the exception
//      function   - the function logging the exception
//
//   Remarks:
//       Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogException( LPCWSTR message, const Exception& e, const char* function )
{
    String LogMessage, FunctionName, Insert1, Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }

        // Make the log entry
        if ( function )
        {
            FunctionName.SetAnsi( function );
            LogMessage  = L"Error logged by function '";
            LogMessage += FunctionName;
            LogMessage += L"'. ";
        }

        if ( message )
        {
            LogMessage += message;
            LogMessage += PXS_CHAR_SPACE;
        }
        LogMessage += e.GetMessage();
        PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                          e.GetErrorType(),
                          e.GetErrorCode(),
                          false, LogMessage.c_str(), Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }     // Ignore
}

//===============================================================================================//
//  Description:
//      Write a network error message to the application's log
//
//  Parameters:
//      status     - the network status error code
//      pszMessage - the message with substitutable parameters
//      Insert1    - string insert parameter, any %%1 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogNetError1( NET_API_STATUS status, LPCWSTR pszMessage, const String& Insert1 )
{
    String Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                          PXS_ERROR_TYPE_NETWORK, status, true, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a network error message to the application's log
//
//  Parameters:
//      status     - the network status error code
//      pszMessage - the message with substitutable parameters
//      Insert1    - string insert parameter, any %%1 are replaced
//      Insert2    - string insert parameter, any %%2 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogNetError2( NET_API_STATUS status,
                      LPCWSTR pszMessage, const String& Insert1, const String& Insert2 )
{
    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                          PXS_ERROR_TYPE_NETWORK, status, true, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a LSA error message to the application's log
//
//  Parameters:
//      status    - the LSA NTSTATUS error code
//      pszMessage - the message with substitutable parameters
//      pszInsert1 - optional insertion string where %%1 is present
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogNtStatusError1( NTSTATUS status, LPCWSTR pszMessage, const String& Insert1 )
{
    String Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                          PXS_ERROR_TYPE_NTSTATUS,
                          static_cast<DWORD>( status ), true, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a LSA warning message to the application's log
//
//  Parameters:
//      status    - the LSA NTSTATUS error code
//      pszMessage - the message with substitutable parameters
//      pszInsert1 - optional insertion string where %%1 is present
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogNtStatusWarn1( NTSTATUS status, LPCWSTR pszMessage, const String& Insert1 )
{
    String Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_WARNING,
                          PXS_ERROR_TYPE_NTSTATUS,
                          static_cast<DWORD>( status ), true, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a Windows socket error message to the application's log
//
//  Parameters:
//      wsaError   - the WSA error code
//      pszMessage - the message with substitutable parameters
//
//  Remarks
//      Does not throw
//      WSA errors are sytem errors of int type.
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogSocketError( int wsaError, LPCWSTR pszMessage )
{
    String Insert1, Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogSysError2( static_cast<DWORD>( wsaError ), pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a system error message to the application's log
//
//  Parameters:
//      pszMessage - the message
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogSysError( DWORD errorCode, LPCWSTR pszMessage )
{
    String Insert1, Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogSysError2( errorCode, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a system error message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      Insert1    - string insert parameter, any %%1 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogSysError1( DWORD errorCode,
                      LPCWSTR pszMessage, const String& Insert1 )
{
    String Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogSysError2( errorCode, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a system error message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      Insert1    - string insert parameter, any %%1 are replaced
//      Insert2    - string insert parameter, any %%2 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogSysError2( DWORD errorCode,
                      LPCWSTR pszMessage, const String& Insert1, const String& Insert2 )
{
    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }

        // Do not say error if no error code
        if ( errorCode == ERROR_SUCCESS )
        {
            PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                              PXS_ERROR_TYPE_APPLICATION,
                              0, false, pszMessage, Insert1, Insert2 );
        }
        else
        {
            PXSLogWriteEntry( PXS_SEVERITY_ERROR,
                              PXS_ERROR_TYPE_SYSTEM,
                              errorCode, true, pszMessage, Insert1, Insert2 );
        }
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a system informational message to the application's log
//
//  Parameters:
//      pszMessage - the message
//      Insert1    - string insert parameter, any %%1 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogSysInfo( DWORD errorCode, LPCWSTR pszMessage )
{
    String Insert1, Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_INFORMATION,
                          PXS_ERROR_TYPE_SYSTEM, errorCode, true, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a system informational message to the application's log
//
//  Parameters:
//      pszMessage - the message
//      Insert1    - string insert parameter, any %%1 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogSysInfo1(DWORD errorCode, LPCWSTR pszMessage, const String& Insert1)
{
    String Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_INFORMATION,
                          PXS_ERROR_TYPE_SYSTEM, errorCode, true, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a system informational message to the application's log
//
//  Parameters:
//      pszMessage - the message
//      Insert1    - string insert parameter, any %%1 are replaced
//      Insert2    - string insert parameter, any %%1 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogSysInfo2( DWORD errorCode,
                     LPCWSTR pszMessage, const String& Insert1, const String& Insert2 )
{
    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_INFORMATION,
                          PXS_ERROR_TYPE_SYSTEM, errorCode, true, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a system warning message to the log
//
//  Parameters:
//      pszMessage - the message
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogSysWarn( DWORD errorCode, LPCWSTR pszMessage )
{
    String Insert1, Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_WARNING,
                          PXS_ERROR_TYPE_SYSTEM, errorCode, true, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a system warning message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      Insert1    - string insert parameter, any %%1 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogSysWarn1( DWORD errorCode, LPCWSTR pszMessage, const String& Insert1 )
{
    String Insert2;

    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_WARNING,
                          PXS_ERROR_TYPE_SYSTEM, errorCode, true, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Write a system warning message to the application's log
//
//  Parameters:
//      pszMessage - the message with substitutable parameters
//      Insert1    - string insert parameter, any %%1 are replaced
//      Insert2    - string insert parameter, any %%2 are replaced
//
//  Remarks
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogSysWarn2( DWORD errorCode,
                     LPCWSTR pszMessage, const String& Insert1, const String& Insert2 )
{
    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        PXSLogWriteEntry( PXS_SEVERITY_WARNING,
                          PXS_ERROR_TYPE_SYSTEM, errorCode, true, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//        Write a log entry in the log file
//
//  Parameters:
//      severity      - defined constant of severity e.g. info, warning etc.
//      errorType     - type of error, e.g. system, network etc.
//      errorCode     - a numerical code of the error
//      translateError- flag to indicate want to translate the error code
//      pszMessage    - string message with optional substitution parameters
//      Insert1       - string insert parameter, any %%1 are replaced
//      Insert2       - string insert parameter, any %%2 are replaced
//
//  Remarks:
//      Does not throw exceptions
//
//  Returns:
//      void
//===============================================================================================//
void PXSLogWriteEntry( DWORD severity,
                       DWORD errorType,
                       DWORD errorCode,
                       bool translateError,
                       LPCWSTR pszMessage, const String& Insert1, const String& Insert2 )
{
    // Catch all
    try
    {
        if ( PXSIsAppLogging() == false )
        {
            return;
        }
        g_pApplication->WriteToAppLog( severity,
                                       errorType,
                                       errorCode,
                                       translateError, pszMessage, Insert1, Insert2 );
    }
    catch ( const Exception& )
    { }  // Ignore
}

//===============================================================================================//
//  Description:
//      Limit the value to the specified range
//
//  Parameters:
//      lower  - the lower bound
//      upper  - the upper bound
//      pValue - the value to limit, the value is updated in-place
//
//  Returns:
//      void
//===============================================================================================//
void PXSLimitInt( int lower, int upper, int* pValue )
{
    if ( pValue == nullptr )
    {
        throw NullException( L"pValue", __FUNCTION__ );
    }

    if ( *pValue < lower )
    {
        *pValue = lower;
    }
    else if ( *pValue > upper )
    {
        *pValue = upper;
    }
}

//===============================================================================================//
//  Description:
//      Limit the value to the specified range
//
//  Parameters:
//      lower  - the lower bound
//      upper  - the upper bound
//      pValue - the value to limit, the value is updated in-place
//
//  Returns:
//      void
//===============================================================================================//
void PXSLimitUInt32( DWORD lower, DWORD upper, DWORD* pValue )
{
    if ( pValue == nullptr )
    {
        throw NullException( L"pValue", __FUNCTION__ );
    }

    if ( *pValue < lower )
    {
        *pValue = lower;
    }
    else if ( *pValue > upper )
    {
        *pValue = upper;
    }
}

//===============================================================================================//
//  Description:
//      Limit the value to the specified range
//
//  Parameters:
//      lower  - the lower bound
//      upper  - the upper bound
//      pValue - the value to limit, the value is updated in-place
//
//  Returns:
//      void
//===============================================================================================//
void PXSLimitSizeT( size_t lower, size_t upper, size_t* pValue )
{
    if ( pValue == nullptr )
    {
        throw NullException( L"pValue", __FUNCTION__ );
    }

    if ( *pValue < lower )
    {
        *pValue = lower;
    }
    else if ( *pValue > upper )
    {
        *pValue = upper;
    }
}

//===============================================================================================//
//  Description:
//      Return the maximum of two DWORD values
//
//  Parameters:
//      x - value 1
//      y - value 2
//
//  Returns:
//      DWORD
//===============================================================================================//
DWORD PXSMaxUInt32( DWORD x, DWORD y )
{
    return ( (x > y) ? x : y );
}

//===============================================================================================//
//  Description:
//      Return the maximum of two int values
//
//  Parameters:
//      x - value 1
//      y - value 2
//
//  Returns:
//      int
//===============================================================================================//
int PXSMaxInt( int x, int y )
{
    return ( (x > y) ? x : y );
}

//===============================================================================================//
//  Description:
//      Return the maximum of two size_t values
//
//  Parameters:
//      x - value 1
//      y - value 2
//
//  Returns:
//      size_t
//===============================================================================================//
size_t PXSMaxSizeT( size_t x, size_t y )
{
     return ( (x > y) ? x : y );
}

//===============================================================================================//
//  Description:
//      Return the maximum of two time_t values
//
//  Parameters:
//      x - value 1
//      y - value 2
//
//  Returns:
//      time_t
//===============================================================================================//
time_t PXSMaxTimeT( time_t x, time_t y )
{
    return ( (x > y) ? x : y );
}

//===============================================================================================//
//  Description:
//      Return the minimum of two DWORD values
//
//  Parameters:
//      x - value 1
//      y - value 2
//
//  Returns:
//      DWORD
//===============================================================================================//
DWORD PXSMinUInt32( DWORD x, DWORD y )
{
    return ( (x < y) ? x : y );
}

//===============================================================================================//
//  Description:
//      Return the minimum of two int values
//
//  Parameters:
//      x - value 1
//      y - value 2
//
//  Returns:
//      int
//===============================================================================================//
int PXSMinInt( int x, int y )
{
    return ( (x < y) ? x : y );
}

//===============================================================================================//
//  Description:
//      Return the minimum of two size_t values
//
//  Parameters:
//      x - value 1
//      y - value 2
//
//  Returns:
//      size_t
//===============================================================================================//
size_t PXSMinSizeT( size_t x, size_t y )
{
    return ( (x < y) ? x : y );
}

//===============================================================================================//
//  Description:
//      Get the MD5 of the specified input
//
//  Parameters:
//      pData    - the data
//      numBytes - number of bytes of data
//      pMd5Hash - receives the hash
//
//  Remarks:
//      Called by the unhandled exception routine so will avoid throwing
//
//      See MSDN "Example C Program: Creating an MD-5 Hash From File Content"
//      Test cases from WikiPedia:
//      MD5("") = d41d8cd98f00b204e9800998ecf8427e
//      MD5("The quick brown fox jumps over the lazy dog")
//              = 9e107d9d372bb6826bd81d3542a419d6
//
//  Returns:
//      void
//===============================================================================================//
void PXSGetMd5Hash( BYTE* pData, DWORD numBytes, String* pMd5Hash )
{
    BYTE       bHash[ 16 ] = { 0 };    // Enough for an MD5
    DWORD      i = 0, dwDataLen = 0;
    Formatter  Format;
    HCRYPTPROV hProv = 0;
    HCRYPTHASH hHash = 0;

    if ( pMd5Hash == nullptr )
    {
        throw ParameterException( L"pMd5Hash", __FUNCTION__ );
    }
    pMd5Hash->Allocate( 16 );
    *pMd5Hash = PXS_STRING_EMPTY;

    if ( pData == nullptr )
    {
        return;
    }

    if ( CryptAcquireContext( &hProv, nullptr, nullptr, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT ) == 0 )
    {
        throw SystemException( GetLastError(), L"CryptAcquireContext", __FUNCTION__ );
    }

    if ( CryptCreateHash( hProv, CALG_MD5, 0, 0, &hHash ) == 0 )
    {
        CryptReleaseContext( hProv, 0 );
        throw SystemException(GetLastError(), L"CryptCreateHash", __FUNCTION__);
    }

    if ( CryptHashData( hHash, pData, numBytes, 0 ) == 0 )
    {
        CryptDestroyHash( hHash );
        CryptReleaseContext( hProv, 0 );
        throw SystemException(GetLastError(), L"CryptCreateHash", __FUNCTION__);
    }

    // Get the bits
    dwDataLen = ARRAYSIZE( bHash );
    if ( CryptGetHashParam( hHash, HP_HASHVAL, bHash, &dwDataLen, 0 ) == 0 )
    {
        CryptDestroyHash( hHash );
        CryptReleaseContext( hProv, 0 );
        throw SystemException( GetLastError(), L"CryptGetHashParam", __FUNCTION__ );
    }

    // Format to string, must use the value cbHash
    if ( dwDataLen > ARRAYSIZE( bHash ) )
    {
        throw BoundsException( L"dwDataLen", __FUNCTION__ );
    }

    for ( i = 0; i < dwDataLen; i++ )
    {
        *pMd5Hash += Format.UInt8Hex( bHash[ i ], false );
    }
    CryptDestroyHash( hHash );
    CryptReleaseContext( hProv, 0 );

    return;
}

//===============================================================================================//
//  Description:
//      Multiply two double values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      double
//===============================================================================================//
double PXSMultiplyDouble( double x, double y )
{
    if ( PXSIsNanDouble( x ) || PXSIsNanDouble( y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    double dblTemp1 = log( fabs( x ) );
    double dblTemp2 = log( fabs( y ) );

    if ( ( dblTemp1 + dblTemp2 ) > log( DBL_MAX ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }
    else if ( ( dblTemp1 + dblTemp2) < log( DBL_MIN ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return ( x * y );
}

//===============================================================================================//
//  Description:
//      Multiply two int values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      int
//===============================================================================================//
int PXSMultiplyInt32( int x, int y )
{
    // Test for positive/negative results
    if ( x > 0 )
    {
        if ( ( ( y > 0 ) && ( x > ( INT32_MAX / y ) ) ) ||
             ( ( y < 0 ) && ( x > ( INT32_MIN / y ) ) )  )
        {
            throw ArithmeticException( __FUNCTION__ );
        }
    }
    else
    {
        if ( ( ( y < 0 ) && ( x < ( INT32_MAX / y ) ) ) ||
             ( ( y > 0 ) && ( x < ( INT32_MIN / y ) ) )  )
        {
            throw ArithmeticException( __FUNCTION__ );
        }
    }

    return ( x * y );
}

//===============================================================================================//
//  Description:
//      Multiply two signed 64-bit values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Remarks:
//      Throws on numeric overflow.
//
//  Returns:
//      __int64
//===============================================================================================//
__int64 PXSMultiplyInt64( __int64 x, __int64 y )
{
    // Test for positive/negative results
    if ( x > 0 )
    {
        if ( ( ( y > 0 ) && ( x > ( INT64_MAX / y ) ) ) ||
             ( ( y < 0 ) && ( x > ( INT64_MIN / y ) ) )  )
        {
            throw ArithmeticException( __FUNCTION__ );
        }
    }
    else
    {
        if ( ( ( y < 0 ) && ( x < ( INT64_MAX / y ) ) ) ||
             ( ( y > 0 ) && ( x < ( INT64_MIN / y ) ) )  )
        {
            throw ArithmeticException( __FUNCTION__ );
        }
    }

    return ( x * y );
}

//===============================================================================================//
//  Description:
//      Multiply two long values values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Remarks:
//      Throws on numeric overflow.
//
//  Returns:
//      The product
//===============================================================================================//
long PXSMultiplyLong( long x, long y )
{
    // Test for positive/negative combinations
    if ( x > 0 )
    {
        if ( ( ( y > 0 ) && ( x > ( LONG_MAX / y ) ) ) ||
             ( ( y < 0 ) && ( x > ( LONG_MIN / y ) ) )  )
        {
            throw ArithmeticException( __FUNCTION__ );
        }
    }
    else
    {
        if ( ( ( y < 0 ) && ( x < ( LONG_MAX / y ) ) ) ||
             ( ( y > 0 ) && ( x < ( LONG_MIN / y ) ) )  )
        {
            throw ArithmeticException( __FUNCTION__ );
        }
    }

    return ( x * y );
}

//===============================================================================================//
//  Description:
//      Multiply two size_t values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      The product
//===============================================================================================//
size_t PXSMultiplySizeT( size_t x, size_t y )
{
    // Test for zero
    if ( y == 0 )
    {
        return 0;
    }

    if ( x > ( SIZE_MAX / y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return ( x * y );
}

//===============================================================================================//
//  Description:
//      Multiply two SQLLEN values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      SQLULEN
//===============================================================================================//
SQLLEN  PXSMultiplySqlLen( SQLLEN x, SQLLEN y )
{
    // Test for positive/negative combinations
    if ( x > 0 )
    {
        if ( ( ( y > 0 ) && ( x > ( PXS_SQLLEN_MAX / y ) ) ) ||
             ( ( y < 0 ) && ( x > ( PXS_SQLLEN_MIN / y ) ) )  )
        {
            throw ArithmeticException( __FUNCTION__ );
        }
    }
    else
    {
        if ( ( ( y < 0 ) && ( x < ( PXS_SQLLEN_MAX / y ) ) ) ||
             ( ( y > 0 ) && ( x < ( PXS_SQLLEN_MIN / y ) ) )  )
        {
            throw ArithmeticException( __FUNCTION__ );
        }
    }

    return ( x * y );
}

//===============================================================================================//
//  Description:
//      Multiply two time_t values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      The product
//===============================================================================================//
time_t PXSMultiplyTimeT( time_t x, time_t y )
{
    // Test for zero
    if ( y == 0 )
    {
        return 0;
    }

    if ( x > ( PXS_TIME_MAX / y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return ( x * y );
}

//===============================================================================================//
//  Description:
//      Multiply two SQLULEN values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      SQLULEN
//===============================================================================================//
SQLULEN PXSMultiplySqlULen( SQLULEN x, SQLULEN y )
{
    // Test for zero
    if ( y == 0 )
    {
        return 0;
    }

    if ( x > ( PXS_SQLULEN_MAX / y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return ( x * y );
}

//===============================================================================================//
//  Description:
//      Multiply two unsigned 8-bit values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      WORD
//===============================================================================================//
BYTE PXSMultiplyUInt8( BYTE x, BYTE y )
{
    // Test for zero
    if ( y == 0 )
    {
        return 0;
    }

    if ( x > ( UINT8_MAX / y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return static_cast< BYTE >( x * y );
}

//===============================================================================================//
//  Description:
//      Multiply two unsigned 16-bit values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      WORD
//===============================================================================================//
WORD PXSMultiplyUInt16( WORD x, WORD y )
{
    // Test for zero
    if ( y == 0 )
    {
        return 0;
    }

    if ( x > ( UINT16_MAX / y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return static_cast< WORD >( x * y );
}

//===============================================================================================//
//  Description:
//      Multiply two unsigned 32-bit values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      DWORD
//===============================================================================================//
DWORD PXSMultiplyUInt32( DWORD x, DWORD y )
{
    // Test for zero
    if ( y == 0 )
    {
        return 0;
    }

    if ( x > ( UINT32_MAX / y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return ( x * y );
}

//===============================================================================================//
//  Description:
//      Multiply two unsigned 64-bit values
//
//  Parameters:
//      x - the first value
//      y - the second value
//
//  Returns:
//      UINT64
//===============================================================================================//
UINT64 PXSMultiplyUInt64( UINT64 x, UINT64 y )
{
    // Test for zero
    if ( y == 0 )
    {
        return 0;
    }

    if ( x > ( UINT64_MAX / y ) )
    {
        throw ArithmeticException( __FUNCTION__ );
    }

    return ( x * y );
}

//===============================================================================================//
//  Description:
//      Hook procedure for the Open File Name dialog box
//
//  Parameters:
//      See MSDN OFNHookProc
//
//  Remarks:
//      Change the extension of file name in the dialog box to reflect the
//      user's choice of file type. The extensions are stored in the OFN
//      member lpstrFileTitle as a string of space separated values.
//
//  Returns:
//      zero if the default dialog box procedure processes the message other
//      wise non-zero
//===============================================================================================//
UINT_PTR CALLBACK PXSOFNHookProc( HWND hdlg, UINT uiMsg, WPARAM, LPARAM lParam )
{
    bool  isDirectory  = false;
    DWORD attributes   = 0;
    HWND  hWndParent   = nullptr;
    UINT_PTR  bReturn  = FALSE;
    OFNOTIFY* pNotify  = nullptr;
    OPENFILENAME* pOFN = nullptr;
    WCHAR wzPath[ MAX_PATH + 1 ] = { 0 };
    String FilePath, FileTitle, Drive, Dir, Fname, Ext;
    Directory DirObject;
    StringArray FileExtensions;

    if ( ( uiMsg == WM_NOTIFY ) && lParam )
    {
        pNotify = reinterpret_cast<OFNOTIFY*>( lParam );
        if ( pNotify->hdr.code == CDN_TYPECHANGE )
        {
            // User has changed the file type, get the extensions
            pOFN = pNotify->lpOFN;
            FileTitle.AppendChars( pOFN->lpstrFileTitle, pOFN->nMaxFileTitle );
            FileTitle.ToArray( PXS_CHAR_SPACE, &FileExtensions );
            if ( ( pOFN->nFilterIndex > 0 ) &&                     // one-based
                 ( pOFN->nFilterIndex <= FileExtensions.GetSize() ) )
            {
                // The controls are on a child window but need the dialog's
                // window handle
                hWndParent = GetParent( hdlg );

                // Replace the old file extension with the user's choice.
                // If no extension is present, splitpath will return the
                // directory name as the file name so test for that.
                CommDlg_OpenSave_GetFilePath( hWndParent, wzPath, ARRAYSIZE( wzPath ) );
                wzPath[ ARRAYSIZE( wzPath ) - 1 ] = PXS_CHAR_NULL;
                attributes  = GetFileAttributes( wzPath );
                isDirectory = ( attributes != INVALID_FILE_ATTRIBUTES ) &&
                              ( attributes & FILE_ATTRIBUTE_DIRECTORY );
                if ( isDirectory == false )
                {
                    FilePath = wzPath;
                    DirObject.SplitPath( FilePath, &Drive, &Dir, &Fname, &Ext );
                    Fname += PXS_STRING_DOT;
                    Fname += FileExtensions.Get( pOFN->nFilterIndex - 1 );
                    CommDlg_OpenSave_SetControlText( hWndParent, edt1, Fname.c_str() );
                }
            }
            bReturn = TRUE;     // Handled
        }
    }

    return bReturn;
}

//===============================================================================================//
//  Description:
//      Prompt for a file path using the Open File Name dialogue
//
//  Parameters:
//      hWndOwner       - owner window for the Open File Name dialogue
//      open            - flag to indicate if the dialogue is for Open or Save
//      promptOverwrite - prompt before overwriting a file
//      pszFilter       - the file filter
//      puFilterIndex   - the index of the filter to initially select, on output
//                        receives the filter index chosen
//      pszFileName     - the initial file name to show on the dialogue
//      pszExtensions   - Space separated values to append to file name when
//                        the user changes the file type on the dialog
//      pFilePath       - receives the selected file name
//
//  Returns:
//      true if the user selected a file, otherwise false
//===============================================================================================//
bool PXSPromptForFilePath( HWND    hWndOwner,
                           bool    open,
                           bool    promptOverwrite,
                           LPCWSTR pszFilter,
                           DWORD*  puFilterIndex,
                           LPCWSTR pszFileName, LPCWSTR pszExtensions, String* pFilePath )
{
    File   FileObject;
    BOOL   result  = FALSE;
    HWND   hWndMainFrame;
    DWORD  error = 0;
    wchar_t szFilePath[ MAX_PATH + 1 ]   = { 0 };
    wchar_t wzExtensions[ MAX_PATH + 1 ] = { 0 };
    String  InitialDir, FileName, Drive, DirPath, Fname, Extension;
    Directory DirectoryObject;
    OPENFILENAMEW ofn;

    if ( pFilePath == nullptr )
    {
        throw ParameterException( L"pFilePath", __FUNCTION__ );
    }
    *pFilePath = PXS_STRING_EMPTY;

    // Need the filter index so can return which filter was chosen
    if ( puFilterIndex == nullptr )
    {
        throw ParameterException( L"puFilterIndex", __FUNCTION__ );
    }

    // Set the initial directory, if none use the personal folder
    if ( g_pApplication )
    {
        g_pApplication->GetLastAccessedDirectory( &InitialDir );
    }
    if ( InitialDir.IsEmpty() )
    {
        DirectoryObject.GetSpecialDirectory( CSIDL_PERSONAL, &InitialDir );
    }

    FileName = pszFileName;
    FileName.Trim();
    if ( FileObject.IsValidFileName( FileName ) )
    {
        StringCchCopy( szFilePath, ARRAYSIZE( szFilePath ), FileName.c_str() );
    }

    // Reset and populate the OFN structure
    memset( &ofn, 0, sizeof ( ofn ) );
    ofn.lStructSize     = sizeof ( ofn );
    ofn.hwndOwner       = hWndOwner;
    ofn.lpstrFilter     = pszFilter;
    ofn.nFilterIndex    = *puFilterIndex;
    ofn.lpstrFile       = szFilePath;
    ofn.nMaxFile        = MAX_PATH;
    ofn.lpstrInitialDir = InitialDir.c_str();
    ofn.lpstrTitle      = nullptr;
    ofn.Flags           = OFN_HIDEREADONLY;

    // Configure for file type changes so the hook will be called in order
    // to adjust the file extensions
    if ( pszExtensions )
    {
        ofn.Flags |= ( OFN_EXPLORER | OFN_ENABLEHOOK );
        ofn.lpfnHook      = PXSOFNHookProc;
        ofn.nMaxFileTitle = static_cast<DWORD>(
                                        0xffffffff & wcslen( pszExtensions ) );
        StringCchCopyN( wzExtensions,
                        ARRAYSIZE( wzExtensions ), pszExtensions, ofn.nMaxFileTitle );
        ofn.lpstrFileTitle = wzExtensions;
    }

    // Show the dialogue box
    if ( open )
    {
        ofn.Flags |= OFN_PATHMUSTEXIST;
        result     = GetOpenFileName( &ofn );
    }
    else
    {
        if ( promptOverwrite )
        {
            ofn.Flags |= OFN_OVERWRITEPROMPT;
        }
        result = GetSaveFileName( &ofn );
    }
    szFilePath[ ARRAYSIZE( szFilePath ) - 1 ] = PXS_CHAR_NULL;

    // Test for cancel or error
    if ( result == 0 )
    {
        error = CommDlgExtendedError();
        if ( error )
        {
            throw Exception( PXS_ERROR_TYPE_COMMON_DIALOG,
                             error, L"GetSaveFileName", __FUNCTION__ );
        }
        return false;   // = user cancel
    }
    *puFilterIndex = ofn.nFilterIndex;
    szFilePath[ ARRAYSIZE( szFilePath ) - 1 ] = PXS_CHAR_NULL;
    *pFilePath = szFilePath;
    pFilePath->Trim();

    // Store the last accessed directory
    if ( g_pApplication )
    {
        DirectoryObject.SplitPath( *pFilePath,
                                   &Drive, &DirPath, &Fname, &Extension );
        if ( DirectoryObject.IsOnLocalDrive( DirPath ) )
        {
            g_pApplication->SetLastAccessedDirectory( DirPath );
        }

        hWndMainFrame = g_pApplication->GetHwndMainFrame();
        if ( hWndMainFrame )
        {
            RedrawWindow( hWndMainFrame,
                          nullptr,
                          nullptr,
                          RDW_ERASE | RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN );
        }
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Callback for qsort for a case insensitive ascending string comparison
//
//  Parameters:
//      pArg1 - pointer to string 1
//      pArg2 - pointer to string 2
//
//  Returns:
//      > 0  if arg1 greater than arg2
//      = 0  if arg1 equivalent to arg2
//      < 0  if arg1 less than arg2
//===============================================================================================//
int PXSQSortStringAscending( const void* pArg1, const void* pArg2 )
{
    return PXSCompareString( static_cast<LPCWSTR>( pArg1 ),
                             static_cast<LPCWSTR>( pArg2 ), false );
}

//===============================================================================================//
//  Description:
//      Callback for qsort for a case insensitive descending string comparison
//
//  Parameters:
//      pArg1 - pointer to string 1
//      pArg2 - pointer to string 2
//
//  Returns:
//      < 0  if arg1 less than arg2
//      = 0  if arg1 equivalent to arg2
//      > 0  if arg1 greater than arg2
//===============================================================================================//
int PXSQSortStringDescending( const void* pArg1, const void* pArg2 )
{
    return -PXSQSortStringAscending( pArg1, pArg2 );
}

//===============================================================================================//
//  Description:
//      Callback for qsort for a case insensitive ascending sort of a two
//      PXS_TYPE_STRING_STRING structures. Only sorts on the first string.
//
//  Parameters:
//      pArg1 - pointer to string 1
//      pArg2 - pointer to string 2
//
//  Returns:
//      > 0  if arg1 greater than arg2
//      = 0  if arg1 equivalent to arg2
//      < 0  if arg1 less than arg2
//===============================================================================================//
int PXSQSortStringStringAscending( const void* pArg1, const void* pArg2 )
{
    LPCWSTR pszOne  = nullptr;
    LPCWSTR pszTwo  = nullptr;

    // Compare on first string (pszString1) of each structure
    if ( pArg1 )
    {
        pszOne = ( (const PXS_TYPE_STRING_STRING*)pArg1 )->pszString1;
    }

    if ( pArg2 )
    {
        pszTwo = ( (const PXS_TYPE_STRING_STRING*)pArg2 )->pszString1;
    }

    return PXSCompareString( pszOne, pszTwo, false );
}

//===============================================================================================//
//  Description:
//      Colour shift the specified bitmap by swapping R->G, G->B and B->R
//
//  Parameters:
//      hBitmap - handle to the bitmap
//
//  Remarks:
//      Uses GetPixel and SetPixel so use sparingly
//
//  Returns:
//      void
//===============================================================================================//
void PXSBitmapRgbToGrb( HBITMAP hBitmap )
{
    long     i = 0, j = 0;
    HDC      hCompatibleDC;
    BITMAP   bitmap;
    HGDIOBJ  hOldBitmap;
    COLORREF oldPixel  = CLR_INVALID, newPixel = CLR_INVALID;

    if ( hBitmap == nullptr )
    {
        return;     // Nothing to do
    }

    memset( &bitmap, 0, sizeof ( bitmap ) );
    if ( GetObject( hBitmap, sizeof ( bitmap ), &bitmap ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetObject", __FUNCTION__ );
    }

    hCompatibleDC = CreateCompatibleDC( nullptr );
    if ( hCompatibleDC == nullptr )
    {
        throw SystemException( GetLastError(), L"CreateCompatibleDC", __FUNCTION__ );
    }
    hOldBitmap = SelectObject( hCompatibleDC, hBitmap );

    // Replace each pixel's RGB value with GBR
    for ( i = 0; i < bitmap.bmWidth; i++ )
    {
        for ( j = 0; j < bitmap.bmHeight; j++ )
        {
            oldPixel = GetPixel( hCompatibleDC, i, j );
            newPixel = static_cast<COLORREF>( GetGValue( oldPixel ) |
                                             (GetBValue( oldPixel ) <<  8) |
                                             (GetRValue( oldPixel ) << 16) );
            SetPixelV( hCompatibleDC, i, j, newPixel );
        }
    }

    // Reset
    SelectObject( hCompatibleDC, hOldBitmap );
    DeleteDC( hCompatibleDC );
}

//===============================================================================================//
//  Description:
//      Replace a colour in a bitmap with the specified new one.
//
//  Parameters:
//      hBitmap   - handle to the bitmap
//      oldColour - the old colour
//      newColour - the new colour
//
//  Remarks:
//      For use on small images as its a pixel by pixel swap.
//
//  Returns:
//      void
//===============================================================================================//
void PXSReplaceBitmapColour( HBITMAP hBitmap, COLORREF oldColour, COLORREF newColour )
{
    long     i = 0, j = 0;
    HDC      hCompatibleDC = nullptr;
    BITMAP   bitMap;
    HGDIOBJ  hOldBitmap = nullptr;

    // Inputs
    if ( ( hBitmap == nullptr       ) ||
         ( oldColour == CLR_INVALID ) ||
         ( newColour == CLR_INVALID )  )
    {
        return;     // Nothing to do or bad parameter
    }

    memset( &bitMap, 0, sizeof ( bitMap ) );
    if ( GetObject( hBitmap, sizeof ( bitMap ), &bitMap ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetObject", __FUNCTION__ );
    }

    // Get the nearest colours
    hCompatibleDC = CreateCompatibleDC( nullptr );
    if ( hCompatibleDC == nullptr )
    {
        throw SystemException( GetLastError(), L"CreateCompatibleDC", __FUNCTION__ );
    }
    hOldBitmap = SelectObject( hCompatibleDC, hBitmap );

    // Pixel by pixel replace
    for ( i = 0; i < bitMap.bmWidth; i++)
    {
        for ( j = 0; j < bitMap.bmHeight; j++)
        {
            if ( oldColour == GetPixel( hCompatibleDC, i, j ) )
            {
                SetPixelV( hCompatibleDC, i, j, newColour );
            }
        }
    }

    // Restore
    if ( hOldBitmap )
    {
        SelectObject( hCompatibleDC, hOldBitmap  );
    }
    DeleteDC( hCompatibleDC );
}

//===============================================================================================//
//  Description:
//      Select a font using the standard windows dialogue
//
//  Parameters:
//      hWndOwner   - handle to the dialogue's parent windows
//      pFontObject - on input specified the settings for the CHOOSEFONT
//                    dialog. On output has the user's choices
//
//  Returns:
//      true if the user picked a font
//===============================================================================================//
bool PXSShowChooseFontDialog( HWND hWndOwner, Font* pFontObject )
{
    bool   success = false;
    DWORD  error   = 0;
    LOGFONT    lf;
    CHOOSEFONT cf;

    if ( pFontObject == nullptr )
    {
        throw ParameterException( L"pFontObject", __FUNCTION__ );
    }
    memset( &lf, 0, sizeof ( lf ) );
    pFontObject->GetLogFont( &lf );

    // Show the dialogue
    memset( &cf, 0, sizeof ( cf ) );
    cf.lStructSize = sizeof ( cf );
    cf.hwndOwner   = hWndOwner;
    cf.lpLogFont   = &lf;
    cf.Flags       = CF_SCREENFONTS | CF_INITTOLOGFONTSTRUCT | CF_LIMITSIZE | CF_SCALABLEONLY;
    cf.nSizeMin    = 8;
    cf.nSizeMax    = 36;
    if ( ChooseFont( &cf ) )
    {
        pFontObject->SetLogFont( &lf );
        pFontObject->Create();
        success = true;
    }
    else
    {
        // CommDlgExtendedError returns zero if the user presses cancel
        error = CommDlgExtendedError();
        if ( error )
        {
            throw Exception( PXS_ERROR_TYPE_COMMON_DIALOG, error, L"ChooseFont", __FUNCTION__ );
        }
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Show an dialog message
//
//  Parameters:
//      e         - the exception
//      hWndOwner - the dialog's owner
//
//  Returns:
//      void
//===============================================================================================//
void PXSShowExceptionDialog( const Exception& e, HWND hWndOwner )
{
    size_t    i = 0, numElements = 0;
    String    Message, Caption, CrashString, ApplicationName;
    Formatter Format;
    StringArray   DiagnosticData;
    MessageDialog ExceptionDialog;

    PXSLogException( e, __FUNCTION__ );

    // Must have an owner window that is visible
    if ( ( hWndOwner == nullptr ) ||
         ( IsWindowVisible( hWndOwner ) == 0 ) )
    {
        return;
    }

    // Reset the cursor in case the user has to press buttons
    SetCursor( LoadCursor( nullptr, IDC_ARROW ) );
    MessageBeep( MB_OK );
    try
    {
        Message  = e.GetMessage();
        Message += PXS_STRING_CRLF;
        if ( e.GetErrorType() == PXS_ERROR_TYPE_EXCEPTION )
        {
            // Show as a two column table
            try
            {
                if ( g_pApplication )
                {
                    g_pApplication->GetDiagnosticData( true, &DiagnosticData );
                }
                numElements = DiagnosticData.GetSize();
                for ( i = 0; i < numElements; i++ )
                {
                    Message += DiagnosticData.Get( i );
                    Message += PXS_STRING_CRLF;
                }
            }
            catch ( const Exception& eDiagnostic )
            {
                PXSLogException( eDiagnostic, __FUNCTION__ );
            }
            Message += PXS_STRING_CRLF;
        }

        // Stack trace
        CrashString = e.GetCrashString();
        if ( CrashString.GetLength() > 0 )
        {
            Message += CrashString;
            Message += PXS_STRING_CRLF;
        }

        // Set the dialogue's properties and show it
        PXSGetApplicationName( &ApplicationName );
        ExceptionDialog.SetTitle( ApplicationName );
        ExceptionDialog.SetIsError( true );
        ExceptionDialog.SetMessage( Message );
        if ( PXS_ERROR_TYPE_EXCEPTION == e.GetErrorType() )
        {
            ExceptionDialog.SetSize( 600, 425 );
        }
        else
        {
            ExceptionDialog.SetSize( 500, 300 );
        }
        ExceptionDialog.Create( hWndOwner );
    }
    catch ( const Exception& )
    {
        // Want to show the input exception
        MessageBox( hWndOwner, e.GetMessage().c_str(), nullptr, MB_OK | MB_ICONEXCLAMATION );
    }
}

//===============================================================================================//
//  Description:
//     Sort a Name-Value array on the names
//
//  Parameters:
//      pNameValues - the array to sort
//
//  Returns:
//      void
//===============================================================================================//
void PXSSortNameValueArray( TArray< NameValue >* pNameValues )
{
    size_t    i = 0, numElements = 0;
    String    Name, Value;
    NameValue Element;
    TArray< NameValue > SortedArray;
    PXS_TYPE_STRING_STRING* pArray = nullptr;

    if ( pNameValues == nullptr )
    {
        return;     // Nothing to do
    }

    numElements = pNameValues->GetSize();
    if ( numElements < 1 )
    {
        return;     // Nothing to do
    }
    pArray = new PXS_TYPE_STRING_STRING[ numElements ];
    if ( pArray == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }

    try
    {
        for ( i = 0; i < numElements; i++ )
        {
            // Get the first field of the record
            pArray[ i ].pszString1 = pNameValues->Get( i ).GetName().c_str();
            pArray[ i ].pszString2 = pNameValues->Get( i ).GetValue().c_str();
        }
        qsort( pArray,
               numElements, sizeof ( PXS_TYPE_STRING_STRING ), PXSQSortStringStringAscending );

        // Replace
        SortedArray.SetSize( numElements );
        for ( i = 0; i < numElements; i++ )
        {
            Name  = pArray[ i ].pszString1;     // false positive C6385 with /analyze
            Value = pArray[ i ].pszString2;
            Element.SetNameValue( Name, Value );
            SortedArray.Set( i, Element );
        }
        *pNameValues = SortedArray;
    }
    catch ( const Exception& )
    {
        delete [] pArray;
        throw;
    }
    delete [] pArray;
}

//===============================================================================================//
//  Description:
//      Exit the program on an unexpected error, typically a
//      structured exception. Almost always an access violation.
//
//  Parameters:
//      pException - a pointer to an exception structure
//
//  Remarks
//      WINAPI == CALLBACK == __stdcall which is required
//      by SetUnhandledExceptionFilter
//
//  Returns:
//      A named EXCEPTION constant, one of:
//          EXCEPTION_EXECUTE_HANDLER       1
//          EXCEPTION_CONTINUE_EXECUTION    0
//          EXCEPTION_CONTINUE_SEARCH      -1
//===============================================================================================//
// Remove MSVC compiler warning of unreachable code because going to
// throw an exception but the function needs a return specifier
#pragma warning( push )
#pragma warning ( disable : 4702 )
LONG WINAPI PXSShowUnhandledExceptionDialog( EXCEPTION_POINTERS* pException )
{
    void*   exceptionAddress = nullptr;
    DWORD   exceptionCode    = 0;
    HWND    hWndMainFrame = nullptr;
    String  ErrorMessage;
    wchar_t szShortMessage[ 128 ] = { 0 };   // Enough for a short message
    UnhandledExceptionDialog ExceptionDialog;

    if ( g_pApplication == nullptr )
    {
        return EXCEPTION_EXECUTE_HANDLER;
    }
    hWndMainFrame = g_pApplication->GetHwndMainFrame();
    SetCursor( LoadCursor( nullptr, IDC_ARROW ) );

    try
    {
        // Make sure only get here once
        if ( 1 != InterlockedIncrement( &m_lExceptionCount ) )
        {
            g_pApplication->GetResourceString( PXS_IDS_118_MULTI_FATAL_ERRORS, &ErrorMessage );
            if ( hWndMainFrame )
            {
                MessageBox( hWndMainFrame, ErrorMessage.c_str(), nullptr, MB_OK | MB_ICONERROR );
            }
            PXSLogAppError( ErrorMessage.c_str() );
            g_pApplication->LogFlush();
            ExitProcess( 3 );       // same as abort
        }

        // Log it
        exceptionCode = 1;  // Say 1, this means failure.
        if ( pException && pException->ExceptionRecord )
        {
            exceptionCode    = pException->ExceptionRecord->ExceptionCode;
            exceptionAddress = pException->ExceptionRecord->ExceptionAddress;
            StringCchPrintf( szShortMessage,
                             ARRAYSIZE( szShortMessage ),
                             L"Fatal error, code=0x%08X, address=%p.",
                             exceptionCode, exceptionAddress );
            PXSLogAppError( szShortMessage );
        }

        // Construct an exception object and get its message
        UnhandledException eUnhandled( pException );
        ErrorMessage = eUnhandled.GetMessage();
        PXSLogAppError( ErrorMessage.c_str() );
        g_pApplication->LogFlush();

        // Show it, even if frame is invisible
        MessageBeep( MB_OK );
        ExceptionDialog.SetExceptionMessage( ErrorMessage );
        ExceptionDialog.Create( hWndMainFrame );
    }
    catch ( const Exception& )
    {
        MessageBox( hWndMainFrame, szShortMessage, nullptr, MB_OK | MB_ICONERROR );
    }

     // Exit
    if ( hWndMainFrame )
    {
        PostQuitMessage( static_cast<int>( exceptionCode ) );
    }
    else
    {
        ExitProcess( exceptionCode );
    }

    // Can't get here but need to return a value.
    return EXCEPTION_EXECUTE_HANDLER;
}
#pragma warning( pop )  // Restore the warnings

//===============================================================================================//
//  Description:
//      Swap the WORD order of the specified DWORD
//
//  Parameters:
//      value - the value
//
//  Returns:
//      DWORD
//===============================================================================================//
DWORD PXSSwapWords( DWORD value )
{
    DWORD low  = ( 0xFFFF & value );
    DWORD high = ( 0xFFFF & ( value >> 16 ) );

    return ( ( ( low ) << 16 ) + high );
}

//===============================================================================================//
//  Description:
//      Handle a call to terminate, i.e. abort
//
//  Parameters:
//      None
//
//  Remarks:
//      set_terminate has an exception specifier of throw()
//
//  Returns:
//      void
//===============================================================================================//
void PXSTerminateHandler( void )
{
    HWND    hWndMainFrame;
    String  ApplicationName;
    LPCWSTR ERROR_MESSAGE = L"The application is terminating unexpectedly.";

    PXSLogSysError( ERROR_PROCESS_ABORTED, ERROR_MESSAGE );
    if ( g_pApplication == nullptr )
    {
        return;
    }
    PXSLogSysError( ERROR_PROCESS_ABORTED, L"PXSTerminateHandler" );
    g_pApplication->LogFlush();

    // If have a window then show a message and exit
    hWndMainFrame = g_pApplication->GetHwndMainFrame();
    if ( hWndMainFrame )
    {
        g_pApplication->GetApplicationName( &ApplicationName );
        MessageBeep( MB_OK );
        MessageBox( hWndMainFrame, ERROR_MESSAGE, ApplicationName.c_str(), MB_OK | MB_ICONERROR );
        PostQuitMessage( 3 );      // Say 3 just like abort
    }
    else
    {
        // No main frame, say 3 just like abort
        ExitProcess( 3 );
    }
}

//===============================================================================================//
//  Description:
//      Convert the specified root path to a string
//
//  Parameters:
//      driveType    - result of GetDriveType
//      pTranslation - receives the translation
//
//  Returns:
//      void
//===============================================================================================//
void PXSTranslateDriveType( UINT driveType, String* pTranslation )
{
    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }
    *pTranslation = PXS_STRING_EMPTY;

    switch ( driveType )
    {
        default:
        case DRIVE_UNKNOWN:
            *pTranslation = L"Unknown";
            break;

        case DRIVE_NO_ROOT_DIR:
            *pTranslation = L"No root directory";
            break;

        case DRIVE_REMOVABLE:
            *pTranslation = L"Removable";
            break;

        case DRIVE_FIXED:
            *pTranslation = L"Fixed";
            break;

        case DRIVE_REMOTE:
            *pTranslation = L"Remote";
            break;

        case DRIVE_CDROM:
            *pTranslation = L"CD-ROM";
            break;

        case DRIVE_RAMDISK:
            *pTranslation = L"RAM disk.";
            break;
    }
}

//===============================================================================================//
//  Description:
//      Remove the opening and closing double quote characters (")
//
//  Parameters:
//      pString - pointer to the string to unquote
//
//  Remarks:
//      pString must be of the form "xxx", ie both opening and closing quotes.
//
//  Returns:
//      void
//===============================================================================================//
void PXSUnQuoteString( String* pString )
{
    size_t length;
    String NewString;

    if ( pString == nullptr )
    {
        return;  // Nothing to do
    }

    length = pString->GetLength();
    if ( length < 2 )
    {
        return;  // Nothing to do
    }

    if ( ( pString->CharAt( 0 )          == PXS_CHAR_QUOTE ) &&
         ( pString->CharAt( length - 1 ) == PXS_CHAR_QUOTE )  )
    {
        pString->SubString( 1, length - 2, &NewString );
        *pString = NewString;
    }
}

//===============================================================================================//
//  Description:
//      Convert the specified VARIANT to a string. Handles arrays, i.e. VT_ARRAY
//
//  Parameters:
//      pVariant - pointer to the variant
//      pValue   - receives the string value
//
//  Remarks
//      For arrays, the output is a string of spce separated values.
//
//  Returns:
//      void
//===============================================================================================//
void PXSVariantToString( const VARIANT* pVariant, String* pValue )
{
    long      i = 0, lBound  = 0, uBound = 0;
    HRESULT   hResult = S_OK;
    VARIANT   Element;
    String    Value;

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    pValue->Zero();

    if ( pVariant == nullptr )
    {
        *pValue = L"<null>";
        return;
    }

    // Test for array
    if ( pVariant->vt & VT_ARRAY )
    {
        hResult = SafeArrayGetLBound( pVariant->parray, 1, &lBound );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"SafeArrayGetLBound", __FUNCTION__ );
        }

        hResult = SafeArrayGetUBound( pVariant->parray, 1, &uBound );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"SafeArrayGetUBound", __FUNCTION__ );
        }

        // Append each value to the output string, space separated
        for ( i = lBound; i <= uBound; i++ )
        {
            // Get the value as a VARIANT, will use the lVal member for SafeArrayGetElement
            // as the members are a union. VARTYPE is a WORD.
            VariantInit( &Element );
            Element.vt = static_cast< VARTYPE >( 0xFFFF & ( pVariant->vt ^ VT_ARRAY ) );
            hResult    = SafeArrayGetElement( pVariant->parray, &i, &Element.lVal );
            if ( FAILED( hResult ) )
            {
                throw ComException( hResult, L"SafeArrayGetElement", __FUNCTION__ );
            }

            try
            {
                Value.Zero();
                PXSVariantValueToString( &Element, &Value );
            }
            catch( Exception& )
            {
                VariantClear( &Element );
                throw;
            }
            VariantClear( &Element );

            pValue->AppendString( Value );
            if ( i != uBound )
            {
                pValue->AppendChar( PXS_CHAR_SPACE );
            }
        }
    }
    else
    {
        PXSVariantValueToString( pVariant, pValue );
    }
}

//===============================================================================================//
//  Description:
//      Convert the specified VARIANT to a string
//
//  Parameters:
//      pVariant - pointer to the variant
//      pValue   - receives the string value
//
//  Remarks
//      Only converting the common data types.
//
//  Returns:
//      void
//===============================================================================================//
void PXSVariantValueToString( const VARIANT* pVariant, String* pValue )
{
    char      szChar[ 8 ] = { 0 };
    String    ErrorMessage;
    Formatter Format;

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    pValue->Zero();

    if ( pVariant == nullptr )
    {
        *pValue = L"<null>";
        return;
    }

    if ( VT_BYREF & pVariant->vt )
    {
        *pValue = L"<reference>";
        return;
    }

    if ( VT_ARRAY & pVariant->vt )
    {
        *pValue = L"<array>";
        return;
    }

    // Check most common data types
    *pValue = PXS_STRING_EMPTY;
    switch ( pVariant->vt )
    {
        default:
            ErrorMessage = Format.StringInt32( L"VARTYPE= %%1.", pVariant->vt );
            throw SystemException( ERROR_INVALID_DATATYPE, ErrorMessage.c_str(), __FUNCTION__ );

        case VT_NULL:
            *pValue = L"<null>";
            break;

        case VT_EMPTY:
            *pValue = L"<empty>";
            break;

        case VT_BOOL:
            if ( pVariant->boolVal == VARIANT_TRUE )
            {
                *pValue = PXS_STRING_ONE;
            }
            else
            {
                *pValue = PXS_STRING_ZERO;
            }
            break;

        case VT_I1:
            szChar[ 0 ] = pVariant->cVal;
            szChar[ 1 ] = 0x00;
            pValue->SetAnsi( szChar );
            break;

        case VT_UI1:
            *pValue = Format.UInt8( pVariant->bVal );
            break;

        case VT_I2:
            *pValue = Format.Int32( pVariant->iVal );
            break;

        case VT_UI2:
            *pValue = Format.UInt16( pVariant->uiVal );
            break;

        case VT_I4:
            *pValue = Format.Int32( pVariant->lVal );
            break;

        case VT_UI4:
            *pValue = Format.UInt32( pVariant->ulVal );
            break;

        case VT_I8:
            *pValue = Format.Int64( pVariant->llVal );
            break;

        case VT_UI8:
            *pValue = Format.UInt64( pVariant->ullVal );
            break;

        case VT_R4:
            *pValue = Format.Float( pVariant->fltVal );
            break;

        case VT_R8:
            *pValue = Format.Double( pVariant->dblVal );
            break;

        case VT_BSTR:
            *pValue = pVariant->bstrVal;
            break;

        case VT_DATE:
            *pValue = Format.OleTimeInIsoFormat( pVariant->date );
            break;
    }
}

//===============================================================================================//
//  Description:
//      Extract the 32-bit low-order part of a WPARAM
//
//  Parameters:
//      value - the value
//
//  Remarks:
//      WPARAM is unsigned int on 32-bit and unsigned __int64 on 64-bit Windowa
//
//  Returns:
//      DWORD
//===============================================================================================//
DWORD PXSWParamToLowUInt32( WPARAM value )
{
    return static_cast<DWORD>( 0xFFFFFFFF & value );
}

//===============================================================================================//
//  Description:
//      Write the specified unhandle exception to the application's log. The
//      logger must be on otherwise has no effect.
//
//  Parameters:
//      pException - pointer to the variant
//
//  Remarks:
//      This is a filter handler for the __except clause.
//
//  Returns:
//      EXCEPTION_EXECUTE_HANDLER
//===============================================================================================//
LONG WINAPI PXSWriteUnhandledExceptionToLog( EXCEPTION_POINTERS* pException )
{
    // Make sure only get here once
    if ( 1 != InterlockedIncrement( &m_lExceptionCount ) )
    {
        return EXCEPTION_EXECUTE_HANDLER;
    }

    if ( g_pApplication == nullptr )
    {
        return EXCEPTION_EXECUTE_HANDLER;
    }

    if ( pException == nullptr )
    {
        return EXCEPTION_EXECUTE_HANDLER;
    }

    try
    {
        UnhandledException eUnhandled( pException );
        PXSLogAppError( eUnhandled.GetMessage().c_str() );
        g_pApplication->LogFlush();
    }
    catch ( const Exception& )
    { }     // Ignore, not much can be done

    return EXCEPTION_EXECUTE_HANDLER;
}
