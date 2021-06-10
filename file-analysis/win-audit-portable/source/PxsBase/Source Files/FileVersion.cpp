///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Binary File Version Class Implementation
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
#include "PxsBase/Header Files/FileVersion.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StringT.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Constructs a new File object
FileVersion::FileVersion()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
FileVersion::~FileVersion()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator - not allowed so no implementation

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get a binary file's version information
//
//  Parameters:
//      Filename         - path of file
//      pManufacturerName- receives the embedded manufacturer's name
//      pVersionString   - receives the embedded version
//      pDescription     - receives the embedded description
//
//  Remarks:
//      Not all files haver version information so will not raise
//      errors if no data found.
//
//  Returns:
//      Void
//===============================================================================================//
void FileVersion::GetVersion( const String& Filename,
                              String* pManufacturerName,
                              String* pVersionString, String* pDescription )
{
    BYTE*     pData    = nullptr;
    DWORD     dwHandle = 0, dwLen = 0, codePage = 0;
    String    SubBlock;
    Formatter Format;
    AllocateBytes AllocBytes;

    if ( Filename.IsEmpty() )
    {
        throw ParameterException( L"Filename", __FUNCTION__ );
    }

    if ( ( pManufacturerName == nullptr ) ||
         ( pVersionString    == nullptr ) ||
         ( pDescription      == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    *pManufacturerName = PXS_STRING_EMPTY;
    *pVersionString    = PXS_STRING_EMPTY;
    *pDescription      = PXS_STRING_EMPTY;

    // Get size of data in version information for buffer alloc.
    dwLen = GetFileVersionInfoSize( Filename.c_str(), &dwHandle );
    if ( ( dwLen == 0 ) ||
         ( dwLen >  0x7FFF8 ) )  // According to the documentation
    {
        return;					// No data or bad data
    }

	dwHandle = 0;				// dwHandle is not used, should be zero
    pData    = AllocBytes.New( dwLen );
    if ( GetFileVersionInfo( Filename.c_str(), dwHandle, dwLen, pData ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetFileVersionInfo", __FUNCTION__ );
    }
    SubBlock.Allocate( 128 );

    // The file version information may exist in multiple languages. The path
    // to each one is of the form \StringFileInfo\lang-codepage\FileVersion.
    // So first read the translation table which is an array of DWORDs.
    // Will use the first one however, the DWORD list is unordered so the
    // its is not necessarily the user's default language. Also, the DWORD
    // is in codepage-lang order no need to swap the WORDs
    SubBlock = L"\\VarFileInfo\\Translation";
    GetVersionInfoDWord( pData, dwLen, SubBlock.c_str(), &codePage );
    codePage = PXSSwapWords( codePage );
    if ( codePage == 0 )
    {
        codePage = 0x040904B0;      // U.S. English (=0x409) and
    }                               // Unicode (= 0x04B0)

    // Company name, format is \\StringFileInfo\\%08X\\CompanyName
    SubBlock  = L"\\StringFileInfo\\";
    SubBlock += Format.UInt32Hex( codePage, false );
    SubBlock += L"\\CompanyName";
    GetVersionInfoString( pData, dwLen, SubBlock.c_str(), pManufacturerName );

    // Version, usual format is \\StringFileInfo\\%08X\\FileVersion but
    // sometimes of the form x.x.x.x ( text )
    SubBlock  = L"\\StringFileInfo\\";
    SubBlock += Format.UInt32Hex( codePage, false );
    SubBlock += L"\\FileVersion";
    GetVersionInfoString( pData, dwLen, SubBlock.c_str(), pVersionString );
    TidyUpVersionString( pVersionString );

    // If did not get the FileVersion, try the ProductVersion
    if ( pVersionString->IsEmpty() )
    {
        // Format is \\StringFileInfo\\%08X\\ProductVersion
        SubBlock  = L"\\StringFileInfo\\";
        SubBlock += Format.UInt32Hex( codePage, false );
        SubBlock += L"\\ProductVersion";
        GetVersionInfoString( pData, dwLen, SubBlock.c_str(), pVersionString );
    }

    // Description, format is "\\StringFileInfo\\%08X\\FileDescription"
    SubBlock  = L"\\StringFileInfo\\";
    SubBlock += Format.UInt32Hex( codePage, false );
    SubBlock += L"\\FileDescription";
    GetVersionInfoString( pData, dwLen, SubBlock.c_str(), pDescription );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the string that resides in the specified sub-block
//
//  Parameters:
//      pBlock     - pointer to the version info block
//      dwLen      - size of the block
//      lpSubBlock - the sub-block from which to fetch the value
//      pValue     - receives the string
//
//  Remarks:
//      The caller is responsible for ensuring the SubBlock is string
//      data
///
//  Returns:
//      void
//===============================================================================================//
void FileVersion::GetVersionInfoString( BYTE* pBlock,
                                        DWORD dwLen, LPCWSTR lpSubBlock, String* pValue )
{
    UINT   uLen    = 0;
    LPWSTR pBuffer = nullptr;

    if ( ( pBlock == nullptr ) || ( dwLen == 0 ) || ( lpSubBlock == nullptr ) )
    {
        return;     // No version information
    }

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = PXS_STRING_EMPTY;

    if ( VerQueryValue( pBlock,
                        lpSubBlock, reinterpret_cast<void**>( &pBuffer ), &uLen ) == 0 )
    {
        return;     // No information
    }

    // Bounds check to verify the data is entirely inside the block.
    // For string information uLen is in characters
    UINT  uTemp  = PXSMultiplyUInt32( uLen, sizeof ( wchar_t ) );
    BYTE* lpTemp = reinterpret_cast<BYTE*>( pBuffer );
    if ( ( lpTemp >= pBlock ) &&
         ( lpTemp + uTemp   ) <= ( pBlock + dwLen ) )
    {
        pValue->AppendChars( pBuffer, uLen );
    }
}

//===============================================================================================//
//  Description:
//      Get the DWORD that resides in the specified sub-block
//
//  Parameters:
//      pBlock     - pointer to the version info block
//      dwLen   - size of the block
//      lpSubBlock - the sub-block from which to fetch the value
//      pValue    - receives the DWORD Value
//
//  Remarks:
//      The caller is responsible for ensuring the data item is a DWORD.
//      In event the subblock contains an array of DWORDs the first one
//      is returned.
//
//  Returns:
//      void
//===============================================================================================//
void FileVersion::GetVersionInfoDWord( BYTE* pBlock,
                                       DWORD dwLen, LPCWSTR lpSubBlock, DWORD* pValue )
{
    UINT   uLen    = 0;
    DWORD* pBuffer = nullptr;

    if ( pBlock == nullptr || dwLen == 0 || lpSubBlock == nullptr )
    {
        return;     // No version information
    }

    if ( pValue == nullptr )
    {
        return;
    }
    *pValue = 0;

    if ( VerQueryValue( pBlock, lpSubBlock, reinterpret_cast<void**>( &pBuffer ), &uLen ) == 0 )
    {
        return;     // No information
    }

    // uLen should be at least the size of a DWORD.
    // Bounds check to verify the data is entirely inside the block.
    BYTE* lpTemp = reinterpret_cast<BYTE*>( pBuffer );
    if ( ( uLen   >= sizeof ( DWORD) ) &&
         ( lpTemp >= pBlock ) &&
         ( lpTemp + uLen    ) <= ( pBlock + dwLen ) )
    {
        *pValue = *pBuffer;
    }
}

//===============================================================================================//
//  Description:
//      Tidy up the specified version information string.
//
//  Parameters:
//      pVersionString - the string
//
//  Remarks:
//      Sometimes a module's version string has additional data other than
//      the file version so will truncate to not show these characters
//
//  Returns:
//      void
//===============================================================================================//
void FileVersion::TidyUpVersionString( String* pVersionString )
{
    size_t  i  = 0, charLength = 0;
    wchar_t ch = 0;

    if ( pVersionString == nullptr )
    {
        return;     // Nothing to do
    }

    // Tidy up, the format is ASCII
    pVersionString->Trim();
    charLength = pVersionString->GetLength();
    for ( i = 0; i < charLength; i++ )
    {
        // If not one of the below characters will ignore the remainder of the
        // version string
        ch = pVersionString->CharAt( i );
        if ( ( ch != 32 ) &&                    // space
             ( ch != 46 ) &&                    // point
             ( ch != 44 ) &&                    // comma
             ( ch != 43 ) &&                    // plus
             ( ch != 45 ) &&                    // minus
             ( ( ch <  48 ) || ( ch > 57 ) ) )  // digit
        {
            pVersionString->Truncate( i );
            pVersionString->Trim();
            break;
        }
    }
}
