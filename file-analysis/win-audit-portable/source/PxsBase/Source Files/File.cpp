///////////////////////////////////////////////////////////////////////////////////////////////////
//
// File Class Implementation
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

// 1. Own Interface
#include "PxsBase/Header Files/File.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/AllocateChars.h"
#include "PxsBase/Header Files/AllocateWChars.h"
#include "PxsBase/Header Files/AutoCloseHandle.h"
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/ByteArray.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Constructs a new File object
File::File()
     :m_bUnicodeLE( false ),
      m_hFile( INVALID_HANDLE_VALUE )
{
}

// Copy constructor - not allowed so no implementation

// Destructor
File::~File()
{
    try
    {
        Close();
    }
    catch ( const Exception& e)
    {
        PXSLogException( e, __FUNCTION__ );
    }
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
//      Close an open file
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void File::Close()
{
    if ( m_hFile == INVALID_HANDLE_VALUE )
    {
        return;     // Nothing to do
    }
    CloseHandle( m_hFile );
    m_hFile      = INVALID_HANDLE_VALUE;
    m_bUnicodeLE = false;
}

//===============================================================================================//
//  Description:
//      Create a new file, overwrite any existing one. The file remains open
//      for use
//
//  Parameters:
//      FileName  - new file's name
//      shareMode - file sharing options
//      Unicode   - flag to determine if want a Unicode file
//
//  Returns:
//      void
//===============================================================================================//
void File::CreateNew( const String& FileName, DWORD shareMode, bool unicode )
{
    BYTE bBOM[ 8 ];     // Holds the Byte Order Marker, two bytes for UTF-16LE

    if ( FileName.IsEmpty() )
    {
        throw ParameterException( L"FileName", __FUNCTION__ );
    }

    if ( IsOpen() )
    {
        Close();
    }

    m_hFile = CreateFile( FileName.c_str(),
                          GENERIC_READ | GENERIC_WRITE,
                          shareMode, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr );

    if ( m_hFile == INVALID_HANDLE_VALUE )
    {
        throw SystemException( GetLastError(), FileName.c_str(), "CreateFile");
    }
    else
    {
        // If want a Unicode file, write the byte order marker (BOM) at the
        // beginning of the file.
        if ( unicode )
        {
            memset( bBOM, 0, sizeof ( bBOM ) );
            bBOM[ 0 ] = 0xFF;
            bBOM[ 1 ] = 0xFE;
            Write( bBOM, 2 );
        }
        m_bUnicodeLE = unicode;
    }
}

//===============================================================================================//
//  Description:
//      Delete a file
//
//  Parameters:
//      FileName - the path to the file
//
//  Returns:
//      void
//===============================================================================================//
void File::Delete( const String& FileName )
{
    DWORD lastError;

    if ( FileName.IsEmpty() )
    {
        throw ParameterException( L"FileName", __FUNCTION__ );
    }

    if ( DeleteFile( FileName.c_str() ) == 0 )
    {
        // Ignore errors of the file not found
        lastError = GetLastError();
        if ( ( lastError != ERROR_FILE_NOT_FOUND ) &&
             ( lastError != ERROR_PATH_NOT_FOUND )  )
        {
            throw SystemException( lastError, FileName.c_str(), "DeleteFile" );
        }
    }
}

//===============================================================================================//
//  Description:
//      Determines if the path exists, GetFileAttributes can cope with
//      UNC paths
//
//  Parameters:
//      FileName - the path to the file
//
//  Returns:
//      true if the path exists, else false
//===============================================================================================//
bool File::Exists( const String& FileName ) const
{
    bool   exists  = false;
    DWORD  attributes = 0;

    if ( ( FileName.GetLength() == 0 ) ||
         ( FileName.GetLength() >  MAX_PATH ) )
    {
        return false;
    }

    // If GetFileAttributes returns INVALID_FILE_ATTRIBUTES will
    // assume it is because the file does not exist rather than an error
    attributes = GetFileAttributes( FileName.c_str() );
    if ( attributes != INVALID_FILE_ATTRIBUTES )
    {
        if ( !( attributes & FILE_ATTRIBUTE_DIRECTORY ) )
        {
            exists = true;
        }
    }

    return exists;
}

//===============================================================================================//
//  Description:
//      Flush the buffers of an open file to disk
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void File::Flush() const
{
    if ( m_hFile == INVALID_HANDLE_VALUE )
    {
        return;     // Nothing to do
    }

    if ( FlushFileBuffers( m_hFile ) == 0 )
    {
        throw SystemException( GetLastError(), L"FlushFileBuffers", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Get the file times of a file
//
//  Parameters:
//      FileName    - the path to the file
//      pCreation   - optional structure to receive the creation time
//      pLastAccess - optional structure  to receive the last access time
//      pLastWrite  - optional structure  to receive the last write time
//
//  Returns:
//      void
//===============================================================================================//
void File::GetTimes( const String& FileName,
                     FILETIME* pCreation,
                     FILETIME* pLastAccess, FILETIME* pLastWrite ) const
{
    HANDLE hFile = INVALID_HANDLE_VALUE;

    if ( FileName.IsEmpty() )
    {
        throw ParameterException( L"FileName", __FUNCTION__ );
    }

    hFile = CreateFile( FileName.c_str(),
                        0,
                        FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_EXISTING, 0, nullptr );
    if ( hFile == INVALID_HANDLE_VALUE )
    {
        throw SystemException( GetLastError(), FileName.c_str(), "CreateFile");
    }
    AutoCloseHandle CloseFileHandle( hFile );

    if ( pCreation )
    {
        memset( pCreation , 0, sizeof ( FILETIME ) );
    }

    if ( pLastAccess )
    {
        memset( pLastAccess, 0, sizeof ( FILETIME ) );
    }

    if ( pLastWrite )
    {
        memset( pLastWrite,  0, sizeof ( FILETIME ) );
    }

    if ( GetFileTime( hFile, pCreation, pLastAccess, pLastWrite ) == 0 )
    {
       throw SystemException( GetLastError(), FileName.c_str(), "GetFileTime");
    }
}

//===============================================================================================//
//  Description:
//      Get the file times of a file in string format
//
//  Parameters:
//      FileName    - the path to the file
//      wantTime    - if set to true, will append time
//      pCreation   - optional string object to receive the creation time
//      pLastAccess - optional string object to receive the last access time
//      pLastWrite  - optional string object to receive the last write time
//
//  Returns:
//      void
//===============================================================================================//
void File::GetTimes( const String& FileName,
                     bool wantTime,
                     String* pCreation, String* pLastAccess, String* pLastWrite ) const
{
    FILETIME  creation;
    FILETIME  lastAccess;
    FILETIME  lastWrite;
    Formatter Format;

    memset( &creation  , 0, sizeof ( creation   ) );
    memset( &lastAccess, 0, sizeof ( lastAccess ) );
    memset( &lastWrite , 0, sizeof ( lastWrite  ) );
    GetTimes( FileName, &creation, &lastAccess, &lastWrite );

    if ( pCreation )
    {
        *pCreation  = Format.FileTimeToLocalTimeUser( creation , wantTime );
    }

    if ( pLastAccess )
    {
        *pLastAccess= Format.FileTimeToLocalTimeUser( lastAccess, wantTime );
    }

    if ( pLastWrite )
    {
        *pLastWrite = Format.FileTimeToLocalTimeUser( lastWrite , wantTime );
    }
}

//===============================================================================================//
//  Description:
//      Get the offset of the file pointer from the start
//
//  Parameters:
//      None
//
//  Returns:
//      unsigned 64-bit value of current current pointer offset in the file
//===============================================================================================//
UINT64 File::GetFilePointer() const
{
    UINT64 filePointer = 0;
    LARGE_INTEGER liDistanceToMove;
    LARGE_INTEGER NewFilePointer;

    // The file must be open
    if ( m_hFile == INVALID_HANDLE_VALUE )
    {
        throw FunctionException( L"m_hFile", __FUNCTION__ );
    }

    memset( &liDistanceToMove, 0, sizeof ( liDistanceToMove ) );
    memset( &NewFilePointer  , 0, sizeof ( NewFilePointer   ) );
    if ( SetFilePointerEx( m_hFile, liDistanceToMove, &NewFilePointer, FILE_CURRENT ) == 0 )
    {
        throw SystemException( GetLastError(), L"SetFilePointerEx", __FUNCTION__ );
    }

    // Want an unsigned value
    if ( NewFilePointer.QuadPart > 0 )
    {
        filePointer = PXSCastInt64ToUInt64( NewFilePointer.QuadPart );
    }

    return filePointer;
}

//===============================================================================================//
//  Description:
//      Get the size of an open file
//
//  Parameters:
//      None
//
//  Returns:
//      unsigned 64-bit integer of file size in bytes
//===============================================================================================//
UINT64 File::GetSize() const
{
    UINT64 size = 0;
    LARGE_INTEGER FileSize;

    if ( m_hFile == INVALID_HANDLE_VALUE )
    {
        throw FunctionException( L"m_hFile", __FUNCTION__ );
    }

    memset( &FileSize, 0, sizeof ( FileSize ) );
    if ( GetFileSizeEx( m_hFile, &FileSize ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetFileSizeEx", __FUNCTION__ );
    }

    // Want an unsigned value
    if ( FileSize.QuadPart > 0 )
    {
        size = PXSCastInt64ToUInt64( FileSize.QuadPart );
    }

    return size;
}

//===============================================================================================//
//  Description:
//      Tell whether or not the file is open
//
//  Parameters:
//      None
//
//  Returns:
//      true if file is open, else false
//===============================================================================================//
bool File::IsOpen() const
{
    if ( m_hFile != INVALID_HANDLE_VALUE )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Tells whether an open file is in Unicode
//
//  Parameters:
//      None
//
//  Remarks:
//      The class scope Unicode flag is set when the file is created or opened.
//
//  Remarks:
//      If no file is open, will not raise an error
//
//  Returns:
//      true if file is Unicode, else false
//===============================================================================================//
bool File::IsUnicode() const
{
    return m_bUnicodeLE;
}

//===============================================================================================//
//  Description:
//      Determine if a string is a valid file name.
//
//  Parameters:
//      FileName - the file name to check
//
//  Remarks:
//      Only the file name is checked not full path so back/forward slash
//      characters are considered invalid
//
//      See MSDN "Naming Files, Paths and Namespaces"
//
//  Returns:
//      true if name is valid as a file name, else false
//===============================================================================================//
bool File::IsValidFileName( const String& FileName ) const
{
    size_t charLengh = 0, i = 0;
    String DeviceNameDot;

    charLengh = FileName.GetLength();
    if ( ( charLengh == 0 ) || ( charLengh > MAX_PATH ) )
    {
        return false;
    }

    // Below ascii 32 not allowed
    for ( i = 0; i < charLengh; i++ )
    {
        if ( FileName.CharAt( i ) < 32 )
        {
            return false;
        }
    }

    // Scan for illegal characters
    wchar_t Illegals[] = { '/', '\\', ':', '*', '?', '"', '<', '>', '|' };

    for ( i = 0; i < ARRAYSIZE( Illegals ); i++ )
    {
        if ( FileName.IndexOf( Illegals[ i ], 0 ) != PXS_MINUS_ONE )
        {
            return false;
        }
    }

    // Name cannot only be dots
    if ( FileName.IsOnlyChar( '.' ) )
    {
       return false;
    }

    // Device names cannot be used for a name nor can it be the first
    // part of a name.
    LPCWSTR DeviceNames[] = { L"CLOCK$",
                              L"AUX" ,
                              L"CON" ,
                              L"NUL" ,
                              L"PRN" ,
                              L"COM1",
                              L"COM2",
                              L"COM3",
                              L"COM4",
                              L"COM5",
                              L"COM6",
                              L"COM7",
                              L"COM8",
                              L"COM9",
                              L"LPT1",
                              L"LPT2",
                              L"LPT3",
                              L"LPT4",
                              L"LPT5",
                              L"LPT6",
                              L"LPT7",
                              L"LPT8",
                              L"LPT9"};

    for ( i = 0; i < ARRAYSIZE( DeviceNames ); i++ )
    {
        // Make a name with a dot, e.g. CLOCK$.
        DeviceNameDot  = DeviceNames[ i ];
        DeviceNameDot += PXS_CHAR_DOT;

        if ( ( FileName.CompareI( DeviceNames[ i ] ) == 0 ) ||
             ( FileName.IndexOfI( DeviceNameDot.c_str() ) == 0 )  )
        {
            return false;
        }
    }

    // MSDN "Security Considerations: International Features
    // "Windows code page and OEM character sets used on Japanese-language
    // systems contain the Yen symbol (00A5) instead of a backslash (\). Thus,
    // the Yen character is a prohibited character for those file systems."
    if ( PRIMARYLANGID( GetUserDefaultLangID() ) == LANG_JAPANESE )
    {
        if ( FileName.IndexOf( 0x00A5, 0 ) != PXS_MINUS_ONE )
        {
            return false;
        }
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Open an existing file for read/write access
//
//  Parameters:
//      FileName      - path to file
//      desiredAccess - flags to indicate access mode
//      shareMode     - flags to indicate share mode
//      numTries      - Number of attempts to try to open the file in case
//                      in case locked by another process.
//      isTextFile    - flag to indicate opening a text file, if so will
//                      detect the Byte Order Marker
//
//    Remarks
//      Upon a successful open, if the file is Unicode Little Endian, the file
//      pointer is on the first byte after the BOM
//
//  Returns:
//      void
//===============================================================================================//
void File::Open( const String& FileName,
                 DWORD desiredAccess, DWORD shareMode, DWORD numTries, bool isTextFile )
{
    bool  shared  = false;
    DWORD counter = 0, lastError = 0;

    if ( FileName.IsEmpty() )
    {
        throw ParameterException( L"FileName", __FUNCTION__ );
    }
    Close();

    // Try to open the file numTries waiting 1s in between tries
    do
    {
        counter++;
        shared    = false;
        lastError = ERROR_SUCCESS;
        m_hFile   = CreateFile( FileName.c_str(),
                                desiredAccess,
                                shareMode,
                                nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr );

        if ( m_hFile == INVALID_HANDLE_VALUE )
        {
            // Wait on sharing error
            lastError = GetLastError();
            if ( ( lastError == ERROR_LOCK_VIOLATION    ) ||
                 ( lastError == ERROR_SHARING_VIOLATION )  )
            {
                shared = true;
                Sleep( 1000 );
            }
        }
    }
    while ( ( true    == shared   ) &&
            ( counter <  numTries ) &&
            ( m_hFile == INVALID_HANDLE_VALUE ) );

    if ( lastError != ERROR_SUCCESS )
    {
        throw SystemException( lastError, FileName.c_str(), "CreateFile" );
    }

    if ( isTextFile )
    {
        OnOpenDetectUtf16Bom();
    }
}

//===============================================================================================//
//  Description:
//      Open an existing binary file for read/write access
//
//  Parameters:
//      FileName - path to file
//
//  Returns:
//      void
//===============================================================================================//
void File::OpenBinaryExclusive( const String& FileName )
{
    Open( FileName, GENERIC_READ | GENERIC_WRITE, 0, 1, false );
}

//===============================================================================================//
//  Description:
//      Open an existing binary file for read/write access in exclusive mode
//
//  Parameters:
//      FileName - path to file
//      numTries - Number of attempts to try to open the file in case
//                 in case locked by another process.//
//  Returns:
//      void
//===============================================================================================//
void File::OpenExclusive( const String& FileName, DWORD numTries )
{
    Open( FileName, GENERIC_READ | GENERIC_WRITE, 0, numTries, false );
}

//===============================================================================================//
//  Description:
//      Open an existing text file for read/write access
//
//  Parameters:
//      FileName - path to file
//
//  Returns:
//      void
//===============================================================================================//
void File::OpenText( const String& FileName )
{
    Open( FileName, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, 1, true );
}

//===============================================================================================//
//  Description:
//      Read bytes from an open file
//
//  Parameters:
//        pBuffer     - pointer to buffer to hold data
//        bufferBytes - size of buffer in bytes, zero is valid
//
//  Remarks:
//      If past the EOF, will return zero
//
//  Returns:
//      the number of bytes read, else zero if at end of file
//===============================================================================================//
size_t File::Read( BYTE* pBuffer, size_t bufferBytes ) const
{
    DWORD nNumberOfBytesToRead = 0, NumberOfBytesRead = 0;

    if ( pBuffer == nullptr )
    {
        throw ParameterException( L"pBuffer", __FUNCTION__ );
    }

    if ( m_hFile == INVALID_HANDLE_VALUE )
    {
        throw FunctionException( L"m_hFile", __FUNCTION__ );
    }

    memset( pBuffer, 0, bufferBytes );
    nNumberOfBytesToRead = PXSCastSizeTToUInt32( bufferBytes );
    if ( ReadFile( m_hFile, pBuffer, nNumberOfBytesToRead, &NumberOfBytesRead, nullptr ) == 0 )
    {
        throw SystemException( GetLastError(), L"ReadFile", __FUNCTION__ );
    }

    return NumberOfBytesRead;
}

//===============================================================================================//
//  Description:
//      Read the entire contents of the open file into a byte array. Will limit
//      to 16MB
//
//  Parameters:
//      pBytes - object to hold data
//
//  Returns:
//      void
//===============================================================================================//
void File::ReadAll( ByteArray* pBytes ) const
{
    const  UINT64 MAX_BUFFER_SIZE = 256 * 1024 *1024;     // 256MB
    BYTE*  pBuffer  = nullptr;
    UINT64 size     = 0;
    size_t fileSize = 0;
    AllocateBytes AllocBytes;

    if ( pBytes == nullptr )
    {
        throw ParameterException( L"pBytes", __FUNCTION__ );
    }
    pBytes->Free();

    if ( m_hFile == INVALID_HANDLE_VALUE )
    {
        throw FunctionException( L"m_hFile", __FUNCTION__ );
    }

    size = GetSize();
    if ( size > MAX_BUFFER_SIZE )
    {
        throw BoundsException( L"size > MAX_BUFFER_SIZE", __FUNCTION__ );
    }
    fileSize = PXSCastUInt64ToSizeT( size );
    pBuffer  = AllocBytes.New( fileSize );
    Seek( 0 );
    Read( pBuffer, fileSize );
    pBytes->Append( pBuffer, fileSize );
}

//===============================================================================================//
//  Description:
//      Read the entire contents of the open file into a string. Will limit
//      to 16MB
//
//  Parameters:
//      pContents - string object to hold data
//
//  Returns:
//      void
//===============================================================================================//
void File::ReadAll( String* pContents ) const
{
    const  UINT64 MAX_BUFFER_SIZE = 16 * 1024 *1024;     // 16MB
    BYTE*  pBuffer  = nullptr;
    UINT64 size     = 0;
    size_t fileSize = 0, bufferSize = 0;
    AllocateBytes AllocBytes;

    if ( pContents == nullptr )
    {
        throw ParameterException( L"pContents", __FUNCTION__ );
    }
    *pContents = PXS_STRING_EMPTY;

    if ( m_hFile == INVALID_HANDLE_VALUE )
    {
        throw FunctionException( L"m_hFile", __FUNCTION__ );
    }

    size = GetSize();
    if ( size == 0 )
    {
        return;     // Nothing to do;
    }

    if ( size > MAX_BUFFER_SIZE )
    {
        throw BoundsException( L"size > MAX_BUFFER_SIZE", __FUNCTION__ );
    }
    fileSize   = PXSCastUInt64ToSizeT( size );
    bufferSize = fileSize;
    bufferSize = PXSAddSizeT( bufferSize, sizeof ( wchar_t ) );    // Terminator
    pBuffer = AllocBytes.New( bufferSize );
    if ( m_bUnicodeLE )
    {
        // Skip over the BOM, two bytes
        if ( fileSize > 2 )
        {
            Seek( 2 );
            Read( pBuffer, fileSize - 2 );
        }
    }
    else
    {
        Seek( 0 );
        Read( pBuffer, fileSize );
    }
    memset( pBuffer + fileSize, 0, sizeof ( wchar_t ) );  // Terminate

    if ( m_bUnicodeLE )
    {
        // Unicode file to Unicode string
        *pContents = reinterpret_cast<LPCWSTR>( pBuffer );
    }
    else
    {   // ANSI file to Unicode string
        pContents->SetAnsi( reinterpret_cast<char*>( pBuffer ) );
    }
}

//===============================================================================================//
//  Description:
//      Read a line of text from an open file. The caller should have set
//      the position after the BOM, if there is one
//
//  Parameters:
//      pLine - string object to receive the line
//
//    Remarks
//      Using a  byte reader approach so only use on small text files
//
//  Returns:
//      Total number of bytes read.
//===============================================================================================//
size_t File::ReadLine( String* pLine ) const
{
    int     nTemp = 0;
    size_t  total = 0;
    BYTE    Buffer[ 8 ];            // Want to read two bytes
    wchar_t szWideChar[ 8 ];        // Holds a wide char + wide terminator
    DWORD   NumberOfBytesRead = 0, charWidth = 0;
    String  TempString;

    if ( pLine == nullptr )
    {
        throw ParameterException( L"pLine", __FUNCTION__ );
    }
    *pLine = PXS_STRING_EMPTY;
    pLine->Allocate( 1024 );

    if ( m_hFile == INVALID_HANDLE_VALUE )
    {
        throw FunctionException( L"m_hFile", __FUNCTION__ );
    }

    // Set the character width
    if ( m_bUnicodeLE )
    {
        charWidth = sizeof ( wchar_t );
    }
    else
    {
        charWidth = sizeof (char );
    }

    // Read 1 character at a time until error, EOF or found CR or LF
    memset( Buffer, 0, sizeof ( Buffer ) );
    while ( ReadFile( m_hFile, Buffer, charWidth, &NumberOfBytesRead, nullptr ) )
    {
        total = PXSAddSizeT( total, NumberOfBytesRead );

        // Break out if EOF or error
        if ( charWidth != NumberOfBytesRead )
        {
            break;
        }

        // See if its an end-of-line character, which is one of 0x0A,
        // 0x0D, 0x0A0D
        if ( ( Buffer[ 0 ] == 0x0A ) || ( Buffer[ 0 ] == 0x0D ) )
        {
            break;
        }

        // Convert buffer to character code in use
        if ( m_bUnicodeLE )
        {
            // Would have read two bytes
            nTemp = ( Buffer[ 1 ] << 8 ) + Buffer[ 0 ];     // Reverse
            szWideChar[ 0 ] = PXSCastInt32ToUInt16 ( nTemp );
            szWideChar[ 1 ] = PXS_CHAR_NULL;
            TempString = szWideChar;
        }
        else
        {
            Buffer[ 1 ] = PXS_CHAR_NULL;
            TempString.SetAnsi( reinterpret_cast<char*>( Buffer ) );
        }
        *pLine += TempString;

        // Reset for next pass
        NumberOfBytesRead = 0;
        memset( Buffer, 0, sizeof ( Buffer ) );
    }

    return total;
}

//===============================================================================================//
//  Description:
//      Read the contents of a text file and put its line into an array
//
//  Parameters:
//      FileName - path to file
//      maxLines - maximum number of lines to read
//      pLines   - receives the lines
//
//  Returns:
//      void
//===============================================================================================//
void File::ReadLineArray( const String& FileName, DWORD maxLines, StringArray* pLines )
{
    const  DWORD NUM_TRIES = 5;
    DWORD  counter   = 0;
    size_t bytesRead = 0;
    String Line;

    if ( pLines == nullptr )
    {
        throw ParameterException( L"pLines", __FUNCTION__ );
    }
    pLines->RemoveAll();
    Open( FileName, GENERIC_READ, FILE_SHARE_READ, NUM_TRIES, true );
    try
    {
        bytesRead = ReadLine( &Line );
        while ( bytesRead && ( counter < maxLines ) )
        {
            counter = PXSAddUInt32( counter, 1 );
            if ( Line.GetLength() )
            {
                pLines->Add( Line );
            }
            Line.Zero();
            bytesRead = ReadLine( &Line );
        }
    }
    catch ( const Exception )
    {
        Close();
        throw;
    }
    Close();
}

//===============================================================================================//
//  Description:
//      Move to a position from the start of the file, i.e. the move
//      method is FILE_BEGIN.
//
//  Parameters:
//      position - Number of bytes to move the pointer.
//
//  Returns:
//      void
//===============================================================================================//
void File::Seek( UINT64 position ) const
{
    LARGE_INTEGER liDistanceToMove;

    if ( m_hFile == INVALID_HANDLE_VALUE )
    {
        throw FunctionException( L"m_hFile", __FUNCTION__ );
    }

    memset( &liDistanceToMove, 0, sizeof ( liDistanceToMove ) );
    liDistanceToMove.QuadPart = PXSCastUInt64ToInt64( position );
    if ( SetFilePointerEx( m_hFile, liDistanceToMove, nullptr, FILE_BEGIN ) == 0 )
    {
        throw SystemException( GetLastError(), L"SetFilePointerEx", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Go to the end of the file
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void File::SeekToEnd() const
{
    LARGE_INTEGER liDistanceToMove;

    if ( m_hFile == INVALID_HANDLE_VALUE )
    {
        throw FunctionException( L"m_hFile", __FUNCTION__ );
    }

    // Set distance to move to zero
    memset( &liDistanceToMove, 0, sizeof ( liDistanceToMove ) );
    if ( SetFilePointerEx( m_hFile, liDistanceToMove, nullptr, FILE_END ) == 0 )
    {
        throw SystemException( GetLastError(), L"SetFilePointerEx", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Set the size of the file
//
//  Parameters:
//      size - the new size in bytes
//
//  Returns:
//      void
//===============================================================================================//
void File::SetSize( UINT64 size )
{
    if ( m_hFile == INVALID_HANDLE_VALUE )
    {
        throw FunctionException( L"m_hFile", __FUNCTION__ );
    }

    Seek( size );
    if ( SetEndOfFile( m_hFile ) == 0 )
    {
        throw SystemException( GetLastError(), L"SetEndOfFile", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Write bytes to an open file
//
//  Parameters:
//      pBuffer  - pointer to buffer of data to write
//      bufferBytes - the size of the buffer in bytes
//
//  Remarks:
//      A zero length buffer is allowed. A value of zero specifies a null
//      write operation which cause the time stamp to change.
//
//  Returns:
//      void
//===============================================================================================//
void File::Write( const void* pBuffer, size_t bufferBytes ) const
{
    DWORD nNumberOfBytesToWrite = 0, NumberOfBytesWritten = 0;

    if ( m_hFile == INVALID_HANDLE_VALUE )
    {
        throw FunctionException( L"m_hFile", __FUNCTION__ );
    }
    nNumberOfBytesToWrite = PXSCastSizeTToUInt32( bufferBytes );
    if ( WriteFile( m_hFile, pBuffer, nNumberOfBytesToWrite, &NumberOfBytesWritten, nullptr ) )
    {
        // Check that all the bytes were written, this is fine as it is
        // not an overlapped operation
        if ( nNumberOfBytesToWrite != NumberOfBytesWritten )
        {
            throw SystemException( ERROR_WRITE_FAULT, L"NumberOfBytesWritten", "WriteFile" );
        }
    }
    else
    {
        throw SystemException( GetLastError(), L"WriteFile", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Write a string to an open file
//
//  Parameters:
//      Text - the text string
//
//  Returns:
//      void
//===============================================================================================//
void File::WriteChars( const String& Text ) const
{
    size_t    bufferBytes = 0;
    char*     pszAnsi     = nullptr;
    Formatter Format;
    AllocateChars AllocChars;

    if ( Text.IsEmpty() )
    {
        return;     // Nothing to do
    }

    if ( m_bUnicodeLE )
    {
        // Input is Unicode, want Unicode output
        bufferBytes = Text.GetLength();
        bufferBytes = PXSMultiplySizeT( bufferBytes, sizeof ( wchar_t ) );
        Write( Text.c_str(), bufferBytes );
    }
    else
    {
        // Input is Unicode, want ANSI output
        bufferBytes = Text.GetAnsiMultiByteLength();
        pszAnsi     = AllocChars.New(  bufferBytes );
        Format.StringToAnsi( Text, pszAnsi, bufferBytes );
        Write( pszAnsi, bufferBytes - 1 );            // Not the NULL
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////


//===============================================================================================//
//  Description:
//      After opening a file, move past any UTF-16 BOM. If the BOM is
//      detected the class scope Unicode flag is set.
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void File::OnOpenDetectUtf16Bom()
{
    BYTE  bBOM[ 2 ] = { 0 };    // The Byte Order Marker, two bytes for UTF-16LE

    if ( m_hFile == INVALID_HANDLE_VALUE )
    {
        throw FunctionException( L"m_hFile", __FUNCTION__ );
    }

    // File pointer should be at zero
    if ( GetFilePointer() )
    {
        throw FunctionException( L"Position", __FUNCTION__ );
    }

    // Read the first 2 bytes and see if it is the Byte Order Marker
    // for UTF-16 little endian
    m_bUnicodeLE = false;
    if ( GetSize() < sizeof ( bBOM ) )
    {
        return;
    }

    if ( Read( bBOM, sizeof ( bBOM ) ) == sizeof ( bBOM ) )
    {
        if ( ( bBOM[ 0 ] == 0xFF ) && ( bBOM[ 1 ] == 0xFE ) )
        {
            m_bUnicodeLE = true;    // BOM matches, set flag
        }
        else
        {
            // Rewind to start of file
            m_bUnicodeLE = false;
            SetFilePointer( m_hFile, 0, nullptr, FILE_BEGIN );
        }
    }
}
