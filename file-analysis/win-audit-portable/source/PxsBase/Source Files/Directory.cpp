///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Directory Class Implementation
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
#include "PxsBase/Header Files/Directory.h"

// 2. C System Files
#include <ShlObj.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/ComException.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor an empty directory object
Directory::Directory()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
Directory::~Directory()
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
//      Create a new directory
//
//  Parameters:
//      Path - string pointer to path of directory
//
//  Remarks:
//      If the directory already exists no error is raised.
//
//  Returns:
//      void
//===============================================================================================//
void Directory::CreateNew( const String& Path ) const
{
    DWORD lastError;

    if ( Path.IsEmpty() )
    {
        throw ParameterException( L"Path", __FUNCTION__ );
    }

    if ( CreateDirectory( Path.c_str(), nullptr ) == 0 )
    {
        lastError = GetLastError();
        if ( ( lastError != ERROR_FILE_EXISTS    ) &&
             ( lastError != ERROR_ALREADY_EXISTS )  )
        {
            throw SystemException( lastError, Path.c_str(), "CreateDirectory" );
        }
    }
}

//===============================================================================================//
//  Description:
//      Deletes an empty directory
//
//  Parameters:
//      Path - pointer to null terminated path
//
//  Remarks:
//      If the directory does not exist, no error is raised.
//
//  Returns:
//      void
//===============================================================================================//
void Directory::Delete( const String& Path ) const
{
    DWORD lastError;

    if ( Path.IsEmpty() )
    {
        throw ParameterException( L"pszPath", __FUNCTION__ );
    }

    if ( RemoveDirectory( Path.c_str() ) == 0 )
    {
        // Check to make sure someone else has not deleted it
        lastError = GetLastError();
        if ( ( lastError != ERROR_FILE_NOT_FOUND ) &&
             ( lastError != ERROR_PATH_NOT_FOUND )  )
        {
            throw SystemException( lastError, Path.c_str(), "RemoveDirectory" );
        }
    }
}

//===============================================================================================//
//  Description:
//        Determines if the directory exists
//
//  Parameters:
//        Path - directory path path
//
//  Remarks:
//  GetFileAttributes() can cope with UNC paths
//
//      Common errors from GetFileAttributes() are
//      ERROR_FILE_NOT_FOUND     = #2
//      ERROR_PATH_NOT_FOUND     = #3
//      ERROR_BAD_NETPATH        = #53
//      ERROR_BAD_PATHNAME       = #161
//      ERROR_INVALID_NAME       = #123
//      ERROR_DIRECTORY          = #267
//      ERROR_NO_NET_OR_BAD_PATH = #1203
//
//  Returns:
//        true if directory exists, else false
//===============================================================================================//
bool Directory::Exists( const String& Path ) const
{
    bool   exists  = false;
    DWORD  attributes = 0;
    String ErrorMessage;

    if ( ( Path.IsEmpty() ) || ( Path.GetLength() > MAX_PATH ) )
    {
        return false;
    }

    // Get file attributes. If get back INVALID_FILE_ATTRIBUTES assume it
    // is because the file does not exist rather than an error in the
    // GetFileAttributes function
    attributes = GetFileAttributes( Path.c_str() );
    if ( attributes != INVALID_FILE_ATTRIBUTES )
    {
        // Make sure it a directory
        if ( attributes & FILE_ATTRIBUTE_DIRECTORY )
        {
            exists = true;
        }
    }

    return exists;
}

//===============================================================================================//
//  Description:
//      Get the specified special folder using
//
//  Parameters:
//      nFolder        - folder ID e.g. CSIDL_ALTSTARTUP
//      pDirectoryPath - string object to receive the directory
//
//  Remarks:
//      Appends a trailing backslash '\' if function succeeds
//
//      Does not check for the existence of the directory/folder
//
//      For Windows 2000+ will use SHGetFolderPath as indicated by
//      The Compatibility Guide
//
//  Returns:
//      void
//===============================================================================================//
void Directory::GetSpecialDirectory( int nFolder, String* pDirectoryPath )
{
    wchar_t szPath[ MAX_PATH + 1 ] = { 0 };  // Must be at least MAX_PATH
    HRESULT hResult;

    if ( pDirectoryPath == nullptr )
    {
        throw ParameterException( L"pDirectoryPath", __FUNCTION__ );
    }
    *pDirectoryPath = PXS_STRING_EMPTY;

    // SHGetFolderPath works on Windows 2000+. Get the folder's current path
    // as opposed to its default path. Use an access token of NULL.
    hResult = SHGetFolderPath( nullptr, nFolder, nullptr, SHGFP_TYPE_CURRENT, szPath );
    if ( SUCCEEDED( hResult ) )
    {
        szPath[ ARRAYSIZE( szPath ) - 1 ] = PXS_CHAR_NULL;
        *pDirectoryPath = szPath;
        pDirectoryPath->Trim();
        if ( ( pDirectoryPath->GetLength() ) &&
             ( pDirectoryPath->EndsWithCharacter(PXS_PATH_SEPARATOR) == false))
        {
            *pDirectoryPath += PXS_PATH_SEPARATOR;
        }
    }
    else
    {
        throw ComException( hResult, L"SHGetFolderPath", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Get the system's temporary directory
//
//  Parameters:
//      pTempDirectory - string object to receive the temp directory
//
//  Remarks:
//      GetTempPath on NT only checks the TMP/TEMP variable, it does not
//      check for the existence the directory.
//
//      Always ensure there is a trailing backslash '\' if function succeeds.
//
//  Returns:
//      void
//===============================================================================================//
void Directory::GetTempDirectory( String* pTempDirectory ) const
{
    String  Path;
    wchar_t Buffer[ MAX_PATH + 1 ] = { 0 };

    if ( pTempDirectory == nullptr )
    {
        throw ParameterException( L"pTempDirectory", __FUNCTION__ );
    }
    *pTempDirectory = PXS_STRING_EMPTY;

    if ( GetTempPath( ARRAYSIZE( Buffer ), Buffer ) == 0 )
    {
        throw SystemException( ERROR_INSUFFICIENT_BUFFER, L"GetTempPath", __FUNCTION__ );
    }
    Buffer[ ARRAYSIZE( Buffer ) - 1 ] = PXS_CHAR_NULL;
    Path = Buffer;

    // Make sure it exits
    if ( Exists( Path ) == false )
    {
        throw SystemException( ERROR_PATH_NOT_FOUND, Path.c_str(), __FUNCTION__ );
    }

    // Ensure have a trailing separator
    if ( ( Path.GetLength() ) &&
         ( Path.EndsWithCharacter( PXS_PATH_SEPARATOR ) == false ) )
    {
         Path += PXS_PATH_SEPARATOR;
    }
    *pTempDirectory = Path;
}

//===============================================================================================//
//  Description:
//      Determine if the specified directory is on a local drive
//
//  Parameters:
//      Path - directory path
//
//  Remarks:
//      Does not check for existence of directory, just if the drive is local.
//
//  Returns:
//      true if the directory is on a local drive otherwise false
//===============================================================================================//
bool Directory::IsOnLocalDrive( const String& Path ) const
{
    bool    local = false;
    UINT    driveType = 0;
    String  DirectoryPath;
    wchar_t RootPathName[ 8 ] = { 0 };  // letter + colon + backslash

    // Check argument
    DirectoryPath = Path;
    DirectoryPath.Trim();
    if ( DirectoryPath.IsEmpty() )
    {
        return false;   // Say its not on a local drive
    }

    // Get the drive letter
    RootPathName[ 0 ] = DirectoryPath.CharAt( 0 );
    RootPathName[ 1 ] = PXS_CHAR_COLON;
    RootPathName[ 2 ] = PXS_PATH_SEPARATOR;
    RootPathName[ 3 ] = PXS_CHAR_NULL;
    driveType = GetDriveType( RootPathName );
    if ( ( driveType == DRIVE_REMOVABLE ) || ( driveType == DRIVE_FIXED ) )
    {
        local = true;  // Its a local drive
    }

    return local;
}

//===============================================================================================//
//  Description:
//      Get a list of file names in a directory which match the specified
//      extension. Does not descend the directory tree.
//
//  Parameters:
//      Path       - directory path
//      DotExt     - file extension to filter for, includes the dot. If
//                   NULL or "" is specified will search for *.*
//      pFileNames - string array to receive file names
//
//  Returns:
//      void
//===============================================================================================//
void Directory::ListFiles( const String& Path,
                           const String& DotExt, StringArray* pFileNames )
{
    DWORD     lastError = 0;
    size_t    arraySize = 0;
    HANDLE    hSearch   = nullptr;
    String    DirectoryPath, Found;
    Formatter Format;
    WIN32_FIND_DATA FindFileData;

    if ( pFileNames == nullptr )
    {
        throw ParameterException( L"pFileNames", __FUNCTION__ );
    }
    pFileNames->RemoveAll();

    if ( Path.IsEmpty() )
    {
        return;     // Nothing to do
    }

    // Add extension, if none supplied
    DirectoryPath = Path;
    if ( DirectoryPath.EndsWithCharacter( PXS_PATH_SEPARATOR ) == false )
    {
        DirectoryPath += PXS_PATH_SEPARATOR;
    }
    if ( DotExt.GetLength() )
    {
        DirectoryPath += L"*";
        DirectoryPath += DotExt;
    }
    else
    {
        DirectoryPath += L"*.*";    // Find all
    }

    memset( &FindFileData, 0, sizeof ( FindFileData  ) );
    hSearch = FindFirstFile( DirectoryPath.c_str(), &FindFileData  );
    if ( hSearch == INVALID_HANDLE_VALUE )
    {
        lastError = GetLastError();
        if ( lastError == ERROR_FILE_NOT_FOUND )
        {
            return;     // Nothing found
        }
        throw SystemException( lastError,
                               DirectoryPath.c_str(), "FindFirstFile" );
    }

    // Catch any exceptions to close the enumeration
    try
    {
        do
        {
            // Add files to the array
            if ( !(FILE_ATTRIBUTE_DIRECTORY & FindFileData.dwFileAttributes ) )
            {
                // N.B. FindFileData.cFileName may not be terminated
                arraySize = ARRAYSIZE( FindFileData.cFileName );
                FindFileData.cFileName[ arraySize - 1 ] = PXS_CHAR_NULL;
                Found = FindFileData.cFileName;
                Found.Trim();
                if ( Found.GetLength() )
                {
                    pFileNames->Add( Found );
                }
            }
        }
        while ( FindNextFile( hSearch, &FindFileData ) );
    }
    catch ( const Exception& )
    {
        FindClose( hSearch );
        throw;
    }
    FindClose( hSearch );
    pFileNames->Sort( true );
}

//===============================================================================================//
//  Description:
//        Split a file path
//
//  Parameters:
//      pszPath - pointer to directory
//      pDrive  - string object to receive the drive
//      pDir    - string object to receive the directory
//      pFname  - string object to receive the file name
//      pExt    - string object to receive the extension, including leading the
//                period (.)
//
//  Returns:
//      void
//===============================================================================================//
void Directory::SplitPath( const String& Path,
                           String* pDrive,
                           String* pDir, String* pFname, String* pExt ) const
{
    wchar_t szDrive[ _MAX_DRIVE + 1 ] = { 0 };
    wchar_t szDir  [ _MAX_DIR   + 1 ] = { 0 };
    wchar_t szFname[ _MAX_FNAME + 1 ] = { 0 };
    wchar_t szExt  [ _MAX_EXT   + 1 ] = { 0 };

    if ( ( pDrive == nullptr ) ||
         ( pDir   == nullptr ) ||
         ( pFname == nullptr ) ||
         ( pExt   == nullptr )  )
    {
        throw ParameterException( L"pDrive,pDir,pFname,pExt", __FUNCTION__ );
    }
    pDrive->Zero();
    pDir->Zero();
    pFname->Zero();
    pExt->Zero();
    if ( Path.IsEmpty() )
    {
        return;     // Nothing to do
    }

    if ( _wsplitpath_s( Path.c_str(),
                        szDrive, ARRAYSIZE( szDrive ),
                        szDir  , ARRAYSIZE( szDir   ),
                        szFname, ARRAYSIZE( szFname ),
                        szExt  , ARRAYSIZE( szExt   ) ) )
    {
        throw ParameterException( L"pszPath", __FUNCTION__ );
    }
    szDrive[ ARRAYSIZE( szDrive ) - 1 ] = PXS_CHAR_NULL;
    szDir  [ ARRAYSIZE( szDir )   - 1 ] = PXS_CHAR_NULL;
    szFname[ ARRAYSIZE( szFname ) - 1 ] = PXS_CHAR_NULL;
    szExt  [ ARRAYSIZE( szExt )   - 1 ] = PXS_CHAR_NULL;

    *pDrive = szDrive;
    pDrive->Truncate( 1 );  // Only want the drive letter
    *pDir   = szDrive;
    *pDir  += szDir;
    *pFname = szFname;
    *pExt   = szExt;

    // Want full directory path with trailing separator
    if ( ( pDir->GetLength() ) &&
         ( pDir->EndsWithCharacter( PXS_PATH_SEPARATOR ) == false ) )
    {
        *pDir += PXS_PATH_SEPARATOR;
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
