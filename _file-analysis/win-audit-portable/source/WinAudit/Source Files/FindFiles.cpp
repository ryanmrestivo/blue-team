///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Find Files Implementation
//
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Copyright 1987-2016 PARMAVEX SERVICES
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
#include "WinAudit/Header Files/FindFiles.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Directory.h"
#include "PxsBase/Header Files/FileVersion.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/SystemInformation.h"
#include "PxsBase/Header Files/TArray.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
FindFiles::FindFiles()
          :MAX_RECURSION_DEPTH( 32 ),
           m_Files()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
FindFiles::~FindFiles()
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
//      Search a given directory for files of the specified extensions
//
//  Parameters:
//      Directory  - directory path
//      Extensions - string of extensions to search for array, must include
//                    the leading dot
//      DirectoryStopList - a list of directories not to search
//
//  Remarks:
//      Fills a class scope list with files
//
//  Returns:
//      void
//===============================================================================================//
void FindFiles::FindInDirectory( const String& DirectoryPath,
                                 const StringArray& Extensions,
                                 const StringArray& DirectoryStopList )
{
    UINT        errorMode;
    DWORD       fileAttributes;
    String      SearchPath;
    Directory   DirObject;
    Formatter   Format;
    StringArray EmptyStopList;

    m_Files.RemoveAll();
    if ( DirectoryPath.IsEmpty() )
    {
        throw ParameterException( L"DirectoryPath", __FUNCTION__ );
    }

    if ( DirObject.Exists( DirectoryPath ) == false )
    {
        throw SystemException( ERROR_PATH_NOT_FOUND, DirectoryPath.c_str(), __FUNCTION__ );
    }

    // Must be a directory
    fileAttributes = GetFileAttributes( DirectoryPath.c_str() );
    if ( !( fileAttributes & FILE_ATTRIBUTE_DIRECTORY ) )
    {
        throw SystemException( ERROR_INVALID_NAME, DirectoryPath.c_str(), __FUNCTION__ );
    }

    // If the directory path is only one character long add a colon
    SearchPath = DirectoryPath;
    if ( SearchPath.GetLength() == 1 )
    {
        SearchPath += PXS_CHAR_COLON;
    }

    // Ensure have a path separator
    if ( SearchPath.EndsWithCharacter( PXS_PATH_SEPARATOR ) == false )
    {
        SearchPath += PXS_PATH_SEPARATOR;
    }

    if ( Extensions.GetSize() == 0 )
    {
        throw ParameterException( L"Extensions", __FUNCTION__ );
    }

    // Catch exceptions so can reset error mode
    errorMode = SetErrorMode( SEM_FAILCRITICALERRORS );
    try
    {
        Find( SearchPath,
              fileAttributes,
              0,              // Start with recursion depth of zero
              Extensions, DirectoryStopList );
    }
    catch ( const Exception& )
    {
        SetErrorMode( errorMode );
        throw;
    }
    SetErrorMode( errorMode );
}

//===============================================================================================//
//  Description:
//      Search all local hard drives for files of the specified extensions
//
//  Parameters:
//      Extensions        - array of extensions to search for array,  must
//                          include the leading cot (.)
//      DirectoryStopList - a list of directories that will not be searched
//
//  Remarks:
//      Fills class scope list with files
//
//  Returns:
//      void
//===============================================================================================//
void FindFiles::FindOnLocalHardDrives( const StringArray& Extensions,
                                       const StringArray& DirectoryStopList )
{
    UINT    errorMode     = 0, driveType = 0;
    DWORD   logicalDrives = 0;
    wchar_t driveLetter   = 0;
    String  Drive;

    m_Files.RemoveAll();
    if ( Extensions.GetSize() == 0 )
    {
        throw ParameterException( L"Extensions", __FUNCTION__ );
    }

    logicalDrives = GetLogicalDrives();
    if ( logicalDrives == 0 )
    {
        throw SystemException( GetLastError(), L"GetLogicalDrives", __FUNCTION__ );
    }

    // Suppress no disk in drive messages
    errorMode = SetErrorMode( SEM_FAILCRITICALERRORS );
    try
    {
        driveLetter = _T( 'A' );
        while ( logicalDrives )
        {
            // This is a lengthy task so see if want to stop
            if ( g_pApplication && g_pApplication->GetStopBackgroundTasks() )
            {
                break;
            }

            if ( logicalDrives & 0x01 )
            {
                // Make the drive path
                Drive  = driveLetter;
                Drive += PXS_CHAR_COLON;
                Drive += PXS_PATH_SEPARATOR;
                driveType = GetDriveType( Drive.c_str() );
                if ( ( driveType == DRIVE_FIXED  ) &&
                     ( driveType != DRIVE_REMOTE )  )
                {
                    PXSLogAppInfo1( L"Scanning drive '%%1'.", Drive );
                    FindInDirectory( Drive, Extensions, DirectoryStopList );
                }
            }
            driveLetter++;  // Next
            logicalDrives = logicalDrives >> 1;
        }
    }
    catch ( const Exception& )
    {
        SetErrorMode( errorMode );
        throw;
    }
    SetErrorMode( errorMode );
}

//===============================================================================================//
//  Description:
//      Get all the files of a given extension as an array of audit records
//
//  Parameters:
//      Extension - pointer to a string of the extension to search for
//      pRecords  - receives the audit records
//
//  Remarks:
//      Must have already executed a method to find the files
//
//  Returns:
//      void
//===============================================================================================//
void FindFiles::GetFileRecords( const String& Extension, TArray< AuditRecord >* pRecords )
{
    String      FilePath, DotExtension, Drive, Dir, Fname, Ext;
    String      Value;
    Directory   DirObject;
    Formatter   Format;
    AuditRecord Record;
    TYPE_SIZE_DATE_PATH* pElement;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    if ( Extension.IsEmpty() || ( Extension.GetLength() >= _MAX_EXT ) )
    {
        throw ParameterException( L"Extension", __FUNCTION__ );
    }

    // Want the extension with a leading dot for splitpath
    if ( Extension.CharAt( 0 ) != PXS_CHAR_DOT )
    {
        DotExtension += PXS_CHAR_DOT;
    }
    DotExtension += Extension;

    // Scan for the matching extension and add to the input array
    if ( m_Files.IsEmpty() )
    {
        return;     // Nothing more to do
    }

    m_Files.Rewind();
    do
    {
        // See if got stop signal
        if ( g_pApplication && g_pApplication->GetStopBackgroundTasks() )
        {
            break;
        }

        pElement = m_Files.GetPointer();
        FilePath = pElement->szFilePath;
        if ( FilePath.GetLength() )
        {
            Drive = PXS_STRING_EMPTY;
            Dir   = PXS_STRING_EMPTY;
            Fname = PXS_STRING_EMPTY;
            Ext   = PXS_STRING_EMPTY;
            DirObject.SplitPath( FilePath, &Drive, &Dir, &Fname, &Ext );
            if ( DotExtension.CompareI( Ext.c_str() ) == 0 )
            {
                Record.Reset( PXS_CATEGORY_FIND_FILES );

                // Add the extension to the file name
                Fname += Ext;
                Record.Add( PXS_FIND_FILES_FILENAME, Fname );

                Value = Format.StorageBytes( pElement->fileSize );
                Record.Add( PXS_FIND_FILES_SIZE, Value );

                // Use local time zone
                Value = Format.FileTimeToLocalTimeIso( pElement->lastWrite );
                Record.Add( PXS_FIND_FILES_MOD_TIMESTAMP, Value );

                Record.Add( PXS_FIND_FILES_DIRECTORY, Dir );

                pRecords->Add( Record );
            }
        }
    } while ( m_Files.Advance() );
}

//===============================================================================================//
//  Description:
//      Get the executable files found in the search that are not in or
//      below the Windows directory
//
//  Parameters:
//      pRecords - array to receive the audit records
//
//  Remarks:
//      Must have already executed the find to populate the list
//      Version information is included
//
//  Returns:
//      void
//===============================================================================================//
void FindFiles::GetNonWindowsExeRecords( TArray< AuditRecord >* pRecords )
{
    File        ExeFile;
    String      Drive, Dir, Fname, Ext, WindowsDirectoryPath, FilePath;
    String      ManufacturerName, VersionString, Description, Modified;
    Formatter   Format;
    Directory   ExeDirectory;
    FileVersion FileVer;
    StringArray Extensions, DirectoryStopList;
    AuditRecord Record;
    SystemInformation  SystemInfo;
    TYPE_SIZE_DATE_PATH* pElement;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();


    // Scan local hard drives for files with an extension of ".exe". Note,
    // must include the leading dot (.). Do not scan the windows dir
    Extensions.Add( L".exe" );
    SystemInfo.GetWindowsDirectoryPath( &WindowsDirectoryPath );
    DirectoryStopList.Add( WindowsDirectoryPath );
    FindOnLocalHardDrives( Extensions, DirectoryStopList );
    if ( m_Files.IsEmpty() )
    {
        return;     // Nothing more to do
    }

    m_Files.Rewind();
    do
    {
        // See if got stop signal
        if ( g_pApplication && g_pApplication->GetStopBackgroundTasks() )
        {
            break;
        }
        pElement = m_Files.GetPointer();
        FilePath = pElement->szFilePath;
        FilePath.Trim();

        // Catch exceptions so can proceed to the next file
        try
        {
            if ( ExeFile.Exists( FilePath ) )
            {
                Drive = PXS_STRING_EMPTY;
                Dir   = PXS_STRING_EMPTY;
                Fname = PXS_STRING_EMPTY;
                Ext   = PXS_STRING_EMPTY;
                ExeDirectory.SplitPath( FilePath, &Drive, &Dir, &Fname, &Ext );

                // Get the version info, catch exceptions and continue
                ManufacturerName = PXS_STRING_EMPTY;
                VersionString    = PXS_STRING_EMPTY;
                Description      = PXS_STRING_EMPTY;
                FileVer.GetVersion( FilePath, &ManufacturerName, &VersionString, &Description );
                // Make the record
                Fname += Ext;
                Modified = Format.FileTimeToLocalTimeIso( pElement->lastWrite );
                Record.Reset( PXS_CATEGORY_NON_WINDOWS_EXE );
                Record.Add( PXS_NON_WINDOWS_EXE_FILENAME , Fname );
                Record.Add( PXS_NON_WINDOWS_EXE_VERSION  , VersionString );
                Record.Add( PXS_NON_WINDOWS_EXE_MODIFIED , Modified );
                Record.Add( PXS_NON_WINDOWS_EXE_MANUFAC  , ManufacturerName);
                Record.Add( PXS_NON_WINDOWS_EXE_DIRECTORY, Dir );
                pRecords->Add( Record );
            }
        }
        catch ( const Exception& e )
        {
            PXSLogException( e, __FUNCTION__ );
        }
    } while ( m_Files.Advance() );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Computer the file size from the specified WIN32_FIND_DATA structure
//
//  Parameters:
//      pFindFileData - pointer to WIN32_FIND_DATA
//
//  Returns:
//      UINT64
//===============================================================================================//
UINT64 FindFiles::ComputeFileSize( const WIN32_FIND_DATA* pFindFileData )
{
    UINT64 fileSize = 0;

    if ( pFindFileData )
    {
        fileSize = PXSMultiplyUInt64( pFindFileData->nFileSizeHigh, 1ULL << 32);
        fileSize = PXSAddUInt64( fileSize, pFindFileData->nFileSizeLow );
    }

    return fileSize;
}

//===============================================================================================//
//  Description:
//      Determine if the specified file name ends with one of the specified
//      extensions
//
//  Parameters:
//      FileName   - the file name to test
//      Extensions - array of extensions, these must not start with a dot
//
//  Returns:
//      true if the file name ends with one of the extensions otherwise
//      false
//===============================================================================================//
bool FindFiles::FileNameHasExtension( const String& FileName, const StringArray& Extensions )
{
    String Drive, Dir, Fname, Ext;
    Directory DirObject;

    DirObject.SplitPath( FileName, &Drive, &Dir, &Fname, &Ext );
    if ( ( Ext.GetLength() ) &&
         ( PXS_MINUS_ONE != Extensions.IndexOf( Ext.c_str(), false ) ) )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Find files matching a one or more file extensions
//
//  Parameters:
//      DirectoryPath    - directory path to search
//      fileAttributes   - the file attributes of DirectoryPath
//      recursionDepth   - the current recursion depth
//      Extensions       - a string array of file extensions to match, must
//                         include the leading dot (.)
//      DirectoryStopList- these directories will be ignored
//
//  Remarks:
//      The list is not emptied before processing because often what to
//      append the results to a previous scan in a sorted manner.
//
//      Will not follow links (aka junctions or reparse-points that have
//      become available since Windows 2000) i.e. if a symbolic link points
//      to a higher up directory c:\temp\dir1\dir2 -> c:\temp\dir1
//
//  Returns:
//      void
//===============================================================================================//
void FindFiles::Find( const String& DirectoryPath,
                      DWORD fileAttributes,
                      DWORD recursionDepth,
                      const StringArray& Extensions, const StringArray& DirectoryStopList )
{
    HANDLE hSearch   = nullptr;
    String Found, SearchPath, FilePath;
    WIN32_FIND_DATA     FindFileData;
    TYPE_SIZE_DATE_PATH Element;

    recursionDepth++;
    if ( false == FollowDirectory( DirectoryPath,
                                   fileAttributes, recursionDepth, DirectoryStopList ) )
    {
        return;
    }

    // Always want a path separator
    SearchPath = DirectoryPath;
    if ( SearchPath.EndsWithCharacter( PXS_PATH_SEPARATOR ) == false )
    {
        SearchPath += PXS_PATH_SEPARATOR;
    }

    // Make the full search path, use *.* as need to get the directories
    // as well as files.
    SearchPath += L"*.*";

    // Get the first file
    memset( &FindFileData, 0, sizeof ( FindFileData ) );
    hSearch = FindFirstFile( SearchPath.c_str(), &FindFileData  );
    if ( hSearch == INVALID_HANDLE_VALUE )
    {
        // Usually access denied so will avoid throwing so as not to stop
        // the operation in progress
        PXSLogSysWarn1( GetLastError(), L"FindFirstFile failed for '%%1'.", SearchPath );
        return;
    }

    Found.Allocate( MAX_PATH + 1 );
    try     // Catch all exceptions as need to close the handle
    {
        do
        {
            // This is a lengthy task so see if need to stop
            if ( g_pApplication && g_pApplication->GetStopBackgroundTasks() )
            {
                break;
            }

            // N.B. member cFileName may not null terminated
            FindFileData.cFileName[
                       ARRAYSIZE( FindFileData.cFileName ) - 1] = PXS_CHAR_NULL;
            Found  = DirectoryPath;
            if ( Found.EndsWithCharacter( PXS_PATH_SEPARATOR ) == false )
            {
                Found += PXS_PATH_SEPARATOR;
            }
            Found += FindFileData.cFileName;

            // For directories go down
            if ( FILE_ATTRIBUTE_DIRECTORY & FindFileData.dwFileAttributes )
            {
                // Ignore "." and ".."
                if ( FindFileData.cFileName[ 0 ] != '.' )
                {
                    Find( Found,
                          FindFileData.dwFileAttributes,
                          recursionDepth, Extensions, DirectoryStopList);
                }
            }
            else if ( FileNameHasExtension( Found, Extensions ) )
            {
                memset( &Element, 0, sizeof ( Element ) );
                Element.fileSize = ComputeFileSize( &FindFileData );
                Element.lastWrite = FindFileData.ftLastWriteTime;
                StringCchCopy( Element.szFilePath,
                               ARRAYSIZE( Element.szFilePath ),
                               Found.c_str() );  // Has path separator
                m_Files.Append( Element );
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
}

//===============================================================================================//
//  Description:
//      Determine if should follow the directory during a find files scan
//
//  Parameters:
//      DirectoryPath     - the directory
//      fileAttributes    - the directories attributes
//      recursionDepth    - the current recursion depth
//      DirectoryStopList - a stop list of directories not to search
//
//  Returns:
//      void
//===============================================================================================//
bool FindFiles::FollowDirectory( const String& DirectoryPath,
                                 DWORD fileAttributes,
                                 DWORD recursionDepth, const StringArray& DirectoryStopList )
{
    if ( DirectoryPath.IsEmpty() )
    {
        return false;
    }

    // Not following a junction (aka reparse-point)
    if ( ( fileAttributes & FILE_ATTRIBUTE_REPARSE_POINT ) )
    {
        return false;
    }

    // Do not follow UNC paths
    if ( DirectoryPath.CharAt( 0 ) == PXS_PATH_SEPARATOR )
    {
        return false;
    }

    // Do not step into parent, i.e "." or ".."
    if ( DirectoryPath.CharAt( 0 ) == '.' )
    {
        return false;
    }

    // Must be a directory
    if ( !( fileAttributes & FILE_ATTRIBUTE_DIRECTORY ) )
    {
        return false;
    }

    if ( recursionDepth > MAX_RECURSION_DEPTH )
    {
        return false;
    }

    //  Do not follow if in the stop list
    if ( DirectoryStopList.IndexOf( DirectoryPath.c_str(), false ) != PXS_MINUS_ONE )
    {
        return false;
    }

    return true;
}
