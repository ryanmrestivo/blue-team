///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Shell Class Implementation
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
#include "PxsBase/Header Files/Shell.h"

// 2. C System Files
#include <ShlObj.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/AutoIUnknownRelease.h"
#include "PxsBase/Header Files/ComException.h"
#include "PxsBase/Header Files/Directory.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/StringT.h"
#include "PxsBase/Header Files/WaitCursor.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
Shell::Shell()
{
}

// Copy constructor - no allowed so no implementation

// Destructor
Shell::~Shell()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator - no allowed so no implementation

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Create a desktop shortcut
//
//  Parameters:
//      ShortCutName - the name of the short cut i.e. name.lnk
//      TargetDir    - the directory containing the target, i.e. "starts in"
//                     property
//      TargetPath   - full path the target file, ie. file launched by short cut
//      Description  - description of short cut
//
//  Remarks:
//      IPersistFile automatically adds quotes when saving so no need
//      to fix up paths.
//
//  Returns:
//      void
//===============================================================================================//
void Shell::CreateDesktopShortCut( const String& ShortCutName,
                                   const String& WorkingDirectory,
                                   const String& TargetPath, const String& Description )
{
    OLECHAR    szFileName[ MAX_PATH + 1 ] = { 0 };
    String     ShortCutFullPath;
    HRESULT    hResult = 0;
    Directory  DirectoryObject;
    IShellLink*   pShellLink   = nullptr;
    IPersistFile* pPersistFile = nullptr;
    AutoIUnknownRelease ReleasePersistFile, ReleaseShellLink;

    if ( ShortCutName.IsEmpty() )
    {
        throw ParameterException( L"ShortCutName", __FUNCTION__ );
    }

    if ( TargetPath.IsEmpty() )
    {
        throw ParameterException( L"TargetPath", __FUNCTION__ );
    }

    // IShellLink interface.
    hResult = CoCreateInstance( CLSID_ShellLink,
                                nullptr,
                                CLSCTX_INPROC_SERVER,
                                IID_IShellLink, reinterpret_cast<void**>( &pShellLink ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IID_IShellLink", "CoCreateInstance" );
    }

    if ( pShellLink == nullptr )
    {
        throw NullException( L"pShellLink", __FUNCTION__ );
    }
    ReleaseShellLink.Set( pShellLink );

    // Set the shortcut's properties
    pShellLink->SetPath( TargetPath.c_str() );
    if ( WorkingDirectory.GetLength() )
    {
        pShellLink->SetWorkingDirectory( WorkingDirectory.c_str() );
    }

    if ( Description.GetLength() )
    {
        pShellLink->SetDescription( Description.c_str() );
    }

    // IPersistFile
    hResult = pShellLink->QueryInterface( IID_IPersistFile,
                                          reinterpret_cast<void**>( &pPersistFile ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IShellLink::QueryInterface", __FUNCTION__ );
    }

    if ( pPersistFile == nullptr )
    {
        throw NullException( L"pPersistFile", __FUNCTION__ );
    }
    ReleasePersistFile.Set( pPersistFile );

    // Make the full path
    DirectoryObject.GetSpecialDirectory( CSIDL_DESKTOP, &ShortCutFullPath );
    if ( ShortCutFullPath.IsEmpty() )
    {
        throw SystemException( ERROR_INVALID_NAME, L"CSIDL_DESKTOP", __FUNCTION__ );
    }

    ShortCutFullPath += ShortCutName;
    if ( ShortCutFullPath.EndsWithStringI( L".lnk" ) == false )
    {
        ShortCutFullPath += L".lnk";
    }

    StringCchCopy( szFileName, ARRAYSIZE( szFileName ), ShortCutFullPath.c_str() );
    hResult = pPersistFile->Save( szFileName, TRUE );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"PersistFile::Save", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Navigate to a given URL by launching the default browser
//
//  Parameters:
//      Url - the URL to point the browser to
//
//  Returns:
//      void
//===============================================================================================//
void Shell::NavigateToUrl( const String& Url )
{
    bool      fileCreated = false;
    File      HtmlFile;
    size_t    i = 0;
    String    HtmlFilePath, ErrorMessage, ExecutablePath, HtmlContent;
    wchar_t   szResult[ MAX_PATH + 1 ] = { 0 };
    Formatter Format;
    HINSTANCE hInstance = nullptr;

    if ( Url.IsEmpty() )
    {
        return;     // Nothing to do
    }

    // Get or make a temporary html file
    if ( g_pApplication )
    {
        g_pApplication->GetApplicationTempPath( true, &HtmlFilePath );
    }
    if ( HtmlFilePath.IsEmpty() )
    {
        PXSGetExeDirectory( &HtmlFilePath );
    }

    // Create a temp html file
    HtmlFilePath += L"temp.html";
    if ( HtmlFile.Exists( HtmlFilePath ) == false )
    {
        HtmlFile.CreateNew( HtmlFilePath, 0, false );
        HtmlContent = L"<html><body><p>Parmavex Services</p></body></html>";
        HtmlFile.WriteChars( HtmlContent );
        HtmlFile.Close();
        fileCreated = true;
    }

    // Get the default browser, make sure result is NULL terminated
    WaitCursor waitCursor;
    memset( szResult, 0, sizeof ( szResult ) );
    hInstance = FindExecutable( HtmlFilePath.c_str(),
                                nullptr,
                                szResult );  // Must be at lease MAX_PATH
    if ( hInstance <= (HINSTANCE)32 )
    {
        // Delete temp file, if created
        if ( fileCreated )
        {
            HtmlFile.Delete( HtmlFilePath );
        }
        ErrorMessage  = L"FindExecutable, ";
        ErrorMessage += HtmlFilePath;
        ErrorMessage += L", hInstance = ";
        ErrorMessage += Format.Pointer( hInstance );
        throw SystemException( ERROR_BAD_PATHNAME, ErrorMessage.c_str(), __FUNCTION__ );
    }
    szResult[ ARRAYSIZE( szResult ) - 1 ] = PXS_CHAR_NULL;

    // Replace nulls returned by FindExecutable to spaces
    // as a result of non-quoted strings in registry
    for ( i = 0; i < ARRAYSIZE( szResult ) - 1; i++ )
    {
        if ( szResult[ i ] == PXS_CHAR_NULL )
        {
            szResult[ i ] = PXS_CHAR_SPACE;
        }
    }
    szResult[ ARRAYSIZE( szResult ) - 1 ] = PXS_CHAR_NULL;
    ExecutablePath = szResult;
    ExecutablePath.Trim();
    hInstance = ShellExecute( GetDesktopWindow(),
                              L"open",
                              ExecutablePath.c_str(), Url.c_str(), nullptr, SW_SHOWDEFAULT );
    // Delete temp file first...
    if ( fileCreated )
    {
        HtmlFile.Delete( HtmlFilePath );
    }

    // ... then test for error
    if ( hInstance <= (HINSTANCE)32 )
    {
        ErrorMessage  = L"ShellExecute, open, ";
        ErrorMessage += ExecutablePath;
        ErrorMessage += L", hInstance = ";
        ErrorMessage += Format.Pointer( hInstance );
        throw SystemException( ERROR_BAD_PATHNAME, ErrorMessage.c_str(), __FUNCTION__ );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
