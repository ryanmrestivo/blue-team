///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Executable Library Class Implementation
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
#include "PxsBase/Header Files/Library.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/StringT.h"
#include "PxsBase/Header Files/SystemInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
Library::Library()
        :m_hLibrary( nullptr )
{
}

// Copy constructor - not allowed so no implementation

// Destructor
Library::~Library()
{
    // Clean up
    if ( m_hLibrary )
    {
        FreeLibrary( m_hLibrary );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

//
// Assignment operator - not allowed so no implementation
//

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Free the library, if it was loaded
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Library::Free()
{
    if ( m_hLibrary )
    {
        FreeLibrary( m_hLibrary );
        m_hLibrary = nullptr;
    }
}

//===============================================================================================//
//  Description:
//      Get the handle to the library
//
//  Parameters:
//      None
//
//  Remarks:
//      Caller must not free the library, this is done in the destructor
//
//  Returns:
//      HINSTANCE
//===============================================================================================//
HINSTANCE Library::GetInstance()
{
    return m_hLibrary;
}

//===============================================================================================//
//  Description:
//      Load a system library
//
//  Parameters:
//      pszLibFileName - pointer to name of library
//
//  Returns:
//      void
//===============================================================================================//
void Library::LoadSystemLibrary( LPCWSTR pszLibFileName )
{
    String LibFileName, LibFullPath;
    SystemInformation SystemInfo;

    // Want a name only, not a path
    LibFileName = pszLibFileName;
    LibFileName.Trim();
    if ( ( LibFileName.IsEmpty() ) ||
         ( LibFileName.IndexOf( PXS_PATH_SEPARATOR, 0 ) != PXS_MINUS_ONE ) )
    {
        throw ParameterException( L"pszLibFileName", __FUNCTION__ );
    }

    SystemInfo.GetSystemDirectoryPath( &LibFullPath );
    LibFullPath += LibFileName;
    LoadFullPath( LibFullPath );
}

//===============================================================================================//
//  Description:
//      Load a library
//
//  Parameters:
//      pszLibFullPath - full path to the library
//
//  Remarks:
//      Suppresses and error messages, in case cannot find library. This can
//      happen if LoadLibrary finds a library not for the computer's operating
//      system. E.g. Attempt to load a library mistakenly put on the wrong
//      operating system.
//
//  Returns:
//      void
//===============================================================================================//
void Library::LoadFullPath( const String& LibFullPath )
{
    UINT  uMode;
    DWORD lastError;

    // Test for existing handle
    if ( m_hLibrary )
    {
        throw FunctionException( L"m_hLibrary", __FUNCTION__ );
    }

    // Make sure it has a path separator
    if ( ( LibFullPath.IsEmpty() ) ||
         ( LibFullPath.IndexOf( PXS_PATH_SEPARATOR, 0 ) == PXS_MINUS_ONE ) )
    {
        throw ParameterException( L"LibFullPath", __FUNCTION__ );
    }

    // Suppress cannot load messages
    uMode = SetErrorMode( SEM_NOOPENFILEERRORBOX );
    SetLastError( ERROR_SUCCESS );
    m_hLibrary = LoadLibrary( LibFullPath.c_str() );
    lastError  = GetLastError();
    SetErrorMode( uMode );
    if ( m_hLibrary == nullptr )
    {
        throw SystemException( lastError, L"LoadLibrary", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      The the address of a procedure in a library, i.e. same as
//      GetProcAddress().
//
//  Parameters:
//      pszProcName - pointer to name of procedure, case sensitive
//
//  Remarks:
//      Library must already have been loaded
//
//  Returns:
//      Pointer to the procedure, NULL is not found
//===============================================================================================//
FARPROC Library::ProcAddress( const char* pszProcName )
{
    String  Details;
    FARPROC pAddress = nullptr;

    if ( m_hLibrary == nullptr )
    {
        throw FunctionException( L"m_hLibrary", __FUNCTION__ );
    }

    // Check input
    if ( ( pszProcName == nullptr ) || ( *pszProcName == PXS_CHAR_NULL ) )
    {
        throw ParameterException( L"pszProcName", __FUNCTION__ );
    }

    pAddress = GetProcAddress( m_hLibrary, pszProcName );
    if ( pAddress == nullptr )
    {
        Details.SetAnsi( pszProcName );
        throw SystemException( GetLastError(), Details.c_str(), __FUNCTION__ );
    }
    return pAddress;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
