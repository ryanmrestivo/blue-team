///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Access Database Class Implementation
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
#include "WinAudit/Header Files/AccessDatabase.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/AutoIUnknownRelease.h"
#include "PxsBase/Header Files/AutoSysFreeString.h"
#include "PxsBase/Header Files/AutoVariantClear.h"
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/ComException.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"

// 5. This Project

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
AccessDatabase::AccessDatabase()
               :m_DatabasePath(),
                m_UserName(),
                m_Password(),
                m_ProviderID(),
                m_bConnected( false ),
                m_pIDBInitialize( nullptr ),
                m_bDatabaseOpen( false ),
                acCmdNewObjectModule( 139 ),
                acModule( 5 ),
                acSaveYes( 1 ),
                m_pIUnknown( nullptr ),
                m_pDispAccess( nullptr )
{
}

// Copy constructor - not allowed so no implementation

// Destructor
AccessDatabase::~AccessDatabase()
{
    try
    {
        // Clean up
        ComQuit( 1 );   // acQuitSaveAll = 1
        OleDbDisconnect();
        if ( m_pIDBInitialize )
        {
            m_pIDBInitialize->Release();
        }
    }
    catch ( const Exception& e )
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
//      Close an open database using Access.CloseCurrentDatabase()
//
//  Parameters:
//      None
//
//  Remarks:
//      Only closes the opened database, does not quit Access
//      Access syntax:
//          Sub CloseCurrentDatabase()
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::ComCloseCurrentDatabase()
{
    VARIANT    result;
    DISPPARAMS dispParams = { nullptr, nullptr, 0, 0 };     // No arguments

    if ( m_bDatabaseOpen == false )
    {
        return;
    }

    if ( m_pDispAccess == nullptr )
    {
        return;
    }

    VariantInit( &result );
    ComMethod( m_pDispAccess, L"CloseCurrentDatabase", &dispParams, &result );
    VariantClear( &result );

    m_bDatabaseOpen = false;
}

//===============================================================================================//
//  Description:
//      Runs a built-in menu or tool bar command using Access.DoCmd.RunCommand
//
//  Parameters:
//      command - named constant of the command to run
//
//  Remarks:
//      Access Syntax
//          DoCmd.RunCommand Command
//
//      Do not need to check if a database has been opened as many commands
//      are available when no database is open
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::ComDoCmd_RunCommand( long command )
{
    VARIANT    doCmd;
    VARIANT    result;
    VARIANTARG arguments[ 1 ];    // DoCmd.Close takes 1 arguments
    DISPPARAMS dispParams = { nullptr, nullptr, 0, 0 };

    // Cannot execute a command unless Access has been started
    if ( m_pDispAccess == nullptr )
    {
       throw FunctionException( L"m_pDispAccess", __FUNCTION__ );
    }

    // Get a dispatch interface to the DoCmd
    VariantInit( &doCmd );
    doCmd.vt = VT_DISPATCH;
    ComProperyGet( m_pDispAccess, L"DoCmd", &dispParams, &doCmd );
    AutoVariantClear VariantClearDoCmd( &doCmd );

    // Assign the Command argument
    VariantInit( &arguments[ 0 ] );
    arguments[ 0 ].vt   = VT_I4;
    arguments[ 0 ].lVal = command;

    // Set the dispatch parameters structure
    memset( &dispParams, 0, sizeof ( dispParams ) );
    dispParams.cArgs  = ARRAYSIZE( arguments );
    dispParams.rgvarg = arguments;

    // Call Access.DoCmd.RunCommand()
    VariantInit( &result );
    ComMethod( doCmd.pdispVal, L"RunCommand", &dispParams, &result );
    VariantClear( &result );
}

//===============================================================================================//
//  Description:
//      Quit Access
//
//  Parameters:
//      options - options on quitting
//
//  Remarks:
//      Access Syntax
//          Sub Quit([Option As AcQuitOption = acQuitSaveAll])
//          Const acQuitPrompt   = 0
//          Const acQuitSaveAll  = 1
//          Const acQuitSaveNone = 2
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::ComQuit( long option )
{
    VARIANT    result;
    VARIANTARG arguments[ 1 ];           // Quit() takes one argument
    DISPPARAMS dispParams = { nullptr, nullptr, 0, 0 };

    // If Access has not been started will not raise an error
    if ( m_pDispAccess == nullptr )
    {
        return;
    }

    ComCloseCurrentDatabase();

    // Assign the exit option
    arguments[ 0 ].vt   = VT_I4;
    arguments[ 0 ].lVal = option;

    dispParams.cArgs  = ARRAYSIZE( arguments );
    dispParams.rgvarg = arguments;

    VariantInit( &result );
    ComMethod( m_pDispAccess, L"Quit", &dispParams, &result );
    VariantClear( &result );

    // Release the interfaces
    if ( m_pDispAccess )
    {
        m_pDispAccess->Release();
        m_pDispAccess = nullptr;  // Reset
    }

    if ( m_pIUnknown )
    {
        m_pIUnknown->Release();
        m_pIUnknown = nullptr;  // Reset
    }
}

//===============================================================================================//
//  Description:
//      Start an instance of Microsoft Access. This creates the class scope
//      dispatch interface.
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::ComStart()
{
    CLSID   accessClassID;
    HRESULT hResult = S_OK;

    // If Access has been started will not raise an error
    if ( m_pDispAccess )
    {
        return;     // Nothing to do
    }

    // In order to create the COM interfaces need to first identify the
    // class ID of an installed instance of Access
    memset( &accessClassID, 0, sizeof ( accessClassID ) );
    GetAccessClassID( &accessClassID );

    // Create an instance of Access if have not already done so
    if ( m_pIUnknown == nullptr )
    {
        // Will use flags CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER to be
        // consistent will OleRun
        hResult = CoCreateInstance( accessClassID,
                                    nullptr,
                                    CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
                                    IID_IUnknown,
                                    reinterpret_cast<void**>( &m_pIUnknown ) );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"CoCreateInstance, IID_IUnknown", __FUNCTION__ );
        }

        if ( m_pIUnknown == nullptr )
        {
            throw NullException( L"m_pIUnknown", __FUNCTION__ );
        }

        // Should be running, but ensure by putting in a runnable state
        // Note, OleRun calls IRunnableObject::Run() which uses
        // CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER
        hResult = OleRun( m_pIUnknown );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"OleRun", __FUNCTION__ );
        }
    }

    hResult = m_pIUnknown->QueryInterface( IID_IDispatch,
                                           reinterpret_cast<void**>( &m_pDispAccess ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CoCreateInstance, IID_IDispatch", __FUNCTION__ );
    }

    if ( m_pDispAccess == nullptr )
    {
        throw NullException( L"m_pDispAccess", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Show/Hide the application using using its Access.Visible property
//
//  Parameters:
//      visible - flag that indicates if want the application to be visible
//
//  Remarks:
//      Access Syntax
//          Property Visible As Boolean
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::ComVisible( bool visible )
{
    DISPID     dispIDNamedArgs = 0;
    VARIANT    result;
    VARIANTARG arguments[ 1 ];
    DISPPARAMS dispParams = { nullptr, nullptr, 0, 0 };

    // Cannot execute a command unless Access has been started
    if ( m_pDispAccess == nullptr )
    {
       throw FunctionException( L"m_pDispAccess", __FUNCTION__ );
    }

    VariantInit( &arguments[ 0 ] );
    arguments[ 0 ].vt = VT_BOOL;
    if ( visible )
    {
        arguments[ 0 ].boolVal = TRUE;    // -1
    }
    else
    {
        arguments[ 0 ].boolVal = FALSE;
    }

    memset( &dispParams, 0, sizeof ( dispParams ) );
    dispParams.cArgs  = 1;
    dispParams.rgvarg = arguments;

    // When putting a property need to set it as a named argument and use
    // DISPID_PROPERTYPUT as the dispatch ID of the named argument
    dispIDNamedArgs = DISPID_PROPERTYPUT;
    dispParams.cNamedArgs = 1;
    dispParams.rgdispidNamedArgs = &dispIDNamedArgs;

    VariantInit( &result );
    ComProperyPut( m_pDispAccess, L"Visible", &dispParams, &result );
    VariantClear( &result );
}

//===============================================================================================//
//  Description:
//      Create the class scope IDBInitialize interface
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::CreateIDBInitialize()
{
    CLSID    clsidJetOleDb;
    HRESULT  hResult = S_OK;

    // Cannot create this interface without a provider name
    if ( m_ProviderID.IsEmpty() )
    {
       throw FunctionException( L"m_ProviderID", __FUNCTION__ );
    }

    // Only need to create if do not have one as this object can be reused
    if ( m_pIDBInitialize == nullptr )
    {
        memset( &clsidJetOleDb, 0, sizeof ( clsidJetOleDb ) );
        hResult = CLSIDFromProgID( m_ProviderID.c_str(), &clsidJetOleDb );
        if ( hResult != NOERROR )
        {
           throw ComException( hResult, m_ProviderID.c_str(), "CLSIDFromProgID" );
        }

        hResult = CoCreateInstance(
                             clsidJetOleDb,
                             nullptr,
                             CLSCTX_INPROC_SERVER,
                             IID_IDBInitialize,
                             reinterpret_cast<void**>( &m_pIDBInitialize ) );
        if ( FAILED( hResult ) )
        {
            throw ComException( hResult, L"CoCreateInstance, IDBInitialize", __FUNCTION__ );
        }

        if ( m_pIDBInitialize == nullptr )
        {
            throw NullException( L"m_pIDBInitialize", __FUNCTION__ );
        }
    }
}

//===============================================================================================//
//  Description:
//      Create an empty Access Database at the specified path
//
//  Parameters:
//      DatabasePath - the full path to the database;
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::OleDbCreate( const String& DatabasePath )
{
    const DWORD NUM_DBPROPS = 1;
    CLSID     clsidJetOleDb;
    BStr      BStrPath;
    File      DatabaseFile;
    String    OleDbErrorString;
    DBPROP    rgProps[ NUM_DBPROPS ];
    HRESULT   hResult = S_OK;
    Formatter Format;
    DBPROPSET PropSet;
    IDBInitialize*      pIDBInitialize      = nullptr;
    IDBDataSourceAdmin* pIDBDataSourceAdmin = nullptr;
    AutoIUnknownRelease ReleaseIDBInitialize;
    AutoIUnknownRelease ReleaseIDBDataSourceAdmin;

    // Cannot execute a command unless Access has been started
    if ( m_ProviderID.IsEmpty() )
    {
       throw FunctionException( L"m_ProviderID", __FUNCTION__ );
    }

    // Need a path
    if ( DatabasePath.IsEmpty() )
    {
        throw ParameterException( L"DatabasePath", __FUNCTION__ );
    }

    // Verify that a file does not exist here
    if ( DatabaseFile.Exists( DatabasePath ) )
    {
        throw SystemException( ERROR_FILE_EXISTS, DatabasePath.c_str(), __FUNCTION__ );
    }

    // Need the provider's class id
    memset( &clsidJetOleDb, 0, sizeof ( clsidJetOleDb ) );
    hResult = CLSIDFromProgID( m_ProviderID.c_str(), &clsidJetOleDb );
    if ( hResult != NOERROR )
    {
        throw ComException( hResult, m_ProviderID.c_str(), __FUNCTION__ );
    }

    // Different databases/versions have different constants for the CLSID.
    // For ORACLE its CLSID_MSDAORA
    hResult = CoCreateInstance( clsidJetOleDb,
                                nullptr,
                                CLSCTX_INPROC_SERVER,
                                IID_IDBInitialize,
                                reinterpret_cast<void**>( &pIDBInitialize) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IID_IDBInitialize", __FUNCTION__ );
    }

    if ( pIDBInitialize == nullptr )
    {
        throw NullException( L"pIDBInitialize", __FUNCTION__ );
    }
    ReleaseIDBInitialize.Set( pIDBInitialize );

    // Get the DataSourceAdmin object
    hResult = pIDBInitialize->QueryInterface( IID_IDBDataSourceAdmin,
                                              reinterpret_cast<void**>( &pIDBDataSourceAdmin ) );
    if ( FAILED( hResult ) )
    {
        GetOleDbErrorString( &OleDbErrorString );
        throw ComException( hResult, OleDbErrorString.c_str(), "IDBInitialize::QueryInterface" );
    }

    if ( pIDBDataSourceAdmin == nullptr )
    {
        throw NullException( L"pIDBDataSourceAdmin", __FUNCTION__ );
    }
    ReleaseIDBDataSourceAdmin.Set( pIDBDataSourceAdmin );

    // Need a BSTR for the path in order to set it as a property
    BStrPath.Allocate( DatabasePath );
    VariantInit( &rgProps[ 0 ].vValue );
    rgProps[ 0 ].dwPropertyID   = DBPROP_INIT_DATASOURCE;
    rgProps[ 0 ].vValue.vt      = VT_BSTR;
    rgProps[ 0 ].dwOptions      = DBPROPOPTIONS_REQUIRED;
    rgProps[ 0 ].colid          = DB_NULLID;
    rgProps[ 0 ].vValue.bstrVal = BStrPath.b_str();

    PropSet.rgProperties    = rgProps;
    PropSet.cProperties     = NUM_DBPROPS;
    PropSet.guidPropertySet = DBPROPSET_DBINIT;

    hResult = pIDBDataSourceAdmin->CreateDataSource( 1,
                                                     &PropSet,
                                                     nullptr, IID_IDBDataSourceAdmin, nullptr );
    if ( FAILED( hResult ) )
    {
        GetOleDbErrorString( &OleDbErrorString );
        throw ComException( hResult,
                            OleDbErrorString.c_str(), "IDBDataSourceAdmin::CreateDataSource" );
    }
}

//===============================================================================================//
//  Description:
//      Connect to the specified database using the specified credentials
//
//  Parameters:
//      DatabsePath - the file path to the database
//      UserName    - the user name
//      Password    - the password
//
//  Remarks:
//      If no user name is supplied, then will use "Admin" otherwise get error
//      DB_SEC_E_AUTH_FAILED.
//
//      The BSTR strings are stored at class scope in case need their string
//      pointers to be valid while the connection is active.
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::OleDbConnect( const String& DatabsePath,
                                   const String& UserName, const String& Password )
{
    const DWORD NUM_DBPROPS = 3;
    String     OleDbErrorString, UserNameClean, PasswordClean;
    DBPROP     rgProps[ NUM_DBPROPS ];
    HRESULT    hResult = S_OK;
    Formatter  Format;
    DBPROPSET  PropSet;
    IDBProperties* pIDBProperties = nullptr;
    AutoIUnknownRelease ReleaseIDBProperties;

    // Do a disconnect even if not connected to reset class scope members
    OleDbDisconnect();
    CreateIDBInitialize();
    if ( m_pIDBInitialize == nullptr )
    {
        throw NullException( PXS_STRING_EMPTY, __FUNCTION__ );
    }

    // If no user name was supplied then use "Admin" otherwise get
    // DB_SEC_E_AUTH_FAILED
    UserNameClean = UserName;
    UserNameClean.Trim();
    if ( UserNameClean.IsEmpty() )
    {
        UserNameClean = L"Admin";
    }

    // If no password was supplied, then user ""
    PasswordClean = Password;
    if ( PasswordClean.IsEmpty() )
    {
        PasswordClean = PXS_STRING_EMPTY;
    }

    // Data source - allocate a BSTR and set it as a property
    m_DatabasePath.Allocate( DatabsePath );
    VariantInit( &rgProps[ 0 ].vValue );
    rgProps[ 0 ].dwPropertyID   = DBPROP_INIT_DATASOURCE;   // The db path
    rgProps[ 0 ].vValue.vt      = VT_BSTR;
    rgProps[ 0 ].dwOptions      = DBPROPOPTIONS_REQUIRED;   // Required
    rgProps[ 0 ].colid          = DB_NULLID;
    rgProps[ 0 ].vValue.bstrVal = m_DatabasePath.b_str();

    // UserName - allocate a BSTR and set it as a property
    m_UserName.Allocate( UserNameClean );
    VariantInit( &rgProps[ 1 ].vValue );
    rgProps[ 1 ].dwPropertyID   = DBPROP_AUTH_USERID;
    rgProps[ 1 ].vValue.vt      = VT_BSTR;
    rgProps[ 1 ].dwOptions      = DBPROPOPTIONS_OPTIONAL;
    rgProps[ 1 ].colid          = DB_NULLID;
    rgProps[ 1 ].vValue.bstrVal = m_UserName.b_str();

    // Password - allocate a BSTR and set it as a property
    m_Password.Allocate( PasswordClean );
    VariantInit( &rgProps[ 2 ].vValue );
    rgProps[ 2 ].dwPropertyID   = DBPROP_AUTH_PASSWORD;
    rgProps[ 2 ].vValue.vt      = VT_BSTR;
    rgProps[ 2 ].dwOptions      = DBPROPOPTIONS_OPTIONAL;
    rgProps[ 2 ].colid          = DB_NULLID;
    rgProps[ 2 ].vValue.bstrVal = m_Password.b_str();

    // Fill the structure containing the properties.
    PropSet.rgProperties    = rgProps;
    PropSet.cProperties     = NUM_DBPROPS;
    PropSet.guidPropertySet = DBPROPSET_DBINIT;

    // Need IID_IDBProperties in order to set the initialization properties.
    hResult = m_pIDBInitialize->QueryInterface( IID_IDBProperties,
                                                reinterpret_cast<void**>( &pIDBProperties ) );
    if ( FAILED( hResult ) )
    {
        GetOleDbErrorString( &OleDbErrorString );
        throw ComException( hResult,
                            OleDbErrorString.c_str(), "QueryInterface, IID_IDBProperties" );
    }

    if ( pIDBProperties == nullptr )
    {
        throw NullException( L"pIDBProperties", __FUNCTION__ );
    }
    ReleaseIDBProperties.Set( pIDBProperties );

    // Now can set the properties
    hResult = pIDBProperties->SetProperties( 1, &PropSet );
    if ( FAILED( hResult ) )
    {
        GetOleDbErrorString( &OleDbErrorString );
        throw ComException( hResult, OleDbErrorString.c_str(), "IDBProperties::SetProperties" );
    }

    hResult = m_pIDBInitialize->Initialize();
    if ( FAILED( hResult ) )
    {
        GetOleDbErrorString( &OleDbErrorString );
        throw ComException( hResult, OleDbErrorString.c_str(), "IDBInitialize::Initialize" );
    }

    m_bConnected = true;
}

//===============================================================================================//
//  Description:
//      Disconnect from the database
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::OleDbDisconnect()
{
    // Un-initialize the DB object, so can use it later
    if ( m_pIDBInitialize )
    {
        m_pIDBInitialize->Uninitialize();
    }
    m_bConnected = false;
}

//===============================================================================================//
//  Description:
//      Execute a command
//
//  Parameters:
//      SqlQuery - the SQL command
//
//  Remarks:
//      Does not return any records
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::OleDbExecute( const String& SqlQuery )
{
    HRESULT   hResult = S_OK;
    String    OleDbErrorString;
    Formatter Format;
    ICommandText*     pICommandText     = nullptr;
    IDBCreateSession* pIDBCreateSession = nullptr;
    IDBCreateCommand* pIDBCreateCommand = nullptr;
    AutoIUnknownRelease ReleaseICommandText;
    AutoIUnknownRelease ReleaseIDBCreateSession;
    AutoIUnknownRelease ReleaseIDBCreateCommand;

    // Class scope validation
    if ( m_pIDBInitialize == nullptr )
    {
        throw FunctionException( L"m_pIDBInitialize", __FUNCTION__ );
    }

    // Get the DB session object.
    hResult = m_pIDBInitialize->QueryInterface( IID_IDBCreateSession,
                                                reinterpret_cast<void**>( &pIDBCreateSession ) );
    if ( FAILED( hResult ) )
    {
        GetOleDbErrorString( &OleDbErrorString );
        throw ComException( hResult,
                            OleDbErrorString.c_str(), "QueryInterface, IID_IDBCreateSession" );
    }

    if ( pIDBCreateSession == nullptr )
    {
        throw NullException( L"pIDBCreateSession", __FUNCTION__ );
    }
    ReleaseIDBCreateSession.Set( pIDBCreateSession );

    // Create the interface for command creation.
    hResult = pIDBCreateSession->CreateSession( nullptr,
                                                IID_IDBCreateCommand,
                                                reinterpret_cast<IUnknown**>(&pIDBCreateCommand) );
    if ( FAILED( hResult ) )
    {
        GetOleDbErrorString( &OleDbErrorString );
        throw ComException( hResult,
                            OleDbErrorString.c_str(), "CreateSession, IID_IDBCreateCommand" );
    }

    if ( pIDBCreateCommand == nullptr )
    {
        throw NullException( L"pIDBCreateCommand", __FUNCTION__ );
    }
    ReleaseIDBCreateCommand.Set( pIDBCreateCommand );

    // Create the interface for the command text
    hResult = pIDBCreateCommand->CreateCommand( nullptr,
                                                IID_ICommandText,
                                                reinterpret_cast<IUnknown**>( &pICommandText) );
    if ( FAILED( hResult ) )
    {
        GetOleDbErrorString( &OleDbErrorString );
        throw ComException( hResult, OleDbErrorString.c_str(), "CreateCommand, IID_ICommandText" );
    }

    if ( pICommandText == nullptr )
    {
        throw NullException( L"pICommandText", __FUNCTION__ );
    }
    ReleaseICommandText.Set( pICommandText );

    // Set it
    hResult = pICommandText->SetCommandText( DBGUID_DEFAULT, SqlQuery.c_str() );
    if ( FAILED( hResult ) )
    {
        GetOleDbErrorString( &OleDbErrorString );
        throw ComException( hResult, OleDbErrorString.c_str(), "ICommandText::SetCommandText" );
    }
    PXSLogAppInfo1( L"ICommandText::Execute: '%%1'", SqlQuery );

    // Execute, no rows are returned so can use IID_NULL
    hResult = pICommandText->Execute( nullptr,          // pUnkOuter
                                      IID_NULL,         // riid
                                      nullptr,          // pParams
                                      nullptr,          // pcRowsAffected
                                      nullptr );        // ppRowset
    if ( FAILED( hResult ) )
    {
        GetOleDbErrorString( &OleDbErrorString );
        OleDbErrorString += PXS_STRING_CRLF;
        OleDbErrorString += L"SQL:\r\n";
        OleDbErrorString += SqlQuery;
        throw ComException( hResult, OleDbErrorString.c_str(), "ICommandText::Execute");
    }
}

//===============================================================================================//
//  Description:
//      Get if currently connected to the database
//
//  Parameters:
//      None
//
//  Returns:
//      true if connected, otherwise false
//===============================================================================================//
bool AccessDatabase::OleDbIsConnected() const
{
    return m_bConnected;
}

//===============================================================================================//
//  Description:
//      Set the provider string, aka ProgID
//
//  Parameters:
//      ProviderID - the provider
//
//  Remarks:
//      This is used to identify the CLSID in the registry with CLSIDFromProgID
//      e.g. Access 97/Jet 3.51   = Microsoft.Jet.OLEDB.3.51
//           Access 2000/Jet 4.0  = Microsoft.Jet.OLEDB.4.0
//           Access 2007/Ace 12.0 = Microsoft.Ace.OleDB.12.0
//           Access 2010/Ace 14.0 = Microsoft.Ace.OleDB.14.0
//
//  Returns:
//     void
//===============================================================================================//
void AccessDatabase::OleDbSetProvider( const String& ProviderID )
{
    m_ProviderID = ProviderID;
    m_ProviderID.Trim();
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Call IDispatch::Invoke() to access properties and methods exposed by
//      an object.
//
//  Parameters:
//      pDispatch   - pointer to the IDispatch interface
//      pwzName     - pointer to the name of the method or property
//      flags       - flags for the IDispatch::Invoke method
//      pDispParams - pointer to the dispatch parameter array
//      pResult     - pointer to a variant to receive the result
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::ComInvoke( IDispatch* pDispatch,
                                LPCWSTR pwzName,
                                WORD flags, DISPPARAMS* pDispParams, VARIANT* pResult )
{
    UINT      argErr   = 0;
    size_t    numChars = 0;
    String    ErrorMessage;
    DISPID    dispID   = 0;
    HRESULT   hResult  = S_OK;
    OLECHAR   wzOleName[ 512 ] = { 0 };     // Big enough for a name
    LPOLESTR  pOleName = nullptr;
    Formatter Format;
    EXCEPINFO excepInfo;

    // Validate
    if ( ( pwzName     == nullptr ) ||
         ( pDispatch   == nullptr ) ||
         ( pDispParams == nullptr )  )
    {
        throw NullException( PXS_STRING_EMPTY, __FUNCTION__ );
    }

    // Must have a name, make sure its not too big, copy it to a wide buffer
    numChars = wcslen( pwzName );
    if ( ( numChars == 0 ) ||
         ( numChars >= ARRAYSIZE( wzOleName ) ) )
    {
        throw ParameterException( L"pszName", __FUNCTION__ );
    }
    StringCchCopy( wzOleName, ARRAYSIZE( wzOleName ), pwzName );

    // Get the ID for this method/property name
    pOleName = wzOleName;
    hResult  = pDispatch->GetIDsOfNames( IID_NULL, &pOleName, 1, LOCALE_USER_DEFAULT, &dispID );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IDispatch::GetIDsOfNames", __FUNCTION__ );
    }

    // Invoke the method. Set to -1 as it receives the index of the first
    // incorrect argument
    memset( &excepInfo, 0, sizeof ( excepInfo ) );
    argErr  = UINT32_MAX;     // -1
    hResult = pDispatch->Invoke( dispID,
                                 IID_NULL,
                                 LOCALE_SYSTEM_DEFAULT,
                                 flags, pDispParams, pResult, &excepInfo, &argErr );
    if ( FAILED( hResult ) )
    {
        // Test for exception information
        if ( hResult == DISP_E_EXCEPTION )
        {
            ErrorMessage  = L"hResult        : DISP_E_EXCEPTION\r\n";
            ErrorMessage += L"bstrSource     : ";
            ErrorMessage += excepInfo.bstrSource;
            ErrorMessage += PXS_STRING_CRLF;
            ErrorMessage += L"bstrDescription: ";
            ErrorMessage += excepInfo.bstrDescription;
            ErrorMessage += PXS_STRING_CRLF;
            ErrorMessage += L"scode          : ";
            ErrorMessage += Format.Int32( excepInfo.scode );
            ErrorMessage += PXS_STRING_CRLF;
            ErrorMessage += L"wCode          : ";
            ErrorMessage += Format.Int32( excepInfo.wCode );
            ErrorMessage += PXS_STRING_CRLF;
        }

        // On DISP_E_TYPEMISMATCH or DISP_E_PARAMNOTFOUND uArgErr has
        // the argument index
        ErrorMessage += L"uArgErr        : ";
        ErrorMessage += Format.UInt32( argErr );
        ErrorMessage += PXS_STRING_CRLF;
        throw ComException( hResult, ErrorMessage.c_str(), "IDispatch::Invoke" );
    }
}

//===============================================================================================//
//  Description:
//      Invoke a method using COM
//
//  Parameters:
//      pDispatch   - pointer to the IDispatch interface
//      pwzName     - pointer to the name of the method
//      pDispParams - pointer to the dispatch parameter array
//      pResult     - pointer to a variant to receive the result
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::ComMethod( IDispatch* pDispatch,
                                LPCWSTR pwzName, DISPPARAMS* pDispParams, VARIANT* pResult )
{
    String Name;

    Name = pwzName;
    ComInvoke( pDispatch, pwzName, DISPATCH_METHOD, pDispParams, pResult );
}


//===============================================================================================//
//  Description:
//      Get a property using COM
//
//  Parameters:
//      pDispatch   - pointer to the IDispatch interface
//      lpName      - pointer to the name of the property
//      pDispParams - pointer to the dispatch parameter array
//      pVarResult  - pointer to a variant to receive the result
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::ComProperyGet( IDispatch* pDispatch,
                                    LPCWSTR pszName, DISPPARAMS* pDispParams, VARIANT* pResult )
{
    ComInvoke( pDispatch, pszName, DISPATCH_PROPERTYGET, pDispParams, pResult);
}

//===============================================================================================//
//  Description:
//      Put a property using COM
//
//  Parameters:
//      pDispatch   - pointer to the IDispatch interface
//      pszName   - pointer to the name of the property
//      pDispParams - pointer to the dispatch parameter array
//      pVarResult  - pointer to a variant to receive the result
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::ComProperyPut( IDispatch* pDispatch,
                                    LPCWSTR pszName, DISPPARAMS* pDispParams, VARIANT* pResult )
{
    ComInvoke( pDispatch,
               pszName, DISPATCH_PROPERTYPUT, pDispParams, pResult );
}

//===============================================================================================//
//  Description:
//      Get the registered CLSID of Access
//
//  Parameters:
//      pAccessClassID - receives the class id
//
//  Remarks:
//      Need to get the class identifier in order to create the dispatch
//      interface. These are stored in the registry at HKEY_CLASSES_ROOT
//      \Access.Application and Application.x The key Access.Application
//      has a subkey of CurVer which shows the version in use.
//      Version numbers for Access are:
//          Access 97       = v8
//          Access 2000     = v9
//          Access 2002(XP) = v10
//          Access 2003     = v11
//          Access 2007     = v12
//          Access 2010     = v14
//          Access 2013     = v15
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::GetAccessClassID( CLSID* pAccessClassID )
{
    const int MAX_ACCESS_CURVER = 25;   // i.e. up to Access.Application.25
    int       i;
    HRESULT   hResult = 0;
    wchar_t   wzCurVer[ 8 ];    // Enough for one or two digit CurVer
    OLECHAR   wzProgID[ 64 ];   // Enough for the programme ID

    if ( pAccessClassID == nullptr )
    {
        throw ParameterException( L"pAccessClassID", __FUNCTION__ );
    }

    // Start at v8 = Access 97 and scan forward. This means on a machine
    // with more than one version will use the oldest version
    for ( i = 8; i <= MAX_ACCESS_CURVER; i++ )
    {
        // CLSIDFromProgID wants a wide string
        memset( wzCurVer, 0, sizeof ( wzCurVer ) );
        StringCchPrintf( wzCurVer, ARRAYSIZE( wzCurVer ), L"Access.Application.%d", i );

        memset( pAccessClassID, 0, sizeof ( CLSID ) );
        hResult = CLSIDFromProgID( wzProgID, pAccessClassID );
        if ( SUCCEEDED( hResult ) )
        {
            break;
        }
    }

    // Access may not be installed
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CLSIDFromProgID", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Get the description of the current OLE DB error
//
//  Parameters:
//      pOleDbErrorString - string object to receive the error
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::GetOleDbErrorString( String* pOleDbErrorString )
{
    LCID    lcid = 0;
    ULONG   i = 0, recordCount = 0;
    HRESULT hResult = 0;
    IErrorInfo*    pIErrorInfo    = nullptr;
    IErrorRecords* pIErrorRecords = nullptr;
    AutoIUnknownRelease ReleaseIErrorInfo;
    AutoIUnknownRelease ReleaseIErrorRecords;

    if ( pOleDbErrorString == nullptr )
    {
        throw ParameterException( L"pOleDbErrorString", __FUNCTION__ );
    }
    *pOleDbErrorString = PXS_STRING_EMPTY;

    // Get the error object
    hResult = GetErrorInfo( 0, &pIErrorInfo );
    if ( FAILED( hResult ) || ( pIErrorInfo == nullptr ) )
    {
        return;
    }
    ReleaseIErrorInfo.Set( pIErrorInfo );

    // Get the records
    hResult = pIErrorInfo->QueryInterface( IID_IErrorRecords,
                                           reinterpret_cast<void**>( &pIErrorRecords ) );
    if ( FAILED( hResult ) || ( pIErrorRecords == nullptr ) )
    {
        return;
    }
    ReleaseIErrorRecords.Set( pIErrorRecords );

    hResult = pIErrorRecords->GetRecordCount( &recordCount );
    for ( i = 0; i < recordCount; i++ )
    {
        // Declare inside loop so the interfaces are auto-released
        IErrorInfo*    pIErrorRecord  = nullptr;
        ISQLErrorInfo* pISQLErrorInfo = nullptr;
        AutoIUnknownRelease ReleaseIErrorRecord;
        AutoIUnknownRelease ReleaseISQLErrorInfo;

        hResult = pIErrorRecords->GetErrorInfo( i, lcid, &pIErrorRecord );
        if ( SUCCEEDED( hResult ) )
        {
            ReleaseIErrorRecord.Set( pIErrorRecord );
        }

        hResult = pIErrorRecords->GetCustomErrorObject(
                                                   i,
                                                   IID_ISQLErrorInfo,
                                                   reinterpret_cast<IUnknown**>(&pISQLErrorInfo) );
        if ( SUCCEEDED( hResult ) )
        {
            ReleaseISQLErrorInfo.Set( pISQLErrorInfo );
        }

        MakeOleDbErrorString( i, pIErrorRecord, pISQLErrorInfo, pOleDbErrorString );
    }
}

//===============================================================================================//
//  Description:
//      Make an error description from the specified OLE DB error collection
//
//  Parameters:
//      index          - the zero based error number in the error collection
//      pIErrorInfo    - pointer to the error interface
//      pISQLErrorInfo - pointer to the SQL error interface
//      pOleDbErrorString - string object to receive the error
//
//  Remarks:
//      As this is an error diagnostic, will not allow exceptions to propagate
//      out of this method
//
//  Returns:
//      void
//===============================================================================================//
void AccessDatabase::MakeOleDbErrorString( ULONG index,
                                           IErrorInfo* pIErrorInfo,
                                           ISQLErrorInfo* pISQLErrorInfo,
                                           String* pOleDbErrorString )
{
    LONG  nativeError     = 0;
    BSTR  bstrSQLState    = nullptr;
    BSTR  bstrDescription = nullptr;
    HRESULT   hResult;
    String    Text;
    Formatter Format;

    if ( pOleDbErrorString == nullptr )
    {
        throw ParameterException( L"pOleDbErrorString", __FUNCTION__ );
    }
    *pOleDbErrorString = PXS_STRING_EMPTY;

    // Record number
    PXSGetResourceString( PXS_IDS_124_ERROR_RECORD_NUMBER, &Text );
    *pOleDbErrorString += Text;
    *pOleDbErrorString += Format.UInt32( index + 1 );   // Want one-based
    *pOleDbErrorString += PXS_CHAR_COLON;
    *pOleDbErrorString += PXS_STRING_CRLF;

    // Description
    if ( pIErrorInfo )
    {
        PXSGetResourceString( PXS_IDS_121_DESCRIPTION, &Text );
        *pOleDbErrorString += Text;
        *pOleDbErrorString += PXS_CHAR_COLON;
        *pOleDbErrorString += PXS_CHAR_SPACE;
        hResult = pIErrorInfo->GetDescription( &bstrDescription );
        if ( SUCCEEDED( hResult ) )
        {
           AutoSysFreeString SysFreeDescription( bstrDescription );
           *pOleDbErrorString += bstrDescription;
        }
        *pOleDbErrorString += PXS_STRING_CRLF;
    }

    // Examine ISQLErrorInfo
    if ( pISQLErrorInfo )
    {
        hResult = pISQLErrorInfo->GetSQLInfo( &bstrSQLState, &nativeError );
        if ( SUCCEEDED( hResult ) )
        {
            AutoSysFreeString SysFreeSQLState( bstrSQLState );

            PXSGetResourceString( PXS_IDS_122_SQL_STATE, &Text );
            *pOleDbErrorString += Text;
            *pOleDbErrorString += PXS_CHAR_COLON;
            *pOleDbErrorString += PXS_CHAR_SPACE;
            *pOleDbErrorString += bstrSQLState;
            *pOleDbErrorString += PXS_STRING_CRLF;

            PXSGetResourceString( PXS_IDS_123_NATIVE_ERROR, &Text );
            *pOleDbErrorString += Text;
            *pOleDbErrorString += PXS_CHAR_COLON;
            *pOleDbErrorString += PXS_CHAR_SPACE;
            *pOleDbErrorString += Format.Int32( nativeError );
            *pOleDbErrorString += PXS_STRING_CRLF;
        }
    }
}
