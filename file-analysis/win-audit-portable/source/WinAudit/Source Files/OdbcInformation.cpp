///////////////////////////////////////////////////////////////////////////////////////////////////
//
// ODBC Information Class Implementation
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
#include "WinAudit/Header Files/OdbcInformation.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Directory.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/FileVersion.h"
#include "PxsBase/Header Files/NameValue.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/Registry.h"
#include "PxsBase/Header Files/StringArray.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/Odbc.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
OdbcInformation::OdbcInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
OdbcInformation::~OdbcInformation()
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
//      Get the data source names as an array of audit records
//
//  Parameters:
//      pRecords - array to receive the data
//
//  Remarks:
//
//  Returns:
//      void
//===============================================================================================//
void OdbcInformation::GetDataSourceRecords( TArray< AuditRecord >* pRecords )
{
    Odbc    ODBC;
    size_t  i = 0, passNumber = 0, numElements = 0;
    String  Name, Value, DSNType;
    StringArray FileDataSourceNames;
    AuditRecord Record;
    TArray< NameValue > NameValues;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();


    ////////////////////////////////////////////////////////////////////////////
    // File Data Sources

    ODBC.GetFileDataSourceNames( &FileDataSourceNames );
    numElements = FileDataSourceNames.GetSize();
    for ( i = 0; i < numElements; i++ )
    {
        Name = FileDataSourceNames.Get( i );
        Name.Trim();
        if ( Name.GetLength() )
        {
            // File DSNs have no description
            Record.Reset( PXS_CATEGORY_ODBC_DATA_SOURCES );
            Record.Add( PXS_ODBC_DATA_SOURCES_DSN_TYPE, L"File" );
            Record.Add( PXS_ODBC_DATA_SOURCES_DSN     , Name );
            Record.Add( PXS_ODBC_DATA_SOURCES_DESC    , PXS_STRING_EMPTY );
            pRecords->Add( Record );
        }
    }

    ////////////////////////////////////////////////////////////////////////////
    // User and System Data Sources

    // Do two passes, the first gets the system data source names, the
    // second gets the user data source names
    for ( passNumber = 0; passNumber < 2; passNumber++ )
    {
        NameValues.RemoveAll();
        if ( passNumber == 0 )
        {
            DSNType = L"System";
            ODBC.GetSystemOrUserDSNs( true, &NameValues );
        }
        else
        {
            DSNType = L"User";
            ODBC.GetSystemOrUserDSNs( false, &NameValues );
        }

        numElements = NameValues.GetSize();
        for ( i = 0; i < numElements; i++ )
        {
            Name = NameValues.Get( i ).GetName();
            Name.Trim();
            if ( Name.GetLength() )
            {
                Value = NameValues.Get( i ).GetValue();
                Record.Reset( PXS_CATEGORY_ODBC_DATA_SOURCES );
                Record.Add( PXS_ODBC_DATA_SOURCES_DSN_TYPE, DSNType );
                Record.Add( PXS_ODBC_DATA_SOURCES_DSN     , Name    );
                Record.Add( PXS_ODBC_DATA_SOURCES_DESC    , Value   );
                pRecords->Add( Record );
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the installed drivers as an array of audit records
//
//  Parameters:
//      pRecords - string array to receive the data
//
//  Remarks:
//      For some reason SQLDrivers does not return all documented keywords
//      even though they are present in the appropriate registry key. In
//      particular, the Driver is absent. This is the dll file, so will
//      get the description and read the registry.
//
//  Returns:
//      void
//===============================================================================================//
void OdbcInformation::GetDriversRecords( TArray< AuditRecord >* pRecords )
{
    size_t  i = 0, numDescriptions = 0;
    Odbc    ODBC;
    File    DriverFile;
    String  RegKey, Driver, DriverDescription, DriverODBCVer;
    String  ManufacturerName, VersionString, Description;
    String  Drive, Dir, Fname, Ext;
    Registry    RegObject;
    Directory   DriverPath;
    FileVersion FileVer;
    AuditRecord Record;
    StringArray DriverDescriptions;

    // Registry path to where the drivers data is help
    LPCWSTR      STR_REG_DRIVER_KEY = L"SOFTWARE\\ODBC\\ODBCINST.INI\\";

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();
    ODBC.GetInstalledDrivers( &DriverDescriptions );

    // Connect to the registry and read properties of each driver
    RegObject.Connect( HKEY_LOCAL_MACHINE );
    numDescriptions = DriverDescriptions.GetSize();
    for ( i = 0; i < numDescriptions; i++ )
    {
        // The description is the registry key
        DriverDescription = DriverDescriptions.Get( i );
        DriverDescription.Trim();
        if ( DriverDescription.GetLength() )
        {
            // Complete the key's path
            RegKey  = STR_REG_DRIVER_KEY;
            RegKey += DriverDescription;

            // Driver file
            Driver = PXS_STRING_EMPTY;
            RegObject.GetStringValue( RegKey.c_str(), L"Driver", &Driver );
            Driver.Trim();

            // Get the ODBC version
            DriverODBCVer = PXS_STRING_EMPTY;
            RegObject.GetStringValue( RegKey.c_str(),
                                      L"DriverODBCVer", &DriverODBCVer);
            DriverODBCVer.Trim();

            // Get the version information
            ManufacturerName = PXS_STRING_EMPTY;
            VersionString    = PXS_STRING_EMPTY;
            Description      = PXS_STRING_EMPTY;
            try
            {
                if ( DriverFile.Exists( Driver ) &&
                     Driver.EndsWithStringI( L".dll" ) )
                {
                    FileVer.GetVersion( Driver, &ManufacturerName, &VersionString, &Description );
                }
            }
            catch ( const Exception& eVersion )
            {
                // Log and continue
                PXSLogException( L"Error getting version information.", eVersion, __FUNCTION__ );
            }

            // Split the files path
            Drive = PXS_STRING_EMPTY;
            Dir   = PXS_STRING_EMPTY;
            Fname = PXS_STRING_EMPTY;
            Ext   = PXS_STRING_EMPTY;
            if ( Driver.IndexOf( PXS_PATH_SEPARATOR, 0 ) != PXS_MINUS_ONE )
            {
                DriverPath.SplitPath( Driver, &Drive, &Dir, &Fname, &Ext );
                Fname += Ext;
            }
            else
            {
                // Use the entire driver
                Fname = Driver;
            }

            // Make the audit record
            Record.Reset( PXS_CATEGORY_ODBC_DRIVERS );
            Record.Add( PXS_ODBC_DRIVERS_NAME          ,  DriverDescription );
            Record.Add( PXS_ODBC_DRIVERS_COMPANY       ,  ManufacturerName );
            Record.Add( PXS_ODBC_DRIVERS_FILE_NAME     ,  Fname );
            Record.Add( PXS_ODBC_DRIVERS_FILE_VERSION  ,  VersionString );
            Record.Add( PXS_ODBC_DRIVERS_DRIVER_ODBC_VER, DriverODBCVer );
            pRecords->Add( Record );
        }
    }
    PXSSortAuditRecords( pRecords, PXS_ODBC_DRIVERS_NAME );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
