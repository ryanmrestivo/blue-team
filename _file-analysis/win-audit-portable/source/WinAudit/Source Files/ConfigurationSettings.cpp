///////////////////////////////////////////////////////////////////////////////////////////////////
//
// WinAudit Configuration Settings Class Implementation
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
#include "WinAudit/Header Files/ConfigurationSettings.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/TArray.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
ConfigurationSettings::ConfigurationSettings()
                      :autoStartAudit( true ),
                       systemOverview( true ),
                       installedSoftware( true ),
                       operatingSystem( true ),
                       peripherals( true ),
                       security( true ),
                       groupsAndUsers( true ),
                       scheduledTasks( true ),
                       uptimeStatistics( true ),
                       errorLogs( true ),
                       environmentVars( true ),
                       regionalSettings( true ),
                       windowsNetwork( true ),
                       networkTcpIp( true ),
                       hardwareDevices( true ),
                       displays( true ),
                       displayAdapters( true ),
                       installedPrinters( true ),
                       biosVersion( true ),
                       systemManagement( true ),
                       processors( true ),
                       memory( true ),
                       physicalDisks( true ),
                       drives( true ),
                       commPorts( true ),
                       startupPrograms( true ),
                       services( true ),
                       runningPrograms( true ),
                       odbcInformation( true ),
                       oleDbProviders( true ),
                       softwareMetering( false ),
                       userLogonStats( false ),
                       reportShowSQL( false ),
                       maxErrorRate( PXS_DB_MAX_ERROR_RATE_DEFAULT ),
                       maxAffectedRows( PXS_DB_MAX_AFFECTED_ROWS_DEFAULT ),
                       connectTimeoutSecs( PXS_DB_LOGIN_TIMEOUT_SECS_DEF ),
                       queryTimeoutSecs( PXS_DB_QUERY_TIMEOUT_SECS_DEF ),
                       reportMaxRecords( PXS_REPORT_MAX_RECORDS_DEFAULT ),
                       DBMS(),
                       DatabaseName(),
                       MySqlDriver(),
                       PostgreSqlDriver(),
                       ServerName(),
                       UID(),
                       LastReportName()
{
}

// Copy constructor
ConfigurationSettings::ConfigurationSettings( const ConfigurationSettings& oSettings )
{
    ConfigurationSettings();
    *this = oSettings;
}

// Destructor
ConfigurationSettings::~ConfigurationSettings()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
ConfigurationSettings& ConfigurationSettings::operator= (
                                       const ConfigurationSettings& oSettings )
{
    if ( this == &oSettings ) return *this;

    autoStartAudit     = oSettings.autoStartAudit;
    systemOverview     = oSettings.systemOverview;
    installedSoftware  = oSettings.installedSoftware;
    operatingSystem    = oSettings.operatingSystem;
    peripherals        = oSettings.peripherals;
    security           = oSettings.security;
    groupsAndUsers     = oSettings.groupsAndUsers;
    scheduledTasks     = oSettings.scheduledTasks;
    uptimeStatistics   = oSettings.uptimeStatistics;
    errorLogs          = oSettings.errorLogs;
    environmentVars    = oSettings.environmentVars;
    regionalSettings   = oSettings.regionalSettings;
    windowsNetwork     = oSettings.windowsNetwork;
    networkTcpIp       = oSettings.networkTcpIp;
    hardwareDevices    = oSettings.hardwareDevices;
    displays           = oSettings.displays;
    displayAdapters    = oSettings.displayAdapters;
    installedPrinters  = oSettings.installedPrinters;
    biosVersion        = oSettings.biosVersion;
    systemManagement   = oSettings.systemManagement;
    processors         = oSettings.processors;
    memory             = oSettings.memory;
    physicalDisks      = oSettings.physicalDisks;
    drives             = oSettings.drives;
    commPorts          = oSettings.commPorts;
    startupPrograms    = oSettings.startupPrograms;
    services           = oSettings.services;
    runningPrograms    = oSettings.runningPrograms;
    odbcInformation    = oSettings.odbcInformation;
    oleDbProviders     = oSettings.oleDbProviders;
    softwareMetering   = oSettings.softwareMetering;
    userLogonStats     = oSettings.userLogonStats;
    reportShowSQL      = oSettings.reportShowSQL;
    maxErrorRate       = oSettings.maxErrorRate;
    maxAffectedRows    = oSettings.maxAffectedRows;
    connectTimeoutSecs = oSettings.connectTimeoutSecs;
    queryTimeoutSecs   = oSettings.queryTimeoutSecs;
    reportMaxRecords   = oSettings.reportMaxRecords;
    DBMS               = oSettings.DBMS;
    DatabaseName       = oSettings.DatabaseName;
    MySqlDriver        = oSettings.MySqlDriver;
    PostgreSqlDriver   = oSettings.PostgreSqlDriver;
    ServerName         = oSettings.ServerName;
    UID                = oSettings.UID;
    LastReportName     = oSettings.LastReportName;

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//     Fill an array of the data categories to be included in the report
//
//  Parameters:
//      pCategories - receives the category identifiers
//
//  Returns:
//      void
//===============================================================================================//
void ConfigurationSettings::MakeDataCategoriesArray( TArray< DWORD >* pCategories )
{
    if ( pCategories == nullptr )
    {
        throw ParameterException( L"pCategories", __FUNCTION__ );
    }
    pCategories->RemoveAll();

    if ( systemOverview )
    {
        pCategories->Add( PXS_CATEGORY_SYSTEM_OVERVIEW );
    }

    if ( installedSoftware )
    {
        pCategories->Add( PXS_CATEGORY_INSTALLED_SOFTWARE );
        pCategories->Add( PXS_CATEGORY_ACTIVE_SETUP );
        pCategories->Add( PXS_CATEGORY_INSTALLED_PROGS );
        pCategories->Add( PXS_CATEGORY_SOFTWARE_UPDATES );
    }

    if ( operatingSystem )
    {
        pCategories->Add( PXS_CATEGORY_OS );
    }

    if ( peripherals )
    {
        pCategories->Add( PXS_CATEGORY_PERIPHERALS );
    }

    if ( security )
    {
        pCategories->Add( PXS_CATEGORY_SECURITY );
        pCategories->Add( PXS_CATEGORY_KERBEROS_POLICY );
        pCategories->Add( PXS_CATEGORY_KERBEROS_TICKETS );
        pCategories->Add( PXS_CATEGORY_NET_TIME_PROTOCOL );
        pCategories->Add( PXS_CATEGORY_PERMISSIONS );
        pCategories->Add( PXS_CATEGORY_REG_SEC_VALUES );
        pCategories->Add( PXS_CATEGORY_SECURITY_LOG );
        pCategories->Add( PXS_CATEGORY_SECUR_SETTINGS );
        pCategories->Add( PXS_CATEGORY_SYSRESTORE );
        pCategories->Add( PXS_CATEGORY_USERPRIVS );
        pCategories->Add( PXS_CATEGORY_USER_RIGHTS );
        pCategories->Add( PXS_CATEGORY_WINDOWS_FIREWALL );
    }

    if ( groupsAndUsers )
    {
        pCategories->Add( PXS_CATEGORY_GROUPS_AND_USERS );
        pCategories->Add( PXS_CATEGORY_GROUPS );
        pCategories->Add( PXS_CATEGORY_GROUPMEMBERS );
        pCategories->Add( PXS_CATEGORY_GROUPPOLICY );
        pCategories->Add( PXS_CATEGORY_USERS );
    }

    if ( scheduledTasks )
    {
        pCategories->Add( PXS_CATEGORY_SCHED_TASKS );
    }

    if ( uptimeStatistics )
    {
        pCategories->Add( PXS_CATEGORY_UPTIME );
    }

    if ( errorLogs )
    {
        pCategories->Add( PXS_CATEGORY_ERROR_LOG );
    }

    if ( environmentVars )
    {
        pCategories->Add( PXS_CATEGORY_ENVIRON_VARS );
    }

    if ( regionalSettings )
    {
        pCategories->Add( PXS_CATEGORY_REGION_SETTINGS );
    }

    if ( windowsNetwork )
    {
        pCategories->Add( PXS_CATEGORY_WINDOWS_NETWORK );
        pCategories->Add( PXS_CATEGORY_WIN_NET_FILES );
        pCategories->Add( PXS_CATEGORY_WIN_NET_SESSIONS );
        pCategories->Add( PXS_CATEGORY_WIN_NET_SHARES );
    }

    if ( networkTcpIp )
    {
        pCategories->Add( PXS_CATEGORY_NETWORK_TCPIP );
        pCategories->Add( PXS_CATEGORY_NET_ADAPTERS );
        pCategories->Add( PXS_CATEGORY_OPEN_PORTS );
        pCategories->Add( PXS_CATEGORY_ROUTING_TABLE );
    }

    if ( hardwareDevices )
    {
        pCategories->Add( PXS_CATEGORY_HARDWARE_DEVICES );
    }

    if ( displays )
    {
        pCategories->Add( PXS_CATEGORY_DISPLAY_EDID );
    }

    if ( displayAdapters )
    {
        pCategories->Add( PXS_CATEGORY_DISPLAY_ADAPTERS );
    }

    if ( installedPrinters )
    {
        pCategories->Add( PXS_CATEGORY_PRINTERS );
    }

    if ( biosVersion )
    {
        pCategories->Add( PXS_CATEGORY_BIOS_VERSION );
    }

    if ( systemManagement )
    {
        pCategories->Add( PXS_CATEGORY_SMBIOS );
        pCategories->Add( PXS_CATEGORY_SMBIOS_INFO );
        pCategories->Add( PXS_CATEGORY_SMBIOS_SYSINFO );
        pCategories->Add( PXS_CATEGORY_SMBIOS_BOARD );
        pCategories->Add( PXS_CATEGORY_SMBIOS_CHASSIS );
        pCategories->Add( PXS_CATEGORY_SMBIOS_PROC );
        pCategories->Add( PXS_CATEGORY_SMBIOS_MEMCTRL );
        pCategories->Add( PXS_CATEGORY_SMBIOS_MEMMODULE );
        pCategories->Add( PXS_CATEGORY_SMBIOS_CPUCACHE );
        pCategories->Add( PXS_CATEGORY_SMBIOS_PORTCONN );
        pCategories->Add( PXS_CATEGORY_SMBIOS_SYSSLOT );
        pCategories->Add( PXS_CATEGORY_SMBIOS_MEMARRAY );
        pCategories->Add( PXS_CATEGORY_SMBIOS_MEMDEV );
    }

    if ( processors )
    {
        pCategories->Add( PXS_CATEGORY_CPU_BASIC );
    }

    if ( memory )
    {
        pCategories->Add( PXS_CATEGORY_MEMORY );
    }

    if ( physicalDisks )
    {
        pCategories->Add( PXS_CATEGORY_PHYS_DISKS );
    }

    if ( drives )
    {
        pCategories->Add( PXS_CATEGORY_DRIVES );
    }

    if ( commPorts )
    {
        pCategories->Add( PXS_CATEGORY_COMMPORTS );
    }

    if ( startupPrograms )
    {
        pCategories->Add( PXS_CATEGORY_STARTUP );
    }

    if ( services )
    {
        pCategories->Add( PXS_CATEGORY_NTSERVICES );
    }

    if ( runningPrograms )
    {
        pCategories->Add( PXS_CATEGORY_RUNNING_PROCS );
    }

    if ( odbcInformation )
    {
        pCategories->Add( PXS_CATEGORY_ODBC_INFORMATION );
        pCategories->Add( PXS_CATEGORY_ODBC_DATA_SOURCES );
        pCategories->Add( PXS_CATEGORY_ODBC_DRIVERS );
    }

    if ( oleDbProviders )
    {
        pCategories->Add( PXS_CATEGORY_OLE_DB_PROVIDERS );
    }

    if ( softwareMetering )
    {
        pCategories->Add( PXS_CATEGORY_SOFTWARE_METERING );
    }

    if ( userLogonStats )
    {
        pCategories->Add( PXS_CATEGORY_USER_LOGONS );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
