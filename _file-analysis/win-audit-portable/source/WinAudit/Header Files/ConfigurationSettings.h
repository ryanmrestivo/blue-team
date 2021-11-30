///////////////////////////////////////////////////////////////////////////////////////////////////
//
// WinAudit Configuration Settings Class Header
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

#ifndef WINAUDIT_CONFIGURATION_SETTINGS_H_
#define WINAUDIT_CONFIGURATION_SETTINGS_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/WinAudit.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/StringT.h"

// 5. This Project

// 6. Forwards
template< class T > class TArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class ConfigurationSettings
{
    public:
        // Default constructor
        ConfigurationSettings();

        // Destructor
        ~ConfigurationSettings();

        // Copy constructor
        ConfigurationSettings( const ConfigurationSettings& oSettings );

        // Assignment operator
        ConfigurationSettings& operator= ( const ConfigurationSettings& oSettings );

        // Methods
        void MakeDataCategoriesArray( TArray< DWORD >* pCategories );

        // Data members
        bool    autoStartAudit;         // Automatically start the audit
        bool    systemOverview;
        bool    installedSoftware;
        bool    operatingSystem;
        bool    peripherals;
        bool    security;
        bool    groupsAndUsers;
        bool    scheduledTasks;
        bool    uptimeStatistics;
        bool    errorLogs;
        bool    environmentVars;
        bool    regionalSettings;
        bool    windowsNetwork;
        bool    networkTcpIp;
        bool    hardwareDevices;
        bool    displays;
        bool    displayAdapters;
        bool    installedPrinters;
        bool    biosVersion;
        bool    systemManagement;
        bool    processors;
        bool    memory;
        bool    physicalDisks;
        bool    drives;
        bool    commPorts;
        bool    startupPrograms;
        bool    services;
        bool    runningPrograms;
        bool    odbcInformation;
        bool    oleDbProviders;
        bool    softwareMetering;
        bool    userLogonStats;
        bool    reportShowSQL;        // Show the SQL statement in the report
        DWORD   maxErrorRate;         // The maximum error rate for a export
        DWORD   maxAffectedRows;      // The maximum allowed affected rows
        DWORD   connectTimeoutSecs;   // The database connection timeout
        DWORD   queryTimeoutSecs;     // The statement query timeout
        DWORD   reportMaxRecords;     // The maximum records to show in a report
        String  DBMS;                 // The database management system name
        String  DatabaseName;         // The database name
        String  MySqlDriver;          // The MySQL driver name
        String  PostgreSqlDriver;     // The PostgreSQL driver name
        String  ServerName;           // The database server name
        String  UID;                  // The user's id
        String  LastReportName;       // The last selected database report

    protected:
        // Methods

        // Data members

    private:
        // Methods

        // Data members
};

#endif  // WINAUDIT_CONFIGURATION_SETTINGS_H_
