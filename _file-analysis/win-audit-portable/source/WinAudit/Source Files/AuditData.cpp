///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Audit Data Class Implementation
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
#include "WinAudit/Header Files/AuditData.h"

// 2. C System Files
#include <msdaguid.h>
#include <oledb.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/AllocateWChars.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/AutoIUnknownRelease.h"
#include "PxsBase/Header Files/ComException.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Directory.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/FileVersion.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/Mutex.h"
#include "PxsBase/Header Files/NameValue.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/Registry.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/Wmi.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/CommunicationPortInformation.h"
#include "WinAudit/Header Files/CpuInformation.h"
#include "WinAudit/Header Files/DeviceInformation.h"
#include "WinAudit/Header Files/DisplayInformation.h"
#include "WinAudit/Header Files/DriveInformation.h"
#include "WinAudit/Header Files/DiskInformation.h"
#include "WinAudit/Header Files/EventLogInformation.h"
#include "WinAudit/Header Files/GroupUserInformation.h"
#include "WinAudit/Header Files/KerberosTicketInformation.h"
#include "WinAudit/Header Files/NtServiceInformation.h"
#include "WinAudit/Header Files/ObjectPermissionInformation.h"
#include "WinAudit/Header Files/OdbcInformation.h"
#include "WinAudit/Header Files/OpenNetworkPortInformation.h"
#include "WinAudit/Header Files/PeripheralInformation.h"
#include "WinAudit/Header Files/PrinterInformation.h"
#include "WinAudit/Header Files/ProcessInformation.h"
#include "WinAudit/Header Files/SecurityInformation.h"
#include "WinAudit/Header Files/SoftwareInformation.h"
#include "WinAudit/Header Files/TaskSchedulerInformation.h"
#include "WinAudit/Header Files/TcpIpInformation.h"
#include "WinAudit/Header Files/WindowsFirewallInformation.h"
#include "WinAudit/Header Files/WindowsInformation.h"
#include "WinAudit/Header Files/WindowsNetworkInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
AuditData::AuditData()
          :m_SmbiosInfo()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
AuditData::~AuditData()
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
//      Get BIOS identification data
//
//  Parameters:
//      pRecord - receives the data
//
//  Returns:
//      void
//===============================================================================================//
void AuditData::GetBIOSIdentificationRecord( AuditRecord* pRecord )
{
    String   Value;
    Registry RegObject;
    LPCWSTR  STR_SUB_KEY = L"Hardware\\Description\\System";

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_BIOS_VERSION );

    RegObject.Connect( HKEY_LOCAL_MACHINE );

    Value = PXS_STRING_EMPTY;
    RegObject.GetStringValue( STR_SUB_KEY, L"SystemBiosVersion", &Value  );
    pRecord->Add( PXS_BIOS_VERSION_BIOS_VERSION, Value );

    Value = PXS_STRING_EMPTY;
    RegObject.GetStringValue( STR_SUB_KEY, L"SystemBiosDate", &Value );
    pRecord->Add( PXS_BIOS_VERSION_RELEASE_DATE, Value );

    Value = PXS_STRING_EMPTY;
    RegObject.GetStringValue( STR_SUB_KEY, L"VideoBiosVersion", &Value );
    pRecord->Add( PXS_BIOS_VERSION_VIDEO_BIOS_VER, Value );

    Value = PXS_STRING_EMPTY;
    RegObject.GetStringValue( STR_SUB_KEY, L"VideoBiosDate", &Value );
    pRecord->Add( PXS_BIOS_VERSION_VIDEO_BIOS_DATE, Value );
}

//===============================================================================================//
//  Description:
//      Get the category name in English from the specified category id
//
//  Parameters:
//      categoryID    - defined category number
//      pCategoryName - receives the category name
//
//  Returns:
//      void, "" if not found
//===============================================================================================//
void AuditData::GetCategoryName( DWORD categoryID, String* pCategoryName )
{
    size_t i = 0;

    if ( pCategoryName == nullptr )
    {
        throw ParameterException( L"pCategoryName", __FUNCTION__ );
    }
    *pCategoryName = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( PXS_DATA_CATEGORY_PROPERTIES ); i++ )
    {
        if ( categoryID == PXS_DATA_CATEGORY_PROPERTIES[ i ].categoryID )
        {
            PXSGetResourceString( PXS_DATA_CATEGORY_PROPERTIES[ i ].nameStringID, pCategoryName );
            break;
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the audit records for the specified data category
//
//  Parameters:
//      categoryID - defined category number
//      LocalTime  - the computer local time
//      pRecords   - array to receive the records
//
//  Returns:
//      void
//===============================================================================================//
void AuditData::GetCategoryRecords( DWORD categoryID,
                                    const String& LocalTime, TArray< AuditRecord >* pRecords )
{
    const size_t MAX_EVENT_LOG_ENTRIES = 25;
    bool        isColumnar = false, isNode = false;
    BYTE        indent     = 0;
    DWORD       captionID  = 0;
    size_t      numRecords = 0;
    String      NumberOfRecords, CategoryName;
    Formatter   Format;
    AuditRecord Record;
    TArray< AuditRecord >        SubCategoryRecords;
    CommunicationPortInformation CommunicationPortInfo;
    CpuInformation               CpuInfo;
    DeviceInformation            DeviceInfo;
    DiskInformation              DiskInfo;
    DisplayInformation           DisplayInfo;
    DriveInformation             DriveInfo;
    EventLogInformation          EventLogInfo;
    GroupUserInformation         GroupUserInfo;
    KerberosTicketInformation    KerberosTicketInfo;
    NtServiceInformation         NtServiceInfo;
    OdbcInformation              OdbcInfo;
    ObjectPermissionInformation  ObjectPermissionInfo;
    OpenNetworkPortInformation   OpenNetworkPortInfo;
    PeripheralInformation        PeripheralInfo;
    PrinterInformation           PrinterInfo;
    ProcessInformation           ProcessInfo;
    SecurityInformation          SecurityInfo;
    SoftwareInformation          SoftwareInfo;
    TaskSchedulerInformation     TaskSchedulerInfo;
    TcpIpInformation             TcpIpInfo;
    WindowsFirewallInformation   WindowsFirewallInfo;
    WindowsNetworkInformation    WindowsNetworkInfo;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Get the category name
    PXSGetDataCategoryProperties( categoryID,
                                  &CategoryName, &captionID, &isColumnar, &isNode, &indent );
    PXSLogAppInfo1( L"Start category '%%1'.", CategoryName );

    switch ( categoryID )
    {
        default:
            throw ParameterException( L"categoryID", __FUNCTION__ );

        ///////////////////////////////////////////////////////////////////////
        // Groupings for display purposes

        // Fall through. Nothing to do but will add an empty record to the
        // output to signify this category has been done
        case PXS_CATEGORY_INSTALLED_SOFTWARE:
        case PXS_CATEGORY_SECURITY:
        case PXS_CATEGORY_GROUPS_AND_USERS:
        case PXS_CATEGORY_WINDOWS_NETWORK:
        case PXS_CATEGORY_NETWORK_TCPIP:
        case PXS_CATEGORY_SMBIOS:
        case PXS_CATEGORY_ODBC_INFORMATION:
            break;

        ///////////////////////////////////////////////////////////////////////
        // Data Categories

        case PXS_CATEGORY_SYSTEM_OVERVIEW:
            GetSystemOverviewRecord( LocalTime, &Record );
            pRecords->Add( Record );
            break;

        case PXS_CATEGORY_ACTIVE_SETUP:
            SoftwareInfo.GetActiveSetupRecords( pRecords );
            break;

        case PXS_CATEGORY_INSTALLED_PROGS:
            SoftwareInfo.GetInstalledSoftwareRecords( pRecords );
            break;

        case PXS_CATEGORY_SOFTWARE_UPDATES:
            SoftwareInfo.GetSoftwareUpdateRecords( pRecords );
            break;

        case PXS_CATEGORY_OS:
            GetOperatingSystemRecord( &Record );
            pRecords->Add( Record );
            break;

        case PXS_CATEGORY_PERIPHERALS:
            PeripheralInfo.GetAuditRecords( pRecords );
            break;

        case PXS_CATEGORY_INET_SOFTWARE:
            PXSLogAppError1( L"Category '%%1' is no longer used.",
                             CategoryName );
            break;

        case PXS_CATEGORY_OPEN_PORTS:
            OpenNetworkPortInfo.GetAuditRecords( pRecords );
            break;

        case PXS_CATEGORY_SECURITY_LOG:
            EventLogInfo.GetAuditFailureRecords( MAX_EVENT_LOG_ENTRIES, pRecords );
            break;

        case PXS_CATEGORY_SECUR_SETTINGS:
            SecurityInfo.GetSecuritySettingsRecords( pRecords );
            break;

        case PXS_CATEGORY_SYSRESTORE:
            SecurityInfo.GetSystemRestorePointRecords( pRecords );
            break;

        case PXS_CATEGORY_USERPRIVS:
            SecurityInfo.GetLoggedOnUserPrivilegePrecords( pRecords );
            break;

        case PXS_CATEGORY_WINDOWS_FIREWALL:
            WindowsFirewallInfo.GetWindowsFirewallRecords( pRecords );
            break;

        case PXS_CATEGORY_GROUPS:
            GroupUserInfo.GetGroupRecords( pRecords );
            break;

        case PXS_CATEGORY_GROUPMEMBERS:
            GroupUserInfo.GetGroupMemberRecords( pRecords );
            break;

        case PXS_CATEGORY_GROUPPOLICY:
            GroupUserInfo.GetGroupPolicyRecords( pRecords );
            break;

        case PXS_CATEGORY_USERS:
            GroupUserInfo.GetUserRecords( pRecords );
            break;

        case PXS_CATEGORY_SCHED_TASKS:
            TaskSchedulerInfo.GetScheduledTaskRecords( pRecords );
            break;

        case PXS_CATEGORY_UPTIME:
            EventLogInfo.GetUptimeStatsRecord( &Record );
            pRecords->Add( Record );
            break;

        case PXS_CATEGORY_ERROR_LOG:
            EventLogInfo.GetErrorRecords( MAX_EVENT_LOG_ENTRIES, pRecords );
            break;

        case PXS_CATEGORY_WIN_NET_FILES:
            WindowsNetworkInfo.GetOpenFileRecords( pRecords );
            break;

        case PXS_CATEGORY_WIN_NET_SESSIONS:
            WindowsNetworkInfo.GetSessionRecords( pRecords );
            break;

        case PXS_CATEGORY_WIN_NET_SHARES:
            WindowsNetworkInfo.GetShareRecords( pRecords );
            break;

        case PXS_CATEGORY_NET_ADAPTERS:
            TcpIpInfo.GetAdaptersRecords( pRecords );
            break;

        case PXS_CATEGORY_NETBIOS:
            PXSLogAppError1( L"Category '%%1' is not used.", CategoryName );
            break;

        case PXS_CATEGORY_HARDWARE_DEVICES:
            DeviceInfo.GetAuditRecords( pRecords );
            break;

        case PXS_CATEGORY_DISPLAY_CAPS:
            PXSLogAppError1( L"Category '%%1' is not used.", CategoryName );
            break;

        case PXS_CATEGORY_PRINTERS:
            PrinterInfo.GetAuditRecords( pRecords );
            break;

        case PXS_CATEGORY_BIOS_VERSION:
            GetBIOSIdentificationRecord( &Record );
            pRecords->Add( Record );
            break;

        case PXS_CATEGORY_CPU_BASIC:
            CpuInfo.GetAuditRecords( pRecords );
            break;

        case PXS_CATEGORY_CPU_DETAILED:
            PXSLogAppError1( L"Category '%%1' is not used.", CategoryName );
            break;

        case PXS_CATEGORY_CPU_CACHE:
            PXSLogAppError1( L"Category '%%1' is not used.", CategoryName );
            break;

        case PXS_CATEGORY_CPU_TLB:
            PXSLogAppError1( L"Category '%%1' is not used.", CategoryName );
            break;

        case PXS_CATEGORY_MEMORY:
            GetMemoryInformationRecord( &Record );
            pRecords->Add( Record );
            break;

        case PXS_CATEGORY_PHYS_DISKS:
            DiskInfo.GetAuditRecords( pRecords );
            break;

        case PXS_CATEGORY_DRIVES:
            DriveInfo.GetAuditRecords( pRecords );
            break;

        case PXS_CATEGORY_COMMPORTS:
            CommunicationPortInfo.GetAuditRecords( pRecords );
            break;

        case PXS_CATEGORY_STARTUP:
            SoftwareInfo.GetStartupProgramRecords( pRecords );
            break;

        case PXS_CATEGORY_NTSERVICES:
            NtServiceInfo.GetServiceRecords( pRecords );
            break;

        case PXS_CATEGORY_RUNNING_PROCS:
            ProcessInfo.GetRunningProcesseRecords( pRecords );
            break;

        case PXS_CATEGORY_LOADED_MODULES:
            PXSLogAppError1( L"Category '%%1' is not used.", CategoryName );
            break;

        case PXS_CATEGORY_SYSTEM_FILES:
            PXSLogAppError1( L"Category '%%1' is not used.", CategoryName );
            break;

        case PXS_CATEGORY_FIND_FILES:
            PXSLogAppError1( L"Category '%%1' is not used.", CategoryName );
            break;

        case PXS_CATEGORY_DISPLAY_EDID:
            DisplayInfo.GetIdentificationRecords( pRecords );
            break;

        case PXS_CATEGORY_DISPLAY_ADAPTERS:
            DisplayInfo.GetAdapterRecords( pRecords );
            break;

        case PXS_CATEGORY_ENVIRON_VARS:
            GetEnvironmentVarsRecords( pRecords );
            break;

        case PXS_CATEGORY_REGION_SETTINGS:
            GetRegionalSettingsRecords( pRecords );
            break;

        case PXS_CATEGORY_PERMISSIONS:
            ObjectPermissionInfo.GetAuditRecords( pRecords );
            break;

        case PXS_CATEGORY_USER_LOGONS:
            EventLogInfo.GetUserLogonRecords( pRecords );
            break;

        case PXS_CATEGORY_SOFTWARE_METERING:
            EventLogInfo.GetSoftwareMeteringRecords( pRecords );
            break;

        case PXS_CATEGORY_ODBC_DRIVERS:
            OdbcInfo.GetDriversRecords( pRecords );
            break;

        case PXS_CATEGORY_ODBC_DATA_SOURCES:
            OdbcInfo.GetDataSourceRecords( pRecords );
            break;

        case PXS_CATEGORY_OTHER_AUDIT_DATA:
            // Reserved
            break;

        case PXS_CATEGORY_NON_WINDOWS_EXE:
            PXSLogAppError1( L"Category '%%1' is not used.", CategoryName );
            break;

        case PXS_CATEGORY_SMBIOS_INFO:
            m_SmbiosInfo.GetAuditRecords( PXS_SMBIOS_TYPE_0_BIOS, pRecords );
            break;

        case PXS_CATEGORY_SMBIOS_SYSINFO:
            m_SmbiosInfo.GetAuditRecords( PXS_SMBIOS_TYPE_1_SYSTEM, pRecords );
            break;

        case PXS_CATEGORY_SMBIOS_BOARD:
            m_SmbiosInfo.GetAuditRecords( PXS_SMBIOS_TYPE_2_BASE_BOARD, pRecords );
            break;

        case PXS_CATEGORY_SMBIOS_CHASSIS:
            m_SmbiosInfo.GetAuditRecords( PXS_SMBIOS_TYPE_3_CHASSIS, pRecords );
            break;

        case PXS_CATEGORY_SMBIOS_PROC:
            m_SmbiosInfo.GetAuditRecords( PXS_SMBIOS_TYPE_4_PROCESSOR, pRecords );
            break;

        case PXS_CATEGORY_SMBIOS_MEMCTRL:
            m_SmbiosInfo.GetAuditRecords( PXS_SMBIOS_TYPE_5_MEM_CONTROL, pRecords );
            break;

        case PXS_CATEGORY_SMBIOS_MEMMODULE:
            m_SmbiosInfo.GetAuditRecords( PXS_SMBIOS_TYPE_6_MEMORY_MODULE, pRecords );
            break;

        case PXS_CATEGORY_SMBIOS_CPUCACHE:
            m_SmbiosInfo.GetAuditRecords( PXS_SMBIOS_TYPE_7_CPU_CACHE, pRecords );
            break;

        case PXS_CATEGORY_SMBIOS_PORTCONN:
            m_SmbiosInfo.GetAuditRecords( PXS_SMBIOS_TYPE_8_PORT_CONN, pRecords );
            break;

        case PXS_CATEGORY_SMBIOS_SYSSLOT:
            m_SmbiosInfo.GetAuditRecords( PXS_SMBIOS_TYPE_9_SYSTEM_SLOT, pRecords );
            break;

        case PXS_CATEGORY_SMBIOS_MEMARRAY:
            m_SmbiosInfo.GetAuditRecords( PXS_SMBIOS_TYPE_16_MEM_ARRAY, pRecords );
            break;

        case PXS_CATEGORY_SMBIOS_MEMDEV:
            m_SmbiosInfo.GetAuditRecords( PXS_SMBIOS_TYPE_17_MEMORY_DEVICE, pRecords );
            break;

        case PXS_CATEGORY_NET_TIME_PROTOCOL:
            SecurityInfo.GetNetworkTimeProtocolRecords( pRecords );
            break;

        case PXS_CATEGORY_USER_RIGHTS:
            SecurityInfo.GetUserRightsAssignmentRecords( pRecords );
            break;

        case PXS_CATEGORY_KERBEROS_POLICY:
            KerberosTicketInfo.GetKerberosPolicyRecords( pRecords );
            break;

        case PXS_CATEGORY_KERBEROS_TICKETS:
            KerberosTicketInfo.GetKerberosTicketRecords( pRecords );
            break;

        case PXS_CATEGORY_REG_SEC_VALUES:
            SecurityInfo.GetRegistrySecurityValueRecods( pRecords );
            break;

        case PXS_CATEGORY_OLE_DB_PROVIDERS:
            GetOleDbProviderRecords( pRecords );
            break;

        case PXS_CATEGORY_ROUTING_TABLE:
            TcpIpInfo.GetRoutingTableRecords( pRecords );
            break;
    }

    numRecords      = pRecords->GetSize();
    NumberOfRecords = Format.SizeT( numRecords );
    PXSLogAppInfo2( L"End category '%%1', found %%2 record(s).",
                   CategoryName, NumberOfRecords );

    // Will add an empty record to indicate the category has been done
    if ( numRecords == 0 )
    {
        Record.Reset( categoryID );
        pRecords->Add( Record );
    }

    // To help locate problems will flush the log at the end of each category
    // so messages are written to disk
    if ( g_pApplication )
    {
        g_pApplication->LogFlush();
    }
}

//===============================================================================================//
//  Description:
//      Get the environment variables as an array of audit records
//
//  Parameters:
//      pRecords - array to receive the records
//
//  Returns:
//      void
//===============================================================================================//
void AuditData::GetEnvironmentVarsRecords( TArray< AuditRecord >* pRecords )
{
    const    size_t MAX_DATA_CHARS = 8192;    // maximum characters of data
    bool     atEnd = false;
    size_t   i = 0, numChars = 0, idxStart = 0;
    String   Value, Name, NameValue;
    wchar_t* pszData  = nullptr;
    wchar_t* pszBlock;
    Formatter   Format;
    StringArray Tokens;
    AuditRecord Record;
    AllocateWChars AllocWChars;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    pszBlock = GetEnvironmentStrings();     // Not thread safe
    if ( pszBlock == nullptr )
    {
        throw NullException( L"pszBlock", __FUNCTION__ );
    }

    // Need to free the variables
    try
    {
        // The block of data is double null terminated
        while ( ( ( numChars + 1 ) < MAX_DATA_CHARS ) && ( atEnd == false ) )
        {
            numChars++;
            if ( ( pszBlock[ numChars ]     == PXS_CHAR_NULL ) &&
                 ( pszBlock[ numChars + 1 ] == PXS_CHAR_NULL )  )
            {
                atEnd = true;
                numChars--;     // Back of from the NULL
            }
        }
        PXSLogAppInfo1( L"Read %%1 characters of environment variables.",
                        Format.SizeT( numChars ) );

        // Make a copy
        numChars = PXSAddSizeT( numChars, 1 );       // NULL terminator
        pszData  = AllocWChars.New( numChars );
        memcpy( pszData, pszBlock, sizeof ( wchar_t ) * numChars );
        pszData[ numChars - 1 ] = PXS_CHAR_NULL;
        for ( i = 0; i < numChars; i++ )
        {
            // Format it name=value
            if ( pszData[ i ] == PXS_CHAR_NULL )
            {
                Tokens.RemoveAll();
                NameValue = pszData + idxStart;
                NameValue.ToArray( '=', &Tokens );
                if ( Tokens.GetSize() == 2 )
                {
                    Name  = Tokens.Get( 0 );
                    Value = Tokens.Get( 1 );
                    Record.Reset( PXS_CATEGORY_ENVIRON_VARS );
                    Record.Add( PXS_ENVIRON_VARS_NAME , Name );
                    Record.Add( PXS_ENVIRON_VARS_VALUE, Value );
                    pRecords->Add( Record );
                }
                idxStart = i + 1;   // Set start position for the next item
            }
        }
    }
    catch ( const Exception& )
    {
        FreeEnvironmentStrings( pszBlock );
        throw;
    }
    FreeEnvironmentStrings( pszBlock );
    PXSSortAuditRecords( pRecords, PXS_ENVIRON_VARS_NAME);
}

//===============================================================================================//
//  Description:
//      Get memory information
//
//  Parameters:
//      pRecord - receives the memory data
//
//  Returns:
//      Total RAM in MB
//===============================================================================================//
DWORD AuditData::GetMemoryInformationRecord( AuditRecord* pRecord )
{
    const __int64 DIV_MB  = 1024*1024;
    int       pageFileMB = 0, totalPageFileMB = 0;
    Wmi       WMI;
    DWORD     totalPhysMB = 0, smBiosMB = 0;
    String    Value, LocaleMB;
    Formatter Format;
    MEMORYSTATUSEX  memStatusEx;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_MEMORY );

    PXSGetResourceString( PXS_IDS_136_MB, &LocaleMB );

    // Try to get the maximum page file size using WMI as GlobalMemoryStatusEx
    // returns the total page size available to the calling process
    try
    {
        // There may be more than one page file. Note, Win32_PageFile is
        // deprecated.Docs say its uint32 but it comes back as lVal = VT_I4
        WMI.Connect( L"root\\cimv2" );
        WMI.ExecQuery( L"Select * from Win32_PageFile" );
        while ( WMI.Next() )
        {
            pageFileMB = 0;
            if ( WMI.GetInt32( L"MaximumSize", &pageFileMB ) )
            {
                totalPageFileMB = PXSAddInt32( totalPageFileMB, pageFileMB );
            }
        }
    }
    catch ( const Exception& e )
    {
        // Log and continue
        PXSLogException( L"Error getting page size.", e, __FUNCTION__ );
    }

    memset( &memStatusEx, 0, sizeof ( memStatusEx ) );
    memStatusEx.dwLength = sizeof ( memStatusEx );
    if ( GlobalMemoryStatusEx( &memStatusEx ) == 0 )
    {
        throw SystemException( GetLastError(), L"GlobalMemoryStatusEx", __FUNCTION__ );
    }

    // Total RAM. Use SMBIOS in preference. Sometimes the SMBIOS data is
    // incorrect so check that it is more than obtained via the System API but
    // not too much more.
    m_SmbiosInfo.ReadSmbiosData();
    smBiosMB    = m_SmbiosInfo.GetTotalRamMB();
    totalPhysMB = PXSCastUInt64ToUInt32( memStatusEx.ullTotalPhys / DIV_MB );
    if ( ( smBiosMB >= totalPhysMB ) &&
         ( smBiosMB < (totalPhysMB + 256) ) )
    {
        totalPhysMB = smBiosMB;
    }
    Value  = Format.UInt32( totalPhysMB );
    Value += LocaleMB;
    pRecord->Add( PXS_MEMORY_TOTAL_RAM_MB, Value );

    // Free RAM
    Value  = Format.UInt64( memStatusEx.ullAvailPhys / DIV_MB );
    Value += LocaleMB;
    pRecord->Add( PXS_MEMORY_FREE_RAM_MB , Value );

    // Total page file. Use WMI in preference
    if ( totalPageFileMB > 0 )
    {
        Value = Format.Int32( totalPageFileMB );
    }
    else
    {
        Value = Format.UInt64( memStatusEx.ullTotalPageFile / DIV_MB );
    }
    Value += LocaleMB;
    pRecord->Add( PXS_MEMORY_MAX_SWAP_MB, Value );

    // Available page file
    Value  = Format.UInt64( memStatusEx.ullAvailPageFile / DIV_MB );
    Value += LocaleMB;
    pRecord->Add( PXS_MEMORY_FREE_SWAP_MB, Value );

    return totalPhysMB;
}

//===============================================================================================//
//  Description:
//      Get OLD BD Provider information as an array of audit records
//
//  Parameters:
//      pRecords  - receives the provider data
//
//  Returns:
//      void
//===============================================================================================//
void AuditData::GetOleDbProviderRecords( TArray< AuditRecord >* pRecords )
{
    size_t  i = 0, numProviders = 0;
    String  Name, Description;
    AuditRecord Record;
    TArray< NameValue > OleDbProviders;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    GetOleDbProviders( &OleDbProviders );
    numProviders = OleDbProviders.GetSize();
    for ( i = 0; i < numProviders; i++ )
    {
        // Add to output array
        Name = OleDbProviders.Get( i ).GetName();
        if ( Name.GetLength() )
        {
            Description = OleDbProviders.Get( i ).GetValue();
            Record.Reset( PXS_CATEGORY_OLE_DB_PROVIDERS );
            Record.Add( PXS_OLE_DB_PROVIDERS_NAME       , Name );
            Record.Add( PXS_OLE_DB_PROVIDERS_DESCRIPTION, Description );
            pRecords->Add( Record );
        }
    }
    PXSSortAuditRecords( pRecords, PXS_OLE_DB_PROVIDERS_NAME );
}

//===============================================================================================//
//  Description:
//      Get a list of OLEDB providers on the local machine
//
//  Parameters:
//      pOleDbProviders - receives the providers SOURCES_NAME and
//                        SOURCES_DESCRIPTION as name/value pairs
//  Returns:
//      void
//===============================================================================================//
void AuditData::GetOleDbProviders( TArray< NameValue >* pOleDbProviders )
{
    const DBROWCOUNT MAX_FETCH_ROWS    = 256;
    const DBLENGTH   NUM_RESULT_COLUMNS= 2;  // SOURCES_NAME+SOURCES_DESCRIPTION
    const DBLENGTH   COLUMN_LEN_CHARS  = 256;
    const DBLENGTH   COLUMN_LEN_BYTES  = COLUMN_LEN_CHARS * sizeof ( wchar_t );
    wchar_t   wzBuffer[ NUM_RESULT_COLUMNS * COLUMN_LEN_CHARS ] = { 0 };
    HROW      rghRows[ MAX_FETCH_ROWS ] = { 0 };
    HROW*     pRows = nullptr;
    String    Name, Description;
    HRESULT   hResult;
    NameValue Provider;
    Formatter Format;
    HACCESSOR hAccessor = 0;              // = ULONG_PTR
    DBBINDING rgBindings[ NUM_RESULT_COLUMNS ];
    DBCOUNTITEM   cRowsObtained = 0, i = 0;
    AllocateBytes Alloc;

    IRowset*      pIRowset   = nullptr;
    IAccessor*    pIAccessor = nullptr;
    ISourcesRowset*     pISourcesRowset = nullptr;
    AutoIUnknownRelease AutoReleaseIRowset;
    AutoIUnknownRelease AutoReleaseIAccessor;
    AutoIUnknownRelease AutoReleaseISourcesRowset;

    if ( pOleDbProviders == nullptr )
    {
        throw ParameterException( L"pOleDbProviders", __FUNCTION__ );
    }
    pOleDbProviders->RemoveAll();

    hResult = CoCreateInstance( CLSID_OLEDB_ENUMERATOR,
                                nullptr,
                                CLSCTX_INPROC_SERVER,
                                IID_ISourcesRowset,
                                reinterpret_cast<void**>( &pISourcesRowset ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"CoCreateInstance", __FUNCTION__ );
    }

    if ( pISourcesRowset == nullptr )
    {
        throw NullException( L"pISourcesRowset", __FUNCTION__ );
    }
    AutoReleaseISourcesRowset.Set( pISourcesRowset );

    // Get the rowset
    hResult = pISourcesRowset->GetSourcesRowset( nullptr,
                                                 IID_IRowset,
                                                 0,
                                                 nullptr,
                                                 reinterpret_cast<IUnknown**>( &pIRowset ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"ISourcesRowset::GetSourcesRowset", __FUNCTION__ );
    }

    if ( pIRowset == nullptr )
    {
        throw NullException( L"pIRowset", __FUNCTION__ );
    }
    AutoReleaseIRowset.Set( pIRowset );

    // Get the accessor
    hResult = pIRowset->QueryInterface(IID_IAccessor, reinterpret_cast<void**>( &pIAccessor ) );
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IRowset::QueryInterface", __FUNCTION__ );
    }

    if ( pIAccessor == nullptr )
    {
        throw NullException( L"pIAccessor", __FUNCTION__ );
    }
    AutoReleaseIAccessor.Set( pIAccessor );

    // Result column 1: Column name = SOURCES_NAME, Type indicator = DBTYPE_WSTR
    memset( &rgBindings, 0, sizeof ( rgBindings ) );
    rgBindings[ 0 ].iOrdinal   = 1;
    rgBindings[ 0 ].obValue    = 0;
    rgBindings[ 0 ].dwPart     = DBPART_VALUE;
    rgBindings[ 0 ].dwMemOwner = DBMEMOWNER_CLIENTOWNED;
    rgBindings[ 0 ].eParamIO   = DBPARAMIO_NOTPARAM;
    rgBindings[ 0 ].wType      = DBTYPE_WSTR;
    rgBindings[ 0 ].cbMaxLen   = COLUMN_LEN_BYTES;

    // Result column 3: Column name = SOURCES_DESCRIPTION,
    // Type indicator = DBTYPE_WSTR
    rgBindings[ 1 ].iOrdinal   = 3;
    rgBindings[ 1 ].obValue    = COLUMN_LEN_BYTES;  // Follows after column 1
    rgBindings[ 1 ].dwPart     = DBPART_VALUE;
    rgBindings[ 1 ].dwMemOwner = DBMEMOWNER_CLIENTOWNED;
    rgBindings[ 1 ].eParamIO   = DBPARAMIO_NOTPARAM;
    rgBindings[ 1 ].wType      = DBTYPE_WSTR;
    rgBindings[ 1 ].cbMaxLen   = COLUMN_LEN_BYTES;

    hResult = pIAccessor->CreateAccessor( DBACCESSOR_ROWDATA,
                                          NUM_RESULT_COLUMNS,
                                          rgBindings,
                                          sizeof ( wzBuffer ), &hAccessor, nullptr);
    if ( FAILED( hResult ) )
    {
        throw ComException( hResult, L"IAccessor::CreateAccessor", __FUNCTION__ );
    }

    if ( hAccessor == 0 )
    {
        throw NullException( L"hAccessor", __FUNCTION__ );
    }

    // Fetch
    pRows   = &rghRows[ 0 ];
    hResult = pIRowset->GetNextRows( DB_NULL_HCHAPTER, 0, MAX_FETCH_ROWS, &cRowsObtained, &pRows );
    if ( FAILED( hResult ) )
    {
        pIAccessor->ReleaseAccessor( hAccessor, nullptr );
        throw ComException( hResult, L"IRowset::GetNextRows", __FUNCTION__ );
    }

    // Need to release the accessor
    try
    {
        // Copy results to buffer, one row at a time. Will continue on errors.
        for( i = 0; i < cRowsObtained; i++)
        {
            memset( wzBuffer, 0, sizeof ( wzBuffer ) );
            hResult = pIRowset->GetData( rghRows[ i ], hAccessor, wzBuffer );
            if ( SUCCEEDED( hResult ) )
            {
                // SOURCES_NAME
                wzBuffer[ COLUMN_LEN_CHARS - 1 ] = PXS_CHAR_NULL;
                Name = wzBuffer;
                Name.Trim();
                if ( Name.GetLength() )
                {
                    // SOURCES_DESCRIPTION
                    wzBuffer[ ARRAYSIZE( wzBuffer ) - 1 ] = PXS_CHAR_NULL;
                    Description = wzBuffer + COLUMN_LEN_CHARS;
                    Description.Trim();
                    Provider.SetNameValue( Name, Description );
                    pOleDbProviders->Add( Provider );
                }
            }
            else
            {
                PXSLogComWarn( hResult, L"IRowset::GetData failed." );
            }
        }
    }
    catch ( const Exception& )
    {
        pIAccessor->ReleaseAccessor( hAccessor, nullptr );
        throw;
    }
    pIAccessor->ReleaseAccessor( hAccessor, nullptr );
}

//===============================================================================================//
//  Description:
//      Get operating system information in audit record format
//
//  Parameters:
//      pRecord - receives the data
//
//  Returns:
//      void
//===============================================================================================//
void AuditData::GetOperatingSystemRecord( AuditRecord* pRecord )
{
    String    Value, RegisteredOwner, RegisteredOrgName, ProductID;
    String    PlusVersionNumber, InstallDate;
    Formatter Format;
    Directory DirObject;
    WindowsInformation WindowsInfo;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_OS );

    // OS Name, always want to add a value even it is ""
    Value = PXS_STRING_EMPTY;
    WindowsInfo.GetName( &Value );
    pRecord->Add( PXS_OS_NAME, Value );

    // Edition
    Value = PXS_STRING_EMPTY;
    WindowsInfo.GetEdition( &Value );
    pRecord->Add( PXS_OS_EDITION, Value );

    // Registration details
    WindowsInfo.GetRegistrationDetails( &RegisteredOwner,
                                        &RegisteredOrgName,
                                        &ProductID, &PlusVersionNumber, &InstallDate );
    pRecord->Add( PXS_OS_INSTALL_DATE    , InstallDate );
    pRecord->Add( PXS_OS_REGISTERED_OWNER, RegisteredOwner );
    pRecord->Add( PXS_OS_REGISTERED_ORG  , RegisteredOrgName );
    pRecord->Add( PXS_OS_PRODUCT_ID      , ProductID );

    // Major Version
    Value = Format.UInt32( WindowsInfo.GetMajorVersion() );
    pRecord->Add( PXS_OS_MAJOR_VERSION_NUMBER, Value );

    // Minor Version
    Value = Format.UInt32( WindowsInfo.GetMinorVersion() );
    pRecord->Add( PXS_OS_MINOR_VERSION_NUMBER, Value );

    // Build number
    Value = Format.UInt32( WindowsInfo.GetBuildNumber() );
    pRecord->Add( PXS_OS_BUILD_NUMBER, Value );

    // CSD Version (Service Pack)
    Value = PXS_STRING_EMPTY;
    WindowsInfo.GetCSDVersion( &Value );
    pRecord->Add( PXS_OS_CSD_VERSION, Value );

    // Service Pack major and minor, only NT4(SP6) and NT5+
    Value = PXS_STRING_EMPTY;
    WindowsInfo.GetServicePackMajorDotMinor( &Value );
    pRecord->Add( PXS_OS_SERVICE_PACK_MAJOR_MINOR, Value );

    // Plus! Version number
    pRecord->Add( PXS_OS_PLUS_VERSION_NUMBER, PlusVersionNumber );

    // DirectX Version
    Value = PXS_STRING_EMPTY;
    WindowsInfo.GetDirectXVersionXP( &Value );
    pRecord->Add( PXS_OS_DIRECTX_VERSION, Value );

    // Windows Directory
    Value = PXS_STRING_EMPTY;
    WindowsInfo.GetWindowsDirectoryPath( &Value );
    pRecord->Add( PXS_OS_WINDOWS_DIRECTORY, Value );

    // System directory
    Value = PXS_STRING_EMPTY;
    WindowsInfo.GetSystemDirectoryPath( &Value );
    pRecord->Add( PXS_OS_SYSTEM_DIRECTORY, Value );

    // Temporary directory
    Value = PXS_STRING_EMPTY;
    DirObject.GetTempDirectory( &Value );
    pRecord->Add( PXS_OS_TEMPORARY_DIRECTORY, Value );

    // Add the operating system language
    Value = PXS_STRING_EMPTY;
    WindowsInfo.GetLanguage( &Value );
    pRecord->Add( PXS_OS_LANGUAGE, Value );

    // Add the bit size of the operating system
    Value = Format.UInt32( WindowsInfo.GetNumberOperatingSystemBits() );
    pRecord->Add( PXS_OS_NUMBER_OF_BITS, Value );
}

//===============================================================================================//
//  Description:
//      Reserved for future use
//
//  Parameters:
//      Identifier - receives the identifier
//
//  Remarks:
//      Presently not used. However, the Computer_Master table has a called
//      column named "Other_Identifer" in the event want to store an
//      additional computer identifier over and above the ones already in
//      use such as SMBIOS UUID and MAC Address.
//
//  Returns:
//      void
//===============================================================================================//
void AuditData::GetOtherIdentifier( String* pIdentifier )
{
    if ( pIdentifier == nullptr )
    {
        throw ParameterException( L"pIdentifier", __FUNCTION__ );
    }
    *pIdentifier = PXS_STRING_EMPTY;
}

//===============================================================================================//
//  Description:
//      Fill an array of selected regional settings
//
//  Parameters:
//      pRecords - array to receive regional settings data
//
//  Remarks:
//      On XP, DaylightDate and StandardDate members of TIME_ZONE_INFORMATION
//      filled by GetTimeZoneInformation do not match the behaviour observed
//      by adjusting the system clock.
//
//  Returns:
//      void
//===============================================================================================//
void AuditData::GetRegionalSettingsRecords( TArray< AuditRecord >* pRecords  )
{
    LONG    lBias = 0;
    DWORD   i = 0, dwTZI = 0;
    String  Value, Insert1;
    wchar_t szBuffer[ 512 ] = { 0 };        // Enough for a regional setting
    Formatter   Format;
    AuditRecord Record;
    TIME_ZONE_INFORMATION tzInfo;

    struct TYPE_REGIONAL_DATA
    {
        LCTYPE  LCType;
        LPCWSTR pszType;
        LPCWSTR pszName;
    } Data[] = {
        { LOCALE_SCOUNTRY       , L"Locale"  , L"Country"            },
        { LOCALE_SLANGUAGE      , L"Locale"  , L"Language"           },
        { LOCALE_SNATIVEDIGITS  , L"Number"  , L"Digits"             },
        { LOCALE_STHOUSAND      , L"Number"  , L"Digit Separator"    },
        { LOCALE_SDECIMAL       , L"Number"  , L"Decimal Separator"  },
        { LOCALE_IDIGITS        , L"Number"  , L"Decimal Places"     },
        { LOCALE_IMEASURE       , L"Number"  , L"Measurement System" },
        { LOCALE_SNATIVECURRNAME, L"Currency", L"Name"               },
        { LOCALE_SCURRENCY      , L"Currency", L"Symbol"             },
        { LOCALE_SINTLSYMBOL    , L"Currency", L"ISO Code"           },
        { LOCALE_SMONTHOUSANDSEP, L"Currency", L"Money Separator"    },
        { LOCALE_SMONDECIMALSEP , L"Currency", L"Decimal Separator"  },
        { LOCALE_ICURRDIGITS    , L"Currency", L"Decimal Places"     },
        { LOCALE_SSHORTDATE     , L"Date"    , L"Short Date"         },
        { LOCALE_SLONGDATE      , L"Date"    , L"Long Date"          },
        { LOCALE_ICALENDARTYPE  , L"Date"    , L"Calendar"           },
        { LOCALE_STIMEFORMAT    , L"Time"    , L"Time Format"        },
        { LOCALE_S1159          , L"Time"    , L"Ante Meridiem"      },
        { LOCALE_S2359          , L"Time"    , L"Post Meridiem"      } };

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    for ( i = 0; i < ARRAYSIZE( Data ); i++ )
    {
        Value = PXS_STRING_EMPTY;
        memset( szBuffer, 0, sizeof( szBuffer ) );
        if ( GetLocaleInfo( LOCALE_USER_DEFAULT,
                            Data[ i ].LCType, szBuffer, ARRAYSIZE( szBuffer ) ) == 0 )
        {
            Insert1 = Data[ i ].pszName;
            PXSLogSysError1( GetLastError(), L"GetLocaleInfo failed for '%%1'.", Insert1 );
        }
        szBuffer[ ARRAYSIZE( szBuffer ) - 1 ] = PXS_CHAR_NULL;

        // For a few values, need to interpret the data
        if ( LOCALE_IMEASURE == Data[ i ].LCType )
        {
            if ( szBuffer[ 0 ] == '0' )
            {
                Value = L"Metric";
            }
            else
            {
                Value = L"U.S.";
            }
        }
        else if ( LOCALE_ICALENDARTYPE == Data[ i ].LCType )
        {
            TranslateCalenderType( szBuffer, &Value );
        }
        else
        {
            Value = szBuffer;
        }
        Value.Trim();

        // Add to output array
        Record.Reset( PXS_CATEGORY_REGION_SETTINGS );
        Record.Add( PXS_REGION_SETTINGS_ITEM  ,  Data[ i ].pszType );
        Record.Add( PXS_REGION_SETTINGS_NAME  ,  Data[ i ].pszName );
        Record.Add( PXS_REGION_SETTINGS_SETTING, Value );
        pRecords->Add( Record );
    }

    // Add time zone information
    memset( &tzInfo, 0, sizeof ( tzInfo ) );
    dwTZI = GetTimeZoneInformation( &tzInfo );
    if ( dwTZI == TIME_ZONE_ID_INVALID )
    {
        PXSLogSysError( GetLastError(), L"GetTimeZoneInformation failed." );
    }
    tzInfo.StandardName[ARRAYSIZE( tzInfo.StandardName ) - 1] = PXS_CHAR_NULL;
    tzInfo.DaylightName[ARRAYSIZE( tzInfo.DaylightName ) - 1] = PXS_CHAR_NULL;

    Record.Reset( PXS_CATEGORY_REGION_SETTINGS );
    Record.Add( PXS_REGION_SETTINGS_ITEM, L"Time" );
    Record.Add( PXS_REGION_SETTINGS_NAME, L"Time Zone" );
    if ( dwTZI == TIME_ZONE_ID_DAYLIGHT )
    {
        Record.Add( PXS_REGION_SETTINGS_SETTING, tzInfo.DaylightName );
    }
    else
    {
        Record.Add( PXS_REGION_SETTINGS_SETTING, tzInfo.StandardName );
    }
    pRecords->Add( Record );

    // Time Difference
    Record.Reset( PXS_CATEGORY_REGION_SETTINGS );
    Record.Add( PXS_REGION_SETTINGS_ITEM, L"Time" );
    Record.Add( PXS_REGION_SETTINGS_NAME, L"GMT Difference (mins.)" );
    lBias = tzInfo.Bias;
    if ( dwTZI == TIME_ZONE_ID_DAYLIGHT )
    {
        if ( tzInfo.DaylightDate.wMonth )
        {
            lBias = PXSAddInt32( lBias, tzInfo.DaylightBias );
        }
    }
    else
    {
        if ( tzInfo.StandardDate.wMonth )
        {
            lBias = PXSAddInt32( lBias, tzInfo.StandardBias );
        }
    }
    Value = Format.Int32( lBias );
    Record.Add( PXS_REGION_SETTINGS_SETTING, Value );
    pRecords->Add( Record );
}

//===============================================================================================//
//  Description:
//      Get the system files that match the specified filter
//
//  Parameters:
//      pszFilter - filter for FindFirstFile
//      pRecords  - receives an array of audit records
//
//  Returns:
//      void
//===============================================================================================//
void AuditData::GetSystemFilesRecords( LPCWSTR pszFilter, TArray< AuditRecord >* pRecords )
{
    size_t      i = 0, numFileNames = 0;
    File        FileObject;
    String      SystemDirectoryPath, FullPath, DotExtension;
    String      ManufacturerName, VersionString, Description, LastWrite;
    FILETIME    lastWrite;
    Directory   DirObject;
    Formatter   Format;
    AuditRecord Record;
    FileVersion FileVer;
    StringArray FileNames;
    SystemInformation  SystemInfo;

    if ( ( pszFilter == nullptr ) || ( *pszFilter == PXS_CHAR_NULL ) )
    {
        throw ParameterException( L"pszFilter", __FUNCTION__ );
    }

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Get the system directory, path includes the trailing separator
    SystemInfo.GetSystemDirectoryPath( &SystemDirectoryPath );
    if ( SystemDirectoryPath.IsEmpty() )
    {
        throw SystemException( ERROR_BAD_PATHNAME, L"SystemDirectoryPath = 0", __FUNCTION__ );
    }

    // Want .ext form
    if ( *pszFilter != '.' )
    {
        DotExtension = '.';
    }
    DotExtension += pszFilter;
    DirObject.ListFiles( SystemDirectoryPath, DotExtension, &FileNames );
    FileNames.Sort( true );

    // Allocate some characters for string operations
    LastWrite.Allocate( 256 );
    ManufacturerName.Allocate( 256 );
    VersionString.Allocate( 256 );
    Description.Allocate( 256 );
    numFileNames = FileNames.GetSize();
    for ( i = 0; i < numFileNames; i++ )
    {
        // This is a lengthy task so see if want to stop
        if ( g_pApplication && g_pApplication->GetStopBackgroundTasks() )
        {
            break;
        }

        // Get the last write time and version info for each file
        try
        {
            FullPath  = SystemDirectoryPath;
            FullPath += FileNames.Get( i );

            memset( &lastWrite, 0, sizeof ( lastWrite ) );
            FileObject.GetTimes( FullPath, nullptr, nullptr, &lastWrite );
            LastWrite = Format.FileTimeToLocalTimeIso( lastWrite );

            ManufacturerName = PXS_STRING_EMPTY;
            VersionString    = PXS_STRING_EMPTY;
            FileVer.GetVersion( FullPath, &ManufacturerName, &VersionString, &Description );
            Record.Reset( PXS_CATEGORY_SYSTEM_FILES );
            Record.Add( PXS_SYSTEM_FILES_NAME        ,  FileNames.Get( i ) );
            Record.Add( PXS_SYSTEM_FILES_VERSION     ,  VersionString );
            Record.Add( PXS_SYSTEM_FILES_MODIFIED    ,  LastWrite );
            Record.Add( PXS_SYSTEM_FILES_MANUFACTURER, ManufacturerName );
            pRecords->Add( Record );
        }
        catch ( const Exception& e )
        {
            PXSLogException( L"Version or date error.", e, __FUNCTION__);
        }
    }
}

//===============================================================================================//
//  Description:
//      Get a system overview
//
//  Parameters:
//      SmbiosInfo - SmbiosInformation object with data used for the overview.
//                   Must have loaded this class with the SMBIOS data tables
//      LocalTime  - The computer local time
//      pRecord    - receives the data
//
//  Remarks:
//      LocalTime is the time obtained when the audit is run so that
//      the same value is used everywhere
//
//  Returns:
//      void
//===============================================================================================//
void AuditData::GetSystemOverviewRecord( const String& LocalTime, AuditRecord* pRecord )
{
    BYTE        packageID = 0;
    DWORD       data   = 0, totalPhysMB = 0, numBits = 0;
    time_t      upTime = 0;
    UINT64      hddCapacity = 0;
    String      Value, LocaleMB;
    Formatter   Format;
    AuditRecord MemoryRecord, BiosRecord;
    CpuInformation     CpuInfo;
    DriveInformation   Volumeinfo;
    SystemInformation  SystemInfo;
    DisplayInformation DisplayInfo;
    WindowsInformation WindowsInfo;
    TArray< BYTE >  PackageIDs;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_SYSTEM_OVERVIEW );

    try
    {
        PXSGetResourceString( PXS_IDS_136_MB, &LocaleMB );
        m_SmbiosInfo.ReadSmbiosData();
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"SMBIOS read error.", e, __FUNCTION__ );
    }

    // Computer name
    Value = PXS_STRING_EMPTY;
    try
    {
        SystemInfo.GetComputerNetBiosName( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"NetBIOS name error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_COMPUTER_NAME, Value );

    // Domain Name
    Value = PXS_STRING_EMPTY;
    try
    {
        SystemInfo.GetDomainNetBiosName( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Domain name error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_DOMAIN_NAME, Value );

    // Site Name
    Value = PXS_STRING_EMPTY;
    try
    {
        SystemInfo.GetComputerSiteName( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Site name error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_SITE_NAME, Value );

    // Computer roles
    Value = PXS_STRING_EMPTY;
    try
    {
        SystemInfo.GetComputerRoles( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Computer roles error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_ROLES, Value );

    // Computer description
    Value = PXS_STRING_EMPTY;
    try
    {
        SystemInfo.GetComputerDescription( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Computer description error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_DESCRIPTION, Value );

    // Operating system
    Value = PXS_STRING_EMPTY;
    try
    {
        WindowsInfo.GetNameAndEdition( &Value );
        Value.Trim();
        if ( Value.EndsWithStringI( L"bit" ) == false )
        {
            numBits = SystemInfo.GetNumberOperatingSystemBits();
            Value += PXS_CHAR_SPACE;
            Value += Format.UInt32( numBits );
            Value += L"-Bit";
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Operting system name error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_OPERATING_SYS, Value );

    // SMBIOS - manufacturer
    Value = PXS_STRING_EMPTY;
    try
    {
        m_SmbiosInfo.GetSystemManufacturer( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"System manufacturer name error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_MANUFACTURER, Value );

    // SMBIOS - product name
    Value = PXS_STRING_EMPTY;
    try
    {
        m_SmbiosInfo.GetProductName( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Product name error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_MODEL, Value );

    // SMBIOS - serial number
    Value = PXS_STRING_EMPTY;
    try
    {
        m_SmbiosInfo.GetSystemSerialNumber( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"System serial number error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_SERIAL_NUMBER, Value );

    // SMBIOS - asset tag number
    Value = PXS_STRING_EMPTY;
    try
    {
        m_SmbiosInfo.GetChassisAssetTag( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Chassis asset tag error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_ASSET_TAG, Value );

    // Number of processors, aka packages
    Value = PXS_STRING_EMPTY;
    try
    {
        PackageIDs.RemoveAll();
        CpuInfo.GetPackageIDs( &PackageIDs );
        Value = Format.SizeT( PackageIDs.GetSize() );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Number of processors error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_NUM_OF_PROCS, Value );

    // Processor Name and Speed, use the first found processor
    Value = PXS_STRING_EMPTY;
    try
    {
        PackageIDs.RemoveAll();
        CpuInfo.GetPackageIDs( &PackageIDs );
        if ( PackageIDs.GetSize() )
        {
            packageID = PackageIDs.Get( 0 );
            CpuInfo.GetNameAndSpeedMHz( packageID, &Value, &data );

            // Add speed if do not in the description
            if ( ( Value.IndexOfI( L"Mhz" ) == PXS_MINUS_ONE ) &&
                 ( Value.IndexOfI( L"Ghz" ) == PXS_MINUS_ONE )  )
            {
                Value += L", ";
                Value += Format.UInt32( data );
                Value += L"MHz";
            }
        }
        else
        {
            PXSLogAppWarn( L"No processor package indentifiers found." );
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Processor name/speed error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_PROCESSOR_DESC, Value );

    // Memory
    Value = PXS_STRING_EMPTY;
    try
    {
        totalPhysMB = GetMemoryInformationRecord( &MemoryRecord );
        Value       = Format.UInt32( totalPhysMB );
        Value      += LocaleMB;
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Memory information error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_TOTAL_MEMORY_MB, Value );

    Value = PXS_STRING_EMPTY;
    try
    {
        hddCapacity = Volumeinfo.GetTotalHDDCapacity();
        Value       = Format.StorageBytes( hddCapacity );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Disk capacity error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_TOTAL_HDD, Value );

    // Display
    Value = PXS_STRING_EMPTY;
    try
    {
        DisplayInfo.GetDescriptionString( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Dispaly description error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_DISPLAY, Value );

    // BIOS Version
    Value = PXS_STRING_EMPTY;
    try
    {
        GetBIOSIdentificationRecord( &BiosRecord );
        BiosRecord.GetItemValue( PXS_BIOS_VERSION_BIOS_VERSION, &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"BIOS version error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_BIOS_VERSION, Value );

    // User name
    Value = PXS_STRING_EMPTY;
    try
    {
        SystemInfo.GetCurrentUserName( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"User name error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_USER_ACCOUNT, Value );

    // System Uptime
    Value = PXS_STRING_EMPTY;
    try
    {
        upTime = SystemInfo.GetUpTime();
        if ( upTime )
        {
            Value = Format.SecondsToDDHHMM( upTime );
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"System overview error.", e, __FUNCTION__ );
    }
    pRecord->Add( PXS_SYS_OVERVIEW_SYS_UPTIME_SEC, Value );

    // The local time in ISO format
    pRecord->Add( PXS_SYS_OVERVIEW_LOCAL_TIME, LocalTime );
}

//===============================================================================================//
//  Description:
//      Make the audit master record for a database insert
//
//  Parameters:
//      pAuditMaster - receives the audit record data
//
//  Remarks:
//      Some fields like PXS_COMP_MASTER_DB_USER_NAME are not added because
//      these are known only to the database when the record is inserted or
//      updated
//
//  Returns:
//      void
//===============================================================================================//
void AuditData::MakeAuditMasterRecord( AuditRecord* pAuditMaster )
{
    String     AuditGUID, LocalTime, ComputerUTC;
    Formatter  Format;
    SYSTEMTIME st;

    if ( pAuditMaster == nullptr )
    {
        throw ParameterException( L"pAuditMaster", __FUNCTION__ );
    }
    pAuditMaster->Reset( PXS_CATEGORY_AUDIT_MASTER );

    AuditGUID = Format.CreateGuid();
    pAuditMaster->Add( PXS_AUDIT_MASTER_AUDIT_GUID, AuditGUID );

    // Get the computer local time
    memset( &st, 0, sizeof ( st ) );
    GetLocalTime( &st );
    LocalTime = Format.SystemTimeToIso( st );
    pAuditMaster->Add( PXS_AUDIT_MASTER_COMPUTER_LOCAL, LocalTime );

    // Get the computer UTC time
    memset( &st, 0, sizeof ( st ) );
    GetSystemTime( &st );
    ComputerUTC = Format.SystemTimeToIso( st );
    pAuditMaster->Add( PXS_AUDIT_MASTER_COMPUTER_UTC, ComputerUTC );
}

//===============================================================================================//
//  Description:
//      Make the audit master record for a database insert
//
//  Parameters:
//      pszRegPathComputerGuid  - registry folder containing the WinAuditGuid
//      pszWinAuditComputerGuid - the name of the WinAuditGuid
//      pComputerMaster         - receives the computer_master record
//
//  Remarks:
//      Some fields like PXS_COMP_MASTER_DB_USER_NAME are not added
//      because these are known only to the database when the record is
//      inserted/updated
//
//  Returns:
//      void
//===============================================================================================//
void AuditData::MakeComputerMasterRecord( AuditRecord* pComputerMaster )
{
    File      FileObject;
    String    Value, RegisteredOwner, RegisteredOrgName, ProductID;
    String    PlusVersionNumber, InstallDate, GuidFilePath;
    Formatter Format;
    AuditData Identifier;
    TcpIpInformation   TcpIpInfo;
    SystemInformation  SystemInfo;
    WindowsInformation WindowsInfo;

    if ( pComputerMaster == nullptr )
    {
        throw ParameterException( L"pComputerMaster", __FUNCTION__ );
    }
    pComputerMaster->Reset( PXS_CATEGORY_COMPUTER_MASTER );
    m_SmbiosInfo.ReadSmbiosData();

    // MAC Address
    Value = PXS_STRING_EMPTY;
    try
    {
        TcpIpInfo.GetMacAddress( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting MAC address.", e, __FUNCTION__ );
    }
    pComputerMaster->Add( PXS_COMP_MASTER_MAC_ADDRESS, Value );

    // SMBIOS Universally Unique ID
    Value = PXS_STRING_EMPTY;
    try
    {
        m_SmbiosInfo.GetSystemUUID( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting SMBIOS UUID.", e, __FUNCTION__ );
    }
    pComputerMaster->Add( PXS_COMP_MASTER_SMBIOS_UUID, Value );

    // Asset Tag
    Value = PXS_STRING_EMPTY;
    try
    {
        m_SmbiosInfo.GetChassisAssetTag( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting asset tag.", e, __FUNCTION__ );
    }
    pComputerMaster->Add( PXS_COMP_MASTER_ASSET_TAG, Value );

    // Fully Qualified Domain Name
    Value = PXS_STRING_EMPTY;
    try
    {
        SystemInfo.GetFullyQualifiedDomainName( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting fully qualified domain name.", e, __FUNCTION__ );
    }
    pComputerMaster->Add( PXS_COMP_MASTER_FQDN, Value );

    // Site Name
    Value = PXS_STRING_EMPTY;
    try
    {
        SystemInfo.GetComputerSiteName( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting site name.", e, __FUNCTION__ );
    }
    pComputerMaster->Add( PXS_COMP_MASTER_SITE_NAME, Value );

    // Domain name
    Value = PXS_STRING_EMPTY;
    try
    {
        SystemInfo.GetDomainNetBiosName( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting domain name.", e, __FUNCTION__ );
    }
    pComputerMaster->Add( PXS_COMP_MASTER_DOMAIN_NAME, Value );

    // Computer name
    Value = PXS_STRING_EMPTY;
    try
    {
        SystemInfo.GetComputerNetBiosName( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting computer name.", e, __FUNCTION__ );
    }
    pComputerMaster->Add( PXS_COMP_MASTER_COMPUTER_NAME, Value );

    // OS Product ID
    ProductID = PXS_STRING_EMPTY;
    try
    {
        WindowsInfo.GetRegistrationDetails( &RegisteredOwner,
                                            &RegisteredOrgName,
                                            &ProductID, &PlusVersionNumber, &InstallDate );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting Windows product ID.", e, __FUNCTION__ );
    }
    pComputerMaster->Add( PXS_COMP_MASTER_OS_PRODUCT_ID, ProductID );

    // Other Identifier
    Value = PXS_STRING_EMPTY;
    try
    {
        Identifier.GetOtherIdentifier( &Value );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting other identifier.", e, __FUNCTION__ );
    }
    pComputerMaster->Add( PXS_COMP_MASTER_OTHER_ID, Value );

    // WinAudit GUID, must have this
    Value = PXS_STRING_EMPTY;
    PXSGetWinAuditGuidFilePath( &GuidFilePath );
    if ( FileObject.Exists( GuidFilePath ) == false )
    {
        PXSWriteWinAuditGuidFile();
    }
    PXSReadWinAuditGuidFile( &Value );
    if ( Format.IsValidStringGuid( Value ) == false )
    {
        PXSLogAppInfo2( L"Invalid WinAuditGUID '%%1' in %%2.", Value, GuidFilePath );
        throw SystemException( ERROR_INVALID_DATA, PXS_WINAUDIT_COMPUTER_GUID, __FUNCTION__ );
    }
    PXSLogAppInfo2( L"Found WinAuditGUID '%%1' in %%2.", Value, GuidFilePath );
    pComputerMaster->Add( PXS_COMP_MASTER_WINAUDIT_GUID, Value );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Translate a LOCALE_ICALENDARTYPE to a string
//
//  Parameters:
//      pszCalendar - LOCALE_ICALENDARTYPE value
//      pMeaning    - receives the translation
//
//  Remarks:
//      See "Calendar Identifiers"
//
//  Returns:
//      void
//===============================================================================================//
void AuditData::TranslateCalenderType( LPCWSTR pszCalendar, String* pMeaning )
{
    size_t i = 0;

    struct _MEANINGS
    {
        LPCWSTR pszCalendar;
        LPCWSTR pszMeaning;
    } Meanings[] = { { L"1" , L"Gregorian"                        },
                     { L"2" , L"Gregorian"                        },
                     { L"3" , L"Japanese Emperor Era"             },
                     { L"4" , L"Taiwan Calendar"                  },
                     { L"5" , L"Tangun Era (Korea)"               },
                     { L"6" , L"Hijri (Arabic lunar)"             },
                     { L"7" , L"Thai"                             },
                     { L"8" , L"Hebrew (Lunar)"                   },
                     { L"9" , L"Gregorian Middle East French"     },
                     { L"10", L"Gregorian Arabic"                 },
                     { L"11", L"Gregorian Transliterated English" },
                     { L"12", L"Gregorian Transliterated French"  } };

    if ( pszCalendar == nullptr )
    {
        throw ParameterException( L"pszCalendar", __FUNCTION__ );
    }

    if ( pMeaning == nullptr )
    {
        throw ParameterException( L"pMeaning", __FUNCTION__ );
    }
    *pMeaning = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Meanings ); i++ )
    {
        if ( lstrcmp( pszCalendar, Meanings[ i ].pszCalendar ) == 0 )
        {
            *pMeaning = Meanings[ i ].pszMeaning;
            break;
        }
    }

    // If unrecognised...
    if ( pMeaning->IsEmpty() )
    {
       *pMeaning = pszCalendar;
       PXSLogAppWarn1( L"Unrecognised calendar type '%%1'.", *pMeaning );
    }
}
