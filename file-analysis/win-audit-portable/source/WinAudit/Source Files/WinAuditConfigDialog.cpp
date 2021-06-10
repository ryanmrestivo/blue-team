///////////////////////////////////////////////////////////////////////////////////////////////////
//
// WinAudit Configuration Dialog Class Implementation
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
#include "WinAudit/Header Files/WinAuditConfigDialog.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Font.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StaticControl.h"
#include "PxsBase/Header Files/StringT.h"

// 5. This Project
#include "WinAudit/Header Files/Resources.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default Constructor
WinAuditConfigDialog::WinAuditConfigDialog()
                     :m_Settings(),
                      m_SystemOverviewCheckBox(),
                      m_InstalledSoftwareCheckBox(),
                      m_OperatingSystemCheckBox(),
                      m_PeripheralsCheckBox(),
                      m_SecurityCheckBox(),
                      m_GroupsAndUsersCheckBox(),
                      m_ScheduledTasksCheckBox(),
                      m_UptimeStatisticsCheckBox(),
                      m_ErrorLogsCheckBox(),
                      m_EnvironmentVarsCheckBox(),
                      m_RegionalSettingsCheckBox(),
                      m_WindowsNetworkCheckBox(),
                      m_NetworkTcpIpCheckBox(),
                      m_HardwareDevicesCheckBox(),
                      m_DisplaysCheckBox(),
                      m_DisplayAdaptersCheckBox(),
                      m_InstalledPrintersCheckBox(),
                      m_BiosVersionCheckBox(),
                      m_SystemManagementCheckBox(),
                      m_ProcessorsCheckBox(),
                      m_MemoryCheckBox(),
                      m_PhysicalDisksCheckBox(),
                      m_DrivesCheckBox(),
                      m_CommPortsCheckBox(),
                      m_StartupProgramsCheckBox(),
                      m_ServicesCheckBox(),
                      m_RunningProgramsCheckBox(),
                      m_OdbcInformationsCheckBox(),
                      m_OleDbProvidersCheckBox(),
                      m_SoftwareMeteringCheckBox(),
                      m_UserLogonStatsCheckBox(),
                      m_SelectNoneButton(),
                      m_SelectAllButton(),
                      m_CancelButton(),
                      m_ApplyButton()
{
    try
    {
        SetSize( 650, 340 );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Copy constructor - not allowed so no implementation


// Destructor
WinAuditConfigDialog::~WinAuditConfigDialog()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator - not allowed, no implementation

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the configuration settings
//
//  Parameters:
//      pSettings - receives the settings
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditConfigDialog::GetConfigurationSettings( ConfigurationSettings* pSettings ) const
{
    if ( pSettings == nullptr )
    {
        throw ParameterException( L"pSettings", __FUNCTION__ );
    }
    *pSettings = m_Settings;
}

//===============================================================================================//
//  Description:
//      Set the configuration settings
//
//  Parameters:
//      Settings - the settings
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditConfigDialog::SetConfigurationSettings( const ConfigurationSettings& Settings )
{
    m_Settings = Settings;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Handle WM_COMMAND events.
//
//  Parameters:
//        wParam -
//        lParam -
//
//  Returns:
//      0 if handled, else non-zero.
//===============================================================================================//
LRESULT WinAuditConfigDialog::CommandEvent( WPARAM wParam, LPARAM lParam )
{
    bool     state  = false;
    LRESULT  result = 0;
    WORD     wID    = LOWORD(wParam);  // Item, control, or accelerator
    HWND     hWnd   = reinterpret_cast<HWND>( lParam );    // Window handle

    // Ensure controls have been created
    if ( m_bControlsCreated == false ) return result;

    // Don't let any exceptions propagate out of event handler
    try
    {
        if ( ( wID == IDCANCEL ) ||
             ( IsClickFromButton( m_CancelButton, wParam, lParam ) ) )
        {
            EndDialog( m_hWindow, IDCANCEL );
        }
        else if ( IsClickFromButton( m_ApplyButton, wParam, lParam ) )
        {
          m_Settings.systemOverview   = m_SystemOverviewCheckBox.GetState();
          m_Settings.installedSoftware= m_InstalledSoftwareCheckBox.GetState();
          m_Settings.operatingSystem  = m_OperatingSystemCheckBox.GetState();
          m_Settings.peripherals      = m_PeripheralsCheckBox.GetState();
          m_Settings.security         = m_SecurityCheckBox.GetState();
          m_Settings.groupsAndUsers   = m_GroupsAndUsersCheckBox.GetState();
          m_Settings.scheduledTasks   = m_ScheduledTasksCheckBox.GetState();
          m_Settings.uptimeStatistics = m_UptimeStatisticsCheckBox.GetState();
          m_Settings.errorLogs        = m_ErrorLogsCheckBox.GetState();
          m_Settings.environmentVars  = m_EnvironmentVarsCheckBox.GetState();
          m_Settings.regionalSettings = m_RegionalSettingsCheckBox.GetState();
          m_Settings.windowsNetwork   = m_WindowsNetworkCheckBox.GetState();
          m_Settings.networkTcpIp     = m_NetworkTcpIpCheckBox.GetState();
          m_Settings.hardwareDevices  = m_HardwareDevicesCheckBox.GetState();
          m_Settings.displays         = m_DisplaysCheckBox.GetState();
          m_Settings.displayAdapters  = m_DisplayAdaptersCheckBox.GetState();
          m_Settings.installedPrinters= m_InstalledPrintersCheckBox.GetState();
          m_Settings.biosVersion      = m_BiosVersionCheckBox.GetState();
          m_Settings.systemManagement = m_SystemManagementCheckBox.GetState();
          m_Settings.processors       = m_ProcessorsCheckBox.GetState();
          m_Settings.memory           = m_MemoryCheckBox.GetState();
          m_Settings.physicalDisks    = m_PhysicalDisksCheckBox.GetState();
          m_Settings.drives           = m_DrivesCheckBox.GetState();
          m_Settings.commPorts        = m_CommPortsCheckBox.GetState();
          m_Settings.startupPrograms  = m_StartupProgramsCheckBox.GetState();
          m_Settings.services         = m_ServicesCheckBox.GetState();
          m_Settings.runningPrograms  = m_RunningProgramsCheckBox.GetState();
          m_Settings.odbcInformation  = m_OdbcInformationsCheckBox.GetState();
          m_Settings.oleDbProviders   = m_OleDbProvidersCheckBox.GetState();
          m_Settings.softwareMetering = m_SoftwareMeteringCheckBox.GetState();
          m_Settings.userLogonStats   = m_UserLogonStatsCheckBox.GetState();

          EndDialog( m_hWindow, IDOK );
        }
        else if ( IsClickFromButton( m_SelectAllButton , wParam, lParam ) ||
                  IsClickFromButton( m_SelectNoneButton, wParam, lParam )  )
        {
            if ( hWnd == m_SelectAllButton.GetHwnd() )
            {
                state = true;
            }
            m_SystemOverviewCheckBox.SetState( state );
            m_InstalledSoftwareCheckBox.SetState( state );
            m_OperatingSystemCheckBox.SetState( state );
            m_PeripheralsCheckBox.SetState( state );
            m_SecurityCheckBox.SetState( state );
            m_GroupsAndUsersCheckBox.SetState( state );
            m_ScheduledTasksCheckBox.SetState( state );
            m_UptimeStatisticsCheckBox.SetState( state );
            m_ErrorLogsCheckBox.SetState( state );
            m_EnvironmentVarsCheckBox.SetState( state );
            m_RegionalSettingsCheckBox.SetState( state );
            m_WindowsNetworkCheckBox.SetState( state );
            m_NetworkTcpIpCheckBox.SetState( state );
            m_HardwareDevicesCheckBox.SetState( state );
            m_DisplaysCheckBox.SetState( state );
            m_DisplayAdaptersCheckBox.SetState( state );
            m_InstalledPrintersCheckBox.SetState( state );
            m_BiosVersionCheckBox.SetState( state );
            m_SystemManagementCheckBox.SetState( state );
            m_ProcessorsCheckBox.SetState( state );
            m_MemoryCheckBox.SetState( state );
            m_PhysicalDisksCheckBox.SetState( state );
            m_DrivesCheckBox.SetState( state );
            m_CommPortsCheckBox.SetState( state );
            m_StartupProgramsCheckBox.SetState( state );
            m_ServicesCheckBox.SetState( state );
            m_RunningProgramsCheckBox.SetState( state );
            m_OdbcInformationsCheckBox.SetState( state );
            m_OleDbProvidersCheckBox.SetState( state );
            m_SoftwareMeteringCheckBox.SetState( state );
            m_UserLogonStatsCheckBox.SetState( state );

            m_SelectAllButton.Repaint();
            m_SelectNoneButton.Repaint();
        }
        else
        {
            result = 1;    // Not handled, return non-zero
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }

    return result;
}

//===============================================================================================//
//  Description:
//      Handle the WM_INITDIALOG message
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditConfigDialog::InitDialogEvent()
{
    const   int   GAP         = 10;
    const   int   LINE_HEIGHT = 22;
    const   DWORD MAX_DISPLAY_LENGTH = 25;
    bool    rtlReading;
    int     controlWidth = 0, columnDX = 0;
    SIZE    buttonSize   = { 0, 0 }, dialogSize = { 0, 0 };
    RECT    bounds       = { 0, 0, 0, 0 };
    POINT   location     = { 0, 0 };
    LPCWSTR LPSZ_ELIPSES = L"...";
    Font    FontObject;
    String  Text;
    StaticControl Static;

    if ( m_hWindow == nullptr )
    {
        return;
    }
    rtlReading = IsRightToLeftReading();
    GetClientSize( &dialogSize );

    // Title Caption, centred
    bounds.left   = 0;
    bounds.right  = dialogSize.cx;
    bounds.top    = 5;
    bounds.bottom = 45;
    FontObject.SetPointSize( 12, nullptr );
    FontObject.SetBold( true );
    FontObject.SetUnderlined( true );
    FontObject.Create();
    Static.SetBounds( bounds );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
    PXSGetResourceString( PXS_IDS_1120_SELECT_CATEGORIES, &Text );
    Static.SetText( Text );
    Static.SetFont( FontObject );
    m_Statics.Add( Static );
    Static.Reset();

    // Set the width of the controls based on 3 columns of check boxes
    controlWidth = ( dialogSize.cx - ( 4 * GAP ) ) / 3;

    // Set initial control rectangle
    if ( rtlReading )
    {
        bounds.left = dialogSize.cx - GAP - controlWidth;
        columnDX = -(GAP + controlWidth);   // Next column is to the left
    }
    else
    {
        bounds.left = GAP;
        columnDX  = GAP + controlWidth;
    }
    bounds.right = bounds.left + controlWidth;
    bounds.top   = bounds.bottom;
    bounds.bottom= bounds.top + LINE_HEIGHT;

    // System Overview
    m_SystemOverviewCheckBox.SetBounds( bounds );
    m_SystemOverviewCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1121_SYSTEM_OVERVIEW, &Text );
    if ( Text.GetLength() > MAX_DISPLAY_LENGTH )
    {
        Text.Truncate( MAX_DISPLAY_LENGTH );
        Text += LPSZ_ELIPSES;
    }
    m_SystemOverviewCheckBox.SetText( Text );
    m_SystemOverviewCheckBox.SetState( m_Settings.systemOverview );

    // Installed Software
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_InstalledSoftwareCheckBox.SetBounds( bounds );
    m_InstalledSoftwareCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1122_INSTALLED_SOFTWARE, &Text );
    if ( Text.GetLength() > MAX_DISPLAY_LENGTH )
    {
        Text.Truncate( MAX_DISPLAY_LENGTH );
        Text += LPSZ_ELIPSES;
    }
    m_InstalledSoftwareCheckBox.SetText( Text );
    m_InstalledSoftwareCheckBox.SetState( m_Settings.installedSoftware );

    // Operating System
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_OperatingSystemCheckBox.SetBounds( bounds );
    m_OperatingSystemCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1123_OPERATING_SYSTEM, &Text );
    m_OperatingSystemCheckBox.SetText( Text );
    m_OperatingSystemCheckBox.SetState( m_Settings.operatingSystem );

    // Peripherals
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_PeripheralsCheckBox.SetBounds( bounds );
    m_PeripheralsCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1124_PERIPHERALS, &Text );
    m_PeripheralsCheckBox.SetText( Text );
    m_PeripheralsCheckBox.SetState( m_Settings.peripherals );

    // Security
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_SecurityCheckBox.SetBounds( bounds );
    m_SecurityCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1125_SECURITY, &Text );
    m_SecurityCheckBox.SetText( Text );
    m_SecurityCheckBox.SetState( m_Settings.security );

    // Groups And Users
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_GroupsAndUsersCheckBox.SetBounds( bounds );
    m_GroupsAndUsersCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1126_GROUPS_AND_USERS, &Text );
    m_GroupsAndUsersCheckBox.SetText( Text );
    m_GroupsAndUsersCheckBox.SetState( m_Settings.groupsAndUsers );

    // Scheduled Tasks
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_ScheduledTasksCheckBox.SetBounds( bounds );
    m_ScheduledTasksCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1127_SCHEDULED_TASKS, &Text );
    m_ScheduledTasksCheckBox.SetText( Text );
    m_ScheduledTasksCheckBox.SetState( m_Settings.scheduledTasks );

    // Uptime Statistics
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_UptimeStatisticsCheckBox.SetBounds( bounds );
    m_UptimeStatisticsCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1128_UPTIME_STATISITCS, &Text );
    m_UptimeStatisticsCheckBox.SetText( Text );
    m_UptimeStatisticsCheckBox.SetState( m_Settings.uptimeStatistics );

    // Error Logs
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_ErrorLogsCheckBox.SetBounds( bounds );
    m_ErrorLogsCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1129_ERROR_LOGS, &Text );
    m_ErrorLogsCheckBox.SetText( Text );
    m_ErrorLogsCheckBox.SetState( m_Settings.errorLogs );

    // Environment Variables
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_EnvironmentVarsCheckBox.SetBounds( bounds );
    m_EnvironmentVarsCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1130_ENVIRONMENT_VARIABLES, &Text );
    m_EnvironmentVarsCheckBox.SetText( Text );
    m_EnvironmentVarsCheckBox.SetState( m_Settings.environmentVars );

    // Regional Settings
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_RegionalSettingsCheckBox.SetBounds( bounds );
    m_RegionalSettingsCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1131_REGIONAL_SETTINGS, &Text );
    m_RegionalSettingsCheckBox.SetText( Text );
    m_RegionalSettingsCheckBox.SetState( m_Settings.regionalSettings );

    // New Column
    bounds.left  += columnDX;
    bounds.right  = bounds.left + controlWidth;
    bounds.top    = 45;
    bounds.bottom = bounds.top + LINE_HEIGHT;

    // Windows Network
    m_WindowsNetworkCheckBox.SetBounds( bounds );
    m_WindowsNetworkCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1132_WINDOWS_NETWORK, &Text );
    m_WindowsNetworkCheckBox.SetText( Text );
    m_WindowsNetworkCheckBox.SetState( m_Settings.windowsNetwork );

    // Network TCPIP
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_NetworkTcpIpCheckBox.SetBounds( bounds );
    m_NetworkTcpIpCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1133_NETWORK_TCPIP, &Text );
    m_NetworkTcpIpCheckBox.SetText( Text );
    m_NetworkTcpIpCheckBox.SetState( m_Settings.networkTcpIp );

    // Hardware Devices
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_HardwareDevicesCheckBox.SetBounds( bounds );
    m_HardwareDevicesCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1134_HARDWARE_DEVICES, &Text );
    m_HardwareDevicesCheckBox.SetText( Text );
    m_HardwareDevicesCheckBox.SetState( m_Settings.hardwareDevices );

    // Displays
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_DisplaysCheckBox.SetBounds( bounds );
    m_DisplaysCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1135_DISPLAYS, &Text );
    m_DisplaysCheckBox.SetText( Text );
    m_DisplaysCheckBox.SetState( m_Settings.displays );

    // Display Adapters
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_DisplayAdaptersCheckBox.SetBounds( bounds );
    m_DisplayAdaptersCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1136_DISPLAY_ADAPTERS, &Text );
    m_DisplayAdaptersCheckBox.SetText( Text );
    m_DisplayAdaptersCheckBox.SetState( m_Settings.displayAdapters );

    // Installed Printers
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_InstalledPrintersCheckBox.SetBounds( bounds );
    m_InstalledPrintersCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1137_INSTALLED_PRINTERS, &Text );
    m_InstalledPrintersCheckBox.SetText( Text );
    m_InstalledPrintersCheckBox.SetState( m_Settings.installedPrinters );

    // BIOS Version
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_BiosVersionCheckBox.SetBounds( bounds );
    m_BiosVersionCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1138_BIOS_VERSION, &Text );
    m_BiosVersionCheckBox.SetText( Text );
    m_BiosVersionCheckBox.SetState( m_Settings.biosVersion );

    // System Management
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_SystemManagementCheckBox.SetBounds( bounds );
    m_SystemManagementCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1139_SYSTEM_MANAGEMENT, &Text );
    m_SystemManagementCheckBox.SetText( Text );
    m_SystemManagementCheckBox.SetState( m_Settings.systemManagement );

    // Processors
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_ProcessorsCheckBox.SetBounds( bounds );
    m_ProcessorsCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1140_PROCESSORS, &Text );
    m_ProcessorsCheckBox.SetText( Text );
    m_ProcessorsCheckBox.SetState( m_Settings.processors );

    // Memory
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_MemoryCheckBox.SetBounds( bounds );
    m_MemoryCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1141_MEMORY, &Text );
    m_MemoryCheckBox.SetText( Text );
    m_MemoryCheckBox.SetState( m_Settings.memory );

    // Physical Disks
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_PhysicalDisksCheckBox.SetBounds( bounds );
    m_PhysicalDisksCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1142_PHYSICAL_DISKS, &Text );
    m_PhysicalDisksCheckBox.SetText( Text );
    m_PhysicalDisksCheckBox.SetState( m_Settings.physicalDisks );

    // New Column
    bounds.left  += columnDX;
    bounds.right  = bounds.left + controlWidth;
    bounds.top    = 45;
    bounds.bottom = bounds.top + LINE_HEIGHT;

    // Drives
    m_DrivesCheckBox.SetBounds( bounds );
    m_DrivesCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1143_DRIVES, &Text );
    m_DrivesCheckBox.SetText( Text );
    m_DrivesCheckBox.SetState( m_Settings.drives );

    // Communication Ports
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_CommPortsCheckBox.SetBounds( bounds );
    m_CommPortsCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1144_COMMUNICATION_PORTS, &Text );
    m_CommPortsCheckBox.SetText( Text );
    m_CommPortsCheckBox.SetState( m_Settings.commPorts );

    // Startup Programs
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_StartupProgramsCheckBox.SetBounds( bounds );
    m_StartupProgramsCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1145_STARTUP_PROGRAMMES, &Text );
    m_StartupProgramsCheckBox.SetText( Text );
    m_StartupProgramsCheckBox.SetState( m_Settings.startupPrograms );

    // Services
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_ServicesCheckBox.SetBounds( bounds );
    m_ServicesCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1146_SERVICES, &Text );
    m_ServicesCheckBox.SetText( Text );
    m_ServicesCheckBox.SetState( m_Settings.services );

    // Running Programs
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_RunningProgramsCheckBox.SetBounds( bounds );
    m_RunningProgramsCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1147_RUNNING_PROGRAMMES, &Text );
    m_RunningProgramsCheckBox.SetText( Text );
    m_RunningProgramsCheckBox.SetState( m_Settings.runningPrograms );

    // ODBC Information
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_OdbcInformationsCheckBox.SetBounds( bounds );
    m_OdbcInformationsCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1148_ODBC_INFORMATION, &Text );
    m_OdbcInformationsCheckBox.SetText( Text );
    m_OdbcInformationsCheckBox.SetState( m_Settings.odbcInformation );

    // OLE DB Providers
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_OleDbProvidersCheckBox.SetBounds( bounds );
    m_OleDbProvidersCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1149_OLE_DB_PROVIDERS, &Text );
    m_OleDbProvidersCheckBox.SetText( Text );
    m_OleDbProvidersCheckBox.SetState( m_Settings.oleDbProviders );

    // Software Metering
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_SoftwareMeteringCheckBox.SetBounds( bounds );
    m_SoftwareMeteringCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1150_SOFTWARE_METERING, &Text );
    m_SoftwareMeteringCheckBox.SetText( Text );
    m_SoftwareMeteringCheckBox.SetState( m_Settings.softwareMetering );

    // User Logon Statistics
    OffsetRect( &bounds, 0, LINE_HEIGHT );
    m_UserLogonStatsCheckBox.SetBounds( bounds );
    m_UserLogonStatsCheckBox.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1151_USER_LOGON_STATISTICS, &Text );
    m_UserLogonStatsCheckBox.SetText( Text );
    m_UserLogonStatsCheckBox.SetState( m_Settings.userLogonStats );

    ///////////////////////////////////////////////////////////////////////////
    // Buttons

    // Standard button size
    buttonSize.cx = 80;
    buttonSize.cy = 25;

    location.y = dialogSize.cy - GAP - buttonSize.cy;

    // Horizontal line
    bounds.left   = GAP;
    bounds.right  = dialogSize.cx - GAP;
    bounds.top    = location.y - GAP;
    bounds.bottom = bounds.top + 2;
    Static.SetBounds( bounds );
    Static.SetShape( PXS_SHAPE_SUNK );
    m_Statics.Add( Static );
    Static.Reset();

    // Select None button
    location.x = dialogSize.cx - ( 4 * GAP ) - ( 4 * buttonSize.cx );
    m_SelectNoneButton.SetLocation( location );
    m_SelectNoneButton.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1152_NONE, &Text );
    m_SelectNoneButton.SetText( Text );

    // Select All button
    location.x += ( GAP + buttonSize.cx );
    m_SelectAllButton.SetLocation( location );
    m_SelectAllButton.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1153_ALL, &Text );
    m_SelectAllButton.SetText( Text );

    // Cancel button
    location.x += ( GAP + buttonSize.cx );
    m_CancelButton.SetLocation( location );
    m_CancelButton.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_128_CANCEL, &Text );
    m_CancelButton.SetText( Text );

    // Apply/OK button, set this as the default button. Note, button styles are
    // mutually exclusive so cannot use BM_SETSTYLE, instead store its handle
    location.x += ( GAP + buttonSize.cx );
    m_ApplyButton.SetLocation( location );
    m_ApplyButton.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1154_APPLY, &Text );
    m_ApplyButton.SetText( Text );

    // Mirror for RTL
    if ( rtlReading )
    {
        m_SelectNoneButton.RtlMirror( dialogSize.cx );
        m_SelectAllButton.RtlMirror( dialogSize.cx );
        m_CancelButton.RtlMirror( dialogSize.cx );
        m_ApplyButton.RtlMirror( dialogSize.cx );
    }
    m_bControlsCreated = true;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
