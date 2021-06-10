///////////////////////////////////////////////////////////////////////////////////////////////////
//
// WinAudit Configuration Dialog Class Header
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

#ifndef WINAUDIT_WINAUDIT_CONFIG_DIALOG_H_
#define WINAUDIT_WINAUDIT_CONFIG_DIALOG_H_

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
#include "PxsBase/Header Files/Button.h"
#include "PxsBase/Header Files/CheckBox.h"
#include "PxsBase/Header Files/Dialog.h"

// 5. This Project
#include "WinAudit/Header Files/ConfigurationSettings.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class WinAuditConfigDialog : public Dialog
{
    public:
        // Default constructor
        WinAuditConfigDialog();

        // Destructor
        ~WinAuditConfigDialog();

        // Methods
        void GetConfigurationSettings( ConfigurationSettings* pSettings ) const;
        void SetConfigurationSettings( const ConfigurationSettings& Settings );

    protected:
        // Methods
        LRESULT CommandEvent( WPARAM wParam, LPARAM lParam );
        void    InitDialogEvent();

        // Data members

    private:
        // Copy constructor - not allowed
        WinAuditConfigDialog( const WinAuditConfigDialog& oDialog );

        // Assignment operator - not allowed
        WinAuditConfigDialog& operator= ( const WinAuditConfigDialog& oDialog );

        // Methods

        // Data members
        ConfigurationSettings m_Settings;

        CheckBox    m_SystemOverviewCheckBox;
        CheckBox    m_InstalledSoftwareCheckBox;
        CheckBox    m_OperatingSystemCheckBox;
        CheckBox    m_PeripheralsCheckBox;
        CheckBox    m_SecurityCheckBox;
        CheckBox    m_GroupsAndUsersCheckBox;
        CheckBox    m_ScheduledTasksCheckBox;
        CheckBox    m_UptimeStatisticsCheckBox;
        CheckBox    m_ErrorLogsCheckBox;
        CheckBox    m_EnvironmentVarsCheckBox;
        CheckBox    m_RegionalSettingsCheckBox;
        CheckBox    m_WindowsNetworkCheckBox;
        CheckBox    m_NetworkTcpIpCheckBox;
        CheckBox    m_HardwareDevicesCheckBox;
        CheckBox    m_DisplaysCheckBox;
        CheckBox    m_DisplayAdaptersCheckBox;
        CheckBox    m_InstalledPrintersCheckBox;
        CheckBox    m_BiosVersionCheckBox;
        CheckBox    m_SystemManagementCheckBox;
        CheckBox    m_ProcessorsCheckBox;
        CheckBox    m_MemoryCheckBox;
        CheckBox    m_PhysicalDisksCheckBox;
        CheckBox    m_DrivesCheckBox;
        CheckBox    m_CommPortsCheckBox;
        CheckBox    m_StartupProgramsCheckBox;
        CheckBox    m_ServicesCheckBox;
        CheckBox    m_RunningProgramsCheckBox;
        CheckBox    m_OdbcInformationsCheckBox;
        CheckBox    m_OleDbProvidersCheckBox;
        CheckBox    m_SoftwareMeteringCheckBox;
        CheckBox    m_UserLogonStatsCheckBox;
        Button      m_SelectNoneButton;
        Button      m_SelectAllButton;
        Button      m_CancelButton;
        Button      m_ApplyButton;
};

#endif  // WINAUDIT_WINAUDIT_CONFIG_DIALOG_H_
