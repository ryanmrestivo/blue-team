///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Database Administration Dialog Class Header
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

#ifndef WINAUDIT_DATABASE_ADMINISTRATION_DIALOG_H_
#define WINAUDIT_DATABASE_ADMINISTRATION_DIALOG_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

// In order for the user not to have to login for every action, the database
// connection is established with a class scope object. This connection is
// closed when this dialog is closed.

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
#include "PxsBase/Header Files/ComboBox.h"
#include "PxsBase/Header Files/Dialog.h"
#include "PxsBase/Header Files/ProgressBar.h"
#include "PxsBase/Header Files/Spinner.h"
#include "PxsBase/Header Files/StringT.h"

// 5. This Project
#include "WinAudit/Header Files/AuditDatabase.h"
#include "WinAudit/Header Files/ConfigurationSettings.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class DatabaseAdministrationDialog: public Dialog
{
    public:
        // Default constructor
        DatabaseAdministrationDialog();

        // Destructor
        ~DatabaseAdministrationDialog();

        void    GetConfigurationSettings( ConfigurationSettings* pSettings ) const;
        void    SetConfigurationSettings( const ConfigurationSettings& Settings );

    protected:
        // Methods
        LRESULT CommandEvent( WPARAM wParam, LPARAM lParam );
        void    InitDialogEvent();

        // Data members

    private:
        // Copy constructor - not allowed
        DatabaseAdministrationDialog(
                                 const DatabaseAdministrationDialog& oDialog );

        // Assignment operator - not allowed
        DatabaseAdministrationDialog& operator= ( const DatabaseAdministrationDialog& oDialog );

        // Methods
        void    AddCreateAccessProcedures( StringArray* pStatements );
        void    AddCreateAuditDataSql( const String& IntegerKeyword,
                                       const String& TCharKeyword,
                                       const String& VarTCharKeyword,
                                       StringArray* pStatements );
        void    AddCreateAuditMasterSql( const String& AutoIncrementKeyword,
                                         const String& IntegerKeyword,
                                         const String& TimestampKeyword,
                                         const String& VarCharKeyword,
                                         const String& VarTCharKeyword,
                                         StringArray* pStatements );
        void    AddCreateComputerMasterSql( const String& AutoIncrementKeyword,
                                            const String& IntegerKeyword,
                                            const String& TimestampKeyword,
                                            const String& VarCharKeyword,
                                            const String& VarTCharKeyword,
                                            StringArray* pStatements );
        void    AddCreateDisplayNamesSql( const String& IntegerKeyword,
                                           const String& VarTCharKeyword,
                                           StringArray* pStatements );
        void    AddCreateGrants( StringArray* pStatements );
        void    AddCreateMySqlProcedures( StringArray* pStatements );
        void    AddCreateSqlServerProcedures( StringArray* pStatements );
        void    AddCreateViews( const String& SchemaName, StringArray* pStatements );
        void    ConnectDB( bool serverOnly );
        void    CreateAccessDatabase();
        void    CreateDatabase();
        void    CreateServerDatabase();
        void    DeleteOldAudits();
        void    GetDatabaseSpecificKeywords( String* pAutoIncrementKeyword,
                                             String* pCharKeyword,
                                             String* pIntegerKeyword,
                                             String* pSmallIntKeyword,
                                             String* pTimestampKeyword,
                                             String* pVarCharKeyword,
                                             String* pWCharKeyword,
                                             String* pWVarCharKeyword );
        void    MakeReportSql( const String& UpperKeyword, String* pSqlQuery );
        void    RunReport();
        void    UpdateSettings();
        void    SetProgressMessage( const String& ProgressMessage );

        // Data members
        size_t                  m_uIdxStaticProgress;
        String                  m_OdbcDriver;
        Button                  m_CloseButton;
        Button                  m_CreateButton;
        Button                  m_DeleteHistoryButton;
        Button                  m_RunReportButton;
        CheckBox                m_GrantPublicCheckBox;
        CheckBox                m_ReportShowSqlCheckBox;
        CheckBox                m_UnicodeCheckBox;
        ComboBox                m_AccessComboBox;
        ComboBox                m_ReportNamesComboBox;
        ProgressBar             m_ProgressBar;
        Spinner                 m_MaxReportRowsSpinner;
        AuditDatabase           m_AuditDatabase;
        ConfigurationSettings   m_Settings;
};

#endif  // WINAUDIT_DATABASE_ADMINISTRATION_DIALOG_H_
