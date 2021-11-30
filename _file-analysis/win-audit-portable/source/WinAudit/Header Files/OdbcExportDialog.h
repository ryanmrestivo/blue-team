///////////////////////////////////////////////////////////////////////////////////////////////////
//
// ODBC Database Export Class Header
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

#ifndef WINAUDIT_ODBC_EXPORT_DIALOG_H_
#define WINAUDIT_ODBC_EXPORT_DIALOG_H_

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
#include "PxsBase/Header Files/ComboBox.h"
#include "PxsBase/Header Files/Dialog.h"
#include "PxsBase/Header Files/ImageButton.h"
#include "PxsBase/Header Files/ProgressBar.h"
#include "PxsBase/Header Files/RadioButton.h"
#include "PxsBase/Header Files/StringT.h"
#include "PxsBase/Header Files/Spinner.h"
#include "PxsBase/Header Files/TextField.h"

// 5. This Project
#include "WinAudit/Header Files/AuditDatabase.h"
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/ConfigurationSettings.h"
#include "WinAudit/Header Files/OdbcDatabase.h"

// 6. Forwards
template< class T > class TArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class OdbcExportDialog: public Dialog
{
    public:
        // Default constructor
        OdbcExportDialog();

        // Destructor
        ~OdbcExportDialog();

        bool DidExport() const;
        void DoExport( const String& ConnectionString );
        void ExportRecordsToDatabase( AuditDatabase* pDatabase, String* pResultMessage );
        void GetConfigurationSettings( ConfigurationSettings* pSettings ) const;
        void SetAuditRecords( const AuditRecord& AuditMasterRecord,
                              const AuditRecord& ComputerMasterRecord,
                              const TArray< AuditRecord >& AuditRecords );
        void SetConfigurationSettings( const ConfigurationSettings& Settings );
    protected:
        // Methods
        LRESULT CommandEvent( WPARAM wParam, LPARAM lParam );
        void    InitDialogEvent();

        // Data members

    private:
        // Holds column properties, see SQLBindCol
        typedef struct _TYPE_COLUMN_PROPS
        {
            SQLSMALLINT  TargetType;
            SQLSMALLINT  sqlType;
            SQLLEN       BufferLength;
            BYTE*        TargetValuePtr;
            SQLLEN*      StrLen_or_Ind;
            size_t       numBytes;
        } TYPE_COLUMN_PROPS;

        // Constants
        static const size_t   MAX_COL_SIZE_CHARS = 255;  // String column

        // Copy constructor - not allowed
        OdbcExportDialog( const OdbcExportDialog& oAuditRecord );

        // Assignment operator - not allowed
        OdbcExportDialog& operator= ( const OdbcExportDialog& oAuditRecord );

        // Methods
        void    AllocateAuditDataBindBuffers( AuditDatabase* pDatabase );
        void    ExportRecords();
        void    FillAuditDataBindBuffers( SQLINTEGER auditID,
                                          SQLINTEGER computerID );
        void    FreeAuditDataBindBuffers();
        void    SetProgressMessage( const String& ProgressMessage );
        void    ShowAdminDialog();
        void    UpdateConfigurationSettings();

        // Data members
        bool                  m_bDidExport;
        size_t                m_idxProgressMessageStatic;
        size_t                m_uNumColumns;
        TYPE_COLUMN_PROPS*    m_pColumnProps;
        SQLUSMALLINT*         m_pRowStatus;
        AuditRecord           m_AuditMasterRecord;
        AuditRecord           m_ComputerMasterRecord;
        TArray< AuditRecord > m_AuditRecords;
        ConfigurationSettings m_Settings;
        Spinner               m_MaxErrorRateSpinner;
        Spinner               m_MaxAffectedRowsSpinner;
        Spinner               m_ConnectTimeoutSpinner;
        Spinner               m_QueryTimeoutSpinner;
        Button                m_CloseButton;
        Button                m_BrowseButton;
        Button                m_AdminButton;
        Button                m_ExportButton;
        TextField             m_DatabaseNameTextField;
        ComboBox              m_MySQLComboBox;
        ComboBox              m_PostgreComboBox;
        RadioButton           m_AccessOption;
        RadioButton           m_SqlServerOption;
        RadioButton           m_MySQLOption;
        RadioButton           m_PostgreSQLOption;
        ProgressBar           m_ProgressBar;
};

#endif  // WINAUDIT_ODBC_EXPORT_DIALOG_H_
