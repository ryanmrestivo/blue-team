///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Process Information Class Header
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

#ifndef WINAUDIT_PROCESS_INFORMATION_H_
#define WINAUDIT_PROCESS_INFORMATION_H_

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
#include "PxsBase/Header Files/TArray.h"

// 5. This Project

// 6. Forwards
class AuditRecord;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class ProcessInformation
{
    public:
        // Default constructor
        ProcessInformation();

        // Destructor
        ~ProcessInformation();

        // Methods
        void    ChangePrivilege( LPCWSTR pszPrivilege, bool enable );
        void    FillProcessesInfo();
        void    GetExePath( DWORD processID, String* pExePath );
        void    GetLoadedModuleRecords( TArray< AuditRecord >* pRecords );
        void    GetRunningProcesseRecords( TArray< AuditRecord >* pRecords );
        bool    IsPrivilegeEnabled( LPCWSTR pszPrivilege );

    protected:
        // Methods

        // Data members

    private:
        // Structure to hold basic process information
        typedef struct _TYPE_PROCESS_INFO
        {
            DWORD   processID;                  // Process ID
            size_t  workingSet;                 // Working Set Memory - WinNT
            wchar_t szExeName[ MAX_PATH + 1 ];  // Executable file name
            wchar_t szExePath[ MAX_PATH + 1 ];  // Executable file path
        } TYPE_PROCESS_INFO;

        // Copy constructor - not allowed
        ProcessInformation( const ProcessInformation& oProcess );

        // Assignment operator - not allowed
        ProcessInformation& operator= ( const ProcessInformation& oProcess );

        // Methods
        size_t  GetMemoryUsage( DWORD processID );
        size_t  GetWorkingSetSize( DWORD processID );

        // Data members
        TArray< TYPE_PROCESS_INFO > m_Processes;
};

#endif  // WINAUDIT_PROCESS_INFORMATION_H_
