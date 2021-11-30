///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Access Database Class Header
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

#ifndef WINAUDIT_ACCESS_DATABASE_H_
#define WINAUDIT_ACCESS_DATABASE_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/WinAudit.h"

// 2. C System Files
#include <oledb.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/BStr.h"
#include "PxsBase/Header Files/StringT.h"

// 5. This Project

// 6. Forwards
class StringArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class AccessDatabase
{
    public:
        // Default constructor
        AccessDatabase();

        // Destructor
        ~AccessDatabase();

        // Methods
        void    ComCloseCurrentDatabase();
        void    ComDoCmd_RunCommand( long command );
        void    ComQuit( long option );
        void    ComStart();
        void    ComVisible( bool visible );
        void    OleDbCreate( const String& DatabasePath );
        void    OleDbConnect( const String& DatabsePath,
                              const String& UserName, const String& Password );
        void    OleDbDisconnect();
        void    OleDbExecute( const String& SqlQuery );
        bool    OleDbIsConnected() const;
        void    OleDbSetProvider( const String& ProviderID );

    protected:
        // Methods

        // Methods

    private:
        // Copy constructor - not allowed
        AccessDatabase( const AccessDatabase& oAccessDatabase );

        // Assignment operator - not allowed
        AccessDatabase& operator= ( const AccessDatabase& oAccessDatabase );

        // Methods
        void    ComInvoke( IDispatch* pDispatch,
                           LPCWSTR pwzName,
                           WORD flags, DISPPARAMS* pDispParams, VARIANT* pResult );
        void    ComMethod( IDispatch* pDispatch,
                           LPCWSTR pwzName, DISPPARAMS* pDispParams, VARIANT* pResult );
        void    ComProperyGet( IDispatch* pDispatch,
                               LPCWSTR pwzName, DISPPARAMS* pDispParams, VARIANT* pResult );
        void    ComProperyPut( IDispatch* pDispatch,
                               LPCWSTR pwzName, DISPPARAMS* pDispParams, VARIANT* pResult );
        void    CreateIDBInitialize();
        void    GetAccessClassID( CLSID* pAccessClassID );
        void    GetOleDbErrorString( String* pOleDbErrorString );
        void    MakeOleDbErrorString( ULONG index,
                                      IErrorInfo* pIErrorInfo,
                                      ISQLErrorInfo* pISQLErrorInfo, String* pOleDbErrorString );
        // Data members
        BStr            m_DatabasePath;
        BStr            m_UserName;
        BStr            m_Password;
        String          m_ProviderID;          // aka ProgID

        // Data members - OLE DB
        bool            m_bConnected;
        IDBInitialize*  m_pIDBInitialize;

        // Data mmbers - COM
        bool            m_bDatabaseOpen;
        long            acCmdNewObjectModule;   // Access constant = 139
        long            acModule;               // Access constant = 5
        long            acSaveYes;              // Access constant = 1
        IUnknown*       m_pIUnknown;
        IDispatch*      m_pDispAccess;
};

#endif  // WINAUDIT_ACCESS_DATABASE_H_
