///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Software Information Class Header
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

#ifndef WINAUDIT_SOFTWARE_INFORMATION_H_
#define WINAUDIT_SOFTWARE_INFORMATION_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/WinAudit.h"

// 2. C System Files
#include <Msi.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Registry.h"

// 5. This Project

// 6. Forwards
class AuditRecord;
template< class T > class TArray;
template< class T > class TList;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class SoftwareInformation
{
    public:
        // Default constructor
        SoftwareInformation();

        // Destructor
        ~SoftwareInformation();

        // Methods
        void    GetActiveSetupRecords( TArray< AuditRecord >* pRecords );
        void    GetDiagnostics( String* pDiagnostics );
        void    GetInstalledSoftwareRecords( TArray< AuditRecord >* pRecords );
        void    GetSoftwareUpdateRecords( TArray< AuditRecord >* pRecords );
        void    GetStartupProgramRecords( TArray< AuditRecord >* pRecords );

    protected:
        // Methods

        // Data members

    private:
        // Structure to hold the binary data in the registry
        // Based on SLOWAPPINFO
        typedef struct _TYPE_SLOWINFOCACHE
        {
            DWORD       cbSize;
            DWORD       dwHasName;
            ULONGLONG   ullSize;
            FILETIME    ftLastUsed;
            int         iTimesUsed;
            wchar_t     wImage[ MAX_PATH ];  // not terminated
        } TYPE_SLOWINFOCACHE;

        // Structure to hold update data
        typedef struct _TYPE_UPDATE_DATA
        {
            wchar_t szUpdateID[ 48 ];       // Up to  GUID
            wchar_t szInstalledOn[ 24 ];    // Enough for an ISO timestamp
            wchar_t szDescription[ 256 ];
        } TYPE_UPDATE_DATA;

        // Structure to hold installed software data
        typedef struct _TYPE_INSTALLED_DATA
        {
            int     timesUsed;
            wchar_t szSoftwareKey[ 256 ];       // Registry key, GUID or name
            wchar_t szName[ 256 ];
            wchar_t szVendor[ 256 ];            // aka Publisher
            wchar_t szVersion[ 256 ];
            wchar_t szLanguage[ 256 ];
            wchar_t szInstallDate[ 24 ];
            wchar_t szInstallLocation[ 256 ];
            wchar_t szInstallSource[ 256 ];
            wchar_t szInstallState[ 256 ];
            wchar_t szAssignmentType[ 16 ];     // 'Per User' or 'Per Computer'
            wchar_t szPackageCode[ 256 ];
            wchar_t szPackageName[ 256 ];
            wchar_t szLocalPackage[ 256 ];
            wchar_t szProductID[ 48 ];          // GUID
            wchar_t szRegCompany[ 256 ];
            wchar_t szRegOwner[ 256 ];
            wchar_t szLastUsed[ 32 ];           // File time in user locale
            wchar_t szExePath[ 256 ];
            wchar_t szExeVersion[ 256 ];
            wchar_t szExeDescription[ 256 ];
        } TYPE_INSTALLED_DATA;

        // Copy constructor - not allowed
        SoftwareInformation( const SoftwareInformation& oSoftware );

        // Assignment operator - not allowed
        SoftwareInformation& operator= ( const SoftwareInformation& oSoftware);

        // Methods
        bool  AddSoftwareInfoFromRegistry( Registry* pRegObject,
                                           LPCWSTR pszUnistallPath,
                                           const String& KeyName,
                                           TList<TYPE_INSTALLED_DATA>* pInstalled );
        void  BackFillMissingData( const TYPE_UPDATE_DATA* pUpdate1,
                                   TYPE_UPDATE_DATA* pUpdate2 );
        bool  FindInInstalledData( const String& Name,
                                   TList<TYPE_INSTALLED_DATA>* pInstalled );
        DWORD GetArpCacheData( const String& KeyName,
                               int* pTimesUsed, String* pLastUsed, String* pExePath );
        void  GetInstalledFromMsi( TList<TYPE_INSTALLED_DATA>* pInstalled );
        void  GetInstalledFromRegistry( TList<TYPE_INSTALLED_DATA>* pInstalled);
        void  GetInstalledMsoFromRegistry( TList<TYPE_INSTALLED_DATA>* pInstalled);
        void  GetUpdatesFromRegistry( TList<TYPE_UPDATE_DATA>* pUpdates );
        void  GetUpdatesFromMsi( TList<TYPE_UPDATE_DATA>* pUpdates );
        void  GetUpdatesFromWmi( TList<TYPE_UPDATE_DATA>* pUpdates );
        void  GetUpdatesFromWua( TList<TYPE_UPDATE_DATA>* pUpdates );
        bool  FindInstalledSoftwareElement( const TYPE_INSTALLED_DATA* pData,
                                            TList<TYPE_INSTALLED_DATA>* pInstalled,
                                            TYPE_INSTALLED_DATA* pElement );
        bool  IsValidWmiUpdateInstalledOn( const String& InstalledOn ) const;
        bool  IsSameUpdate( const TYPE_UPDATE_DATA* pUpdate1, const TYPE_UPDATE_DATA* pUpdate2 );
        void  MergeInstalledSoftwareLists( TList<TYPE_INSTALLED_DATA>* pMsiInstalled,
                                           TList<TYPE_INSTALLED_DATA>* pRegInstalled );
        void  MergeSoftwareUpdatesLists( TList<TYPE_UPDATE_DATA>* pList,
                                         TList<TYPE_UPDATE_DATA>* pUpdatesList);
        bool  IsQorKbNumber( LPCWSTR pszString );
        void  TranslateMsiInstallState( INSTALLSTATE state, String* pInstallState );

        // Data members
};

#endif  // WINAUDIT_SOFTWARE_INFORMATION_H_
