///////////////////////////////////////////////////////////////////////////////////////////////////
//
// System Information Class Header
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

#ifndef PXSBASE_SYSTEM_INFORMATION_H_
#define PXSBASE_SYSTEM_INFORMATION_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/PxsBase.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

// 6. Forward

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class SystemInformation
{
    public:
        // Default constructor
        SystemInformation();

        // Copy constructor
        SystemInformation( const SystemInformation& oSystemInfo );

        // Destructor
        ~SystemInformation();

        // Assignment operator
        SystemInformation& operator= ( const SystemInformation& oSystemInfo );

        // Methods
        void    DoHeapCheck( String* pHeapCheck );
        DWORD   GetBuildNumber() const;
        void    GetBuiltInAccountNameFromRID( DWORD RID, String* pAccountName ) const;
        void    GetComputerDescription( String* pDescription ) const;
        void    GetComputerNetBiosName( String* pNetBiosName ) const;
        void    GetComputerRoles( String* pComputerRoles ) const;
        void    GetComputerSiteName( String* pSiteName ) const;
        void    GetCurrentUserName( String* pCurrentUserName ) const;
        void    GetDomainNetBiosName( String* pDomainNetBiosName ) const;
        void    GetEdition( String* pEdition ) const;
        void    GetFullyQualifiedDomainName( String* pFQDN ) const;
        void    GetMajorMinorBuild( String* pMajorMinorBuild ) const;
        DWORD   GetMajorVersion() const;
        DWORD   GetMinorVersion() const;
        DWORD   GetNumberOperatingSystemBits() const;
        DWORD   GetProductInfoType() const;
        void    GetServicePackMajorDotMinor( String* pMajorDotMinor ) const;
        void    GetSessionName( String* pSessionName ) const;
        void    GetSystemDirectoryPath( String* pDirectoryPath ) const;
        void    GetSystemLibraryVersion( LPCWSTR pszLibName,
                                         String* pLibraryVersion ) const;
        time_t  GetUpTime() const;
        BYTE    GetVersionProductType() const;
        void    GetVersionStrings( String* pNetFunction,
                                   String* pWmiWin32,
                                   String* pKernel32, String* pRegistryVer ) const;
        WORD    GetVersionSuiteMask() const;
        void    GetWindowsDirectoryPath( String* pDirectoryPath ) const;
        bool    Is64BitProcessor() const;
        bool    IsAccountEnabled( const String& UserName ) const;
        bool    IsAdministrator() const;
        bool    IsRemoteControl() const;
        bool    IsRemoteSession() const;
        bool    IsPreInstall() const;
        bool    IsWOW64() const;

    protected:
        // Methods
        void FillVersionInfoEx( OSVERSIONINFOEX* pVersionInfoEx ) const;

        // Data members

    private:
        // Methods
        void TranslateXPEdition( WORD suiteMask, String* pEdition ) const;

        // Data members
};

#endif  // PXSBASE_SYSTEM_INFORMATION_H_
