///////////////////////////////////////////////////////////////////////////////////////////////////
//
// CPU Information Class Header
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

#ifndef WINAUDIT_CPU_INFORMATION_H_
#define WINAUDIT_CPU_INFORMATION_H_

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
class String;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class CpuInformation
{
    public:
        // Default constructor
        CpuInformation();

        // Destructor
        ~CpuInformation();

        // Methods
        void    GetAuditRecords( TArray< AuditRecord >* pRecords );
        void    GetBasicDataRecord( BYTE packageID, AuditRecord* pRecord );
        void    GetCacheData( BYTE packageID, TArray< AuditRecord >* pRecords );
        void    GetDiagnostics( String* pDiagnostics );
        void    GetFeatures( BYTE packageID, StringArray* pFeatures );
        void    GetNameAndSpeedMHz( BYTE packageID, String* pName, DWORD* pSpeedMHz );
        void    GetPackageIDs( TArray< BYTE >* pPackageIDs );

    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        CpuInformation( const CpuInformation& oCpuInformation );

        // Assignment operator - not allowed
        CpuInformation& operator= ( const CpuInformation& oCpuInformation );

        // Methods
        void      Cpuid( DWORD  function, DWORD* pEAX, DWORD* pEBX, DWORD* pECX, DWORD* pEDX );
        void      FillLogicalApics();
        BYTE      GetApicID();
        void      GetApicsIDs( BYTE packageID, String* pApicsIDs );
        BYTE      GetBrandID();
        void      GetCacheAndTLBInfo( String* pCacheInfo, String* pTLBInfo );
        void      GetCacheDescriptionsAmd( StringArray* pCacheDescriptionsAmd);
        void      GetCacheDescriptionsIntel( StringArray* pCacheDescriptionsIntel );
        void      GetFeaturesAmd( StringArray* pFeaturesAmd );
        void      GetFeaturesIntel( StringArray* pFeaturesIntel );
        DWORD     GetManufacturerID();
        void      GetManufacturerName( String* pManufacturerName );
        BYTE      GetMaximumLogicalsInPackage();
        void      GetName( String* pName );
        BYTE      GetNumberCoresInPackage();
        BYTE      GetNumberLogicalsInPackage( BYTE packageID );
        BYTE      GetPackageIDFromApicID( BYTE apicID );
        void      GetRegistersAsString( bool standard, String* pRegisters );
        void      GetSignature( BYTE* pStepping,
                                BYTE* pModel,
                                BYTE* pFamily, BYTE* pType, BYTE* pModelExt, BYTE* pFamilyExt );
        DWORD     GetSpeedEstimateMHz();
        DWORD     GetSpeedRegistryMHz( BYTE packageID );
        DWORD     GetStandardFeatureInformation();
        void      GetType( String* pType );
        void      GetVendorString( String* pVendorString );
        DWORD_PTR RunCurrentThreadOnPackageID( BYTE packageID );
        bool      SupportsMultiCore();
        void      TranslateProcessorArchitecture( WORD processorArchitecture,
                                                  String* pTranslation );
        // Data members
        TArray< BYTE > m_LogicalApics;
};

#endif  // WINAUDIT_CPU_INFORMATION_H_
