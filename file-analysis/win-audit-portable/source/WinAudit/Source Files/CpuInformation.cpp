///////////////////////////////////////////////////////////////////////////////////////////////////
//
// CPU Information Class Implementation
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

// Intel ad AMD processors only. The assumption is made that all logical
// processors in a physical processor aka "package" are the same design.
//
// Reference docs:
// AMD64 Architecture Programmer's Manual Volume 3: General-Purpose and
// System Instructions and Intel(R) 64 and IA-32 Architectures Software
// Developer's Manual, May 2013, #24594
//
// Intel(R) 64 and IA-32 Architectures Software Developer's Manual Volume 2A:
// Instruction Set Reference, A-M, 253666-047US, June 2013

///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/CpuInformation.h"

// 2. C System Files
#pragma warning( push )
#pragma warning ( disable : 4995 )  // name was marked as #pragma deprecated
#include <intrin.h>
#pragma warning( pop )

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/Registry.h"
#include "PxsBase/Header Files/SystemException.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
CpuInformation::CpuInformation()
               :m_LogicalApics()

{
}

// Copy constructor - not allowed so no implementation

// Destructor
CpuInformation::~CpuInformation()
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
//      Get basic data about the specified processor package as an array of
//      audit records
//
//  Parameters:
//      pRecords - receives the data
//
//  Returns:
//      void
//===============================================================================================//
void CpuInformation::GetAuditRecords( TArray< AuditRecord >* pRecords )
{
    BYTE   packageID;
    size_t i, numPackages;
    AuditRecord Record;
    TArray< BYTE > PackagesIDs;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    GetPackageIDs( &PackagesIDs );
    numPackages = PackagesIDs.GetSize();
    for ( i = 0; i < numPackages; i++ )
    {
        packageID = PackagesIDs.Get( i );
        GetBasicDataRecord( packageID, &Record );
        pRecords->Add( Record );
    }
}

//===============================================================================================//
//  Description:
//      Get basic data about the specified processor package as an audit
//      record
//
//  Parameters:
//      packageID - the processor "package" id
//      pRecord   - object to receive the data
//
//  Returns:
//      void
//===============================================================================================//
void CpuInformation::GetBasicDataRecord( BYTE packageID, AuditRecord* pRecord )
{
    BYTE   numLogicals = 0;
    DWORD  speedMHz = 0;
    size_t i = 0, numFeatures = 0;
    String Name, Type, ManufacturerName, ApicIDs, Value, CacheInfo, TLBInfo;
    String LocaleMHz;
    Formatter   Format;
    StringArray Features;
    DWORD_PTR oldThreadAffinityMask = 0;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_CPU_BASIC );

    // If have not got the APIC identifiers do so now
    if ( m_LogicalApics.GetSize() == 0 )
    {
        FillLogicalApics();
    }

    PXSGetResourceString( PXS_IDS_1263_MHZ, &LocaleMHz );
    oldThreadAffinityMask = RunCurrentThreadOnPackageID( packageID );

    // Need to reset the thread's mask
    try
    {
        Value.Allocate( 256 );

        // Physical Processor Number, counting starts at 1
        Value = Format.Int32( packageID + 1 );
        pRecord->Add( PXS_CPU_BASIC_ITEM_NUMBER, Value );

        // Processor Name
        GetName( &Name );
        pRecord->Add( PXS_CPU_BASIC_FULL_NAME, Name );

        // Short Name - no longer used
        pRecord->Add( PXS_CPU_BASIC_SHORT_NAME, PXS_STRING_EMPTY );

        // Estimated speed
        Value    = PXS_STRING_EMPTY;
        speedMHz = GetSpeedEstimateMHz();
        if ( speedMHz > 0 )
        {
            Value  = Format.UInt32( speedMHz );
            Value += LocaleMHz;
        }
        pRecord->Add( PXS_CPU_BASIC_SPEED_ESTIMATE_MHZ, Value );

        // Speed read from registry
        Value    = PXS_STRING_EMPTY;
        speedMHz = GetSpeedRegistryMHz( packageID );
        if ( speedMHz > 0 )
        {
            Value  = Format.UInt32( speedMHz );
            Value += LocaleMHz;
        }
        pRecord->Add( PXS_CPU_BASIC_SPEED_REGISTRY_MHZ, Value );

        GetType( &Type );
        pRecord->Add( PXS_CPU_BASIC_TYPE, Type );

        GetManufacturerName( &ManufacturerName );
        pRecord->Add( PXS_CPU_BASIC_MANUFACTURER, ManufacturerName );

        GetApicsIDs( packageID, &ApicIDs );
        pRecord->Add( PXS_CPU_BASIC_APIC_PHYSICAL_ID, ApicIDs );

        // Logicals
        Value = PXS_STRING_ONE;
        numLogicals = GetNumberLogicalsInPackage( packageID );
        if ( numLogicals )
        {
            Value = Format.UInt8( numLogicals );
        }
        pRecord->Add( PXS_CPU_BASIC_NUM_LOGICALS, Value );

        // Cache String
        GetCacheAndTLBInfo( &CacheInfo, &TLBInfo );
        pRecord->Add( PXS_CPU_BASIC_CACHE, CacheInfo );
        pRecord->Add( PXS_CPU_BASIC_TLB, TLBInfo );

        // Features
        Value = PXS_STRING_EMPTY;
        GetFeatures( packageID, &Features );
        numFeatures = Features.GetSize();
        for ( i = 0; i < numFeatures; i++ )
        {
            if ( i )
            {
                Value += L", ";
            }
            Value += Features.Get( i );
        }
        pRecord->Add( PXS_CPU_BASIC_FEATURES, Value );
    }
    catch ( const Exception& )
    {
        SetThreadAffinityMask( GetCurrentThread(), oldThreadAffinityMask );
        throw;
    }
    SetThreadAffinityMask( GetCurrentThread(), oldThreadAffinityMask );
}

//===============================================================================================//
//  Description:
//      Get diagnostic data for the CPU on the system
//
//  Parameters:
//      pDiagnostics - string object to receive the CPUID data
//
//  Returns:
//      void
//===============================================================================================//
void CpuInformation::GetDiagnostics( String* pDiagnostics )
{
    size_t      numPackages = 0;
    BYTE        packageID = 0, byteValue = 0, stepping = 0, model = 0;
    BYTE        family = 0, type = 0, modelExt = 0, familyExt = 0;
    DWORD       i = 0, maxStdFunc = 0, maxExtFunc = 0, speedMhz = 0, EAX = 0;
    HANDLE      hProcess  = nullptr, hThread = nullptr;
    String      Title, Vendor, Value, ApplicationName, Registers, DataString;
    Formatter   Format;
    DWORD_PTR   ProcessAffinityMask = 0, SystemAffinityMask = 0;
    DWORD_PTR   ThreadMask = 0;
    SYSTEM_INFO si;
    TArray< BYTE > PackageIDs;

    if ( pDiagnostics == nullptr )
    {
        throw ParameterException( L"pDiagnostics", __FUNCTION__ );
    }
    *pDiagnostics = PXS_STRING_EMPTY;

    DataString.Allocate( 4096 );
    PXSGetApplicationName( &ApplicationName );
    Title  = L"Processor Data by ";
    Title += ApplicationName;
    DataString += Title;
    DataString += PXS_STRING_CRLF;
    DataString.AppendChar( '=', Title.GetLength() );
    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;

    ////////////////////////////////////////////////////////////////////////////
    // System Information

    memset( &si, 0, sizeof ( si ) );
    GetNativeSystemInfo( &si );

    // Add affinity masks
    DataString += L"SYSTEM INFORMATION\r\n";
    DataString += L"==================\r\n\r\n";

    // Processor Architecture
    Value = PXS_STRING_EMPTY;
    TranslateProcessorArchitecture( si.wProcessorArchitecture, &Value );
    DataString += L"Processor Architecture: ";
    DataString += Value;
    DataString += PXS_STRING_CRLF;

    DataString += L"Active Processor Mask : ";
    DataString += Format.BinarySizeT( si.dwActiveProcessorMask );
    DataString += PXS_STRING_CRLF;

    DataString += L"Number Processors     : ";
    DataString += Format.UInt32( si.dwNumberOfProcessors );
    DataString += PXS_STRING_CRLF;

    DataString += L"Processor Revision    : ";
    DataString += Format.UInt16( si.wProcessorRevision );
    DataString += PXS_STRING_CRLF;

    DataString += PXS_STRING_CRLF;

    // IsProcessorFeaturePresent available on NT4 and above
    // PF_FLOATING_POINT_PRECISION_ERRATA
    DataString += L"PF_FLOATING_POINT_PRECISION_ERRATA: ";
    DataString += Format.Int32YesNo( IsProcessorFeaturePresent( 0 ) );
    DataString += PXS_STRING_CRLF;

    // PF_FLOATING_POINT_EMULATED
    DataString += L"PF_FLOATING_POINT_EMULATED        : ";
    DataString += Format.Int32YesNo( IsProcessorFeaturePresent( 1 ) );
    DataString += PXS_STRING_CRLF;

    // PF_COMPARE_EXCHANGE_DOUBLE
    DataString += L"PF_COMPARE_EXCHANGE_DOUBLE        : ";
    DataString += Format.Int32YesNo( IsProcessorFeaturePresent( 2 ) );
    DataString += PXS_STRING_CRLF;

    // PF_MMX_INSTRUCTIONS_AVAILABLE
    DataString += L"PF_MMX_INSTRUCTIONS_AVAILABLE     : ";
    DataString += Format.Int32YesNo( IsProcessorFeaturePresent( 3 ) );
    DataString += PXS_STRING_CRLF;

    // PF_XMMI_INSTRUCTIONS_AVAILABLE
    DataString += L"PF_XMMI_INSTRUCTIONS_AVAILABLE    : ";
    DataString += Format.Int32YesNo( IsProcessorFeaturePresent( 6 ) );
    DataString += PXS_STRING_CRLF;

    // PF_3DNOW_INSTRUCTIONS_AVAILABLE
    DataString += L"PF_3DNOW_INSTRUCTIONS_AVAILABLE   : ";
    DataString += Format.Int32YesNo( IsProcessorFeaturePresent( 7 ) );
    DataString += PXS_STRING_CRLF;

    // PF_RDTSC_INSTRUCTION_AVAILABLE
    DataString += L"PF_RDTSC_INSTRUCTION_AVAILABLE    : ";
    DataString += Format.Int32YesNo( IsProcessorFeaturePresent( 8 ) );
    DataString += PXS_STRING_CRLF;

    // PF_PAE_ENABLED
    DataString += L"PF_PAE_ENABLED                    : ";
    DataString += Format.Int32YesNo( IsProcessorFeaturePresent( 9 ) );
    DataString += PXS_STRING_CRLF;

    // PF_XMMI64_INSTRUCTIONS_AVAILABLE
    DataString += L"PF_XMMI64_INSTRUCTIONS_AVAILABLE  : ";
    DataString += Format.Int32YesNo( IsProcessorFeaturePresent( 10 ) );
    DataString += PXS_STRING_CRLF;

    // PF_NX_ENABLED
    DataString += L"PF_NX_ENABLED                     : ";
    DataString += Format.Int32YesNo( IsProcessorFeaturePresent( 12 ) );
    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;

    ////////////////////////////////////////////////////////////////////////////
    // Add affinity masks

    DataString += PXS_STRING_CRLF;
    DataString += L"AFFINITY MASKS\r\n";
    DataString += L"==============\r\n\r\n";

    // Get the process handle
    hProcess = GetCurrentProcess();
    hThread  = GetCurrentThread();

    // Add the affinity masks
    if ( GetProcessAffinityMask( hProcess, &ProcessAffinityMask, &SystemAffinityMask ) == 0 )
    {
        DataString += L"GetProcessAffinityMask failed, error code = ";
        DataString += Format.UInt32( GetLastError() );
        DataString += PXS_STRING_CRLF;
        return;
    }

    DataString += L"Process affinity mask   : ";
    DataString += Format.BinarySizeT( ProcessAffinityMask );
    DataString += PXS_STRING_CRLF;

    DataString += L"System affinity mask    : ";
    DataString += Format.BinarySizeT( SystemAffinityMask );
    DataString += PXS_STRING_CRLF;

    DataString += L"All processors available: ";
    if ( ProcessAffinityMask == SystemAffinityMask )
    {
        DataString += PXS_STRING_YES;
    }
    else
    {
        DataString += PXS_STRING_NO;
    }
    DataString += PXS_STRING_CRLF;

    ////////////////////////////////////////////////////////////////////////////
    // Physical Processors

    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;
    DataString += L"PHYSICAL PROCESSORS\r\n";
    DataString += L"===================\r\n";
    DataString += PXS_STRING_CRLF;
    try
    {
        GetPackageIDs( &PackageIDs );
        numPackages = PackageIDs.GetSize();
        DataString += L"Number of packages      : ";
        DataString += Format.SizeT( numPackages );
        DataString += PXS_STRING_CRLF;
        DataString += L"Package identifiers     : ";
        Value = PXS_STRING_EMPTY;
        for ( i = 0; i < numPackages; i++ )
        {
            Value += Format.UInt8( packageID );
            if ( i < ( numPackages - 1 ) )
            {
                Value += L", ";
            }
        }
        DataString += Value;
        DataString += PXS_STRING_CRLF;
        DataString += PXS_STRING_CRLF;
        for ( i = 0; i < numPackages; i++ )
        {
            // Run on this package
            packageID = PackageIDs.Get( i );
            RunCurrentThreadOnPackageID( packageID );

            DataString += L"Physical Processor#     : ";
            DataString += Format.SizeT( i + 1 );
            DataString += PXS_STRING_CRLF;

            DataString += L"Package identifier      : ";
            DataString += Format.UInt8( packageID );
            DataString += PXS_STRING_CRLF;

            Value = PXS_STRING_EMPTY;
            GetApicsIDs( packageID, &Value );
            DataString += L"APIC Identifier(s)      : ";
            DataString += Value;
            DataString += PXS_STRING_CRLF;

            byteValue   = GetBrandID();
            DataString += L"Brand ID                : ";
            DataString += Format.UInt8Hex( byteValue, true );
            DataString += PXS_STRING_CRLF;

            Value = PXS_STRING_EMPTY;
            GetManufacturerName( &Value );
            DataString += L"Manufacturer Name       : ";
            DataString += Value;
            DataString += PXS_STRING_CRLF;

            EAX = 0;
            Cpuid( 0, &EAX, nullptr, nullptr, nullptr );
            DataString += L"Maximum Standard Func.  : ";
            DataString += Format.UInt32( EAX );
            DataString += PXS_STRING_CRLF;

            EAX = 0;
            Cpuid( 0x80000000, &EAX, nullptr, nullptr, nullptr );
            DataString += L"Maximum Extend Function : ";
            DataString += Format.UInt32Hex( EAX, true );
            DataString += PXS_STRING_CRLF;

            byteValue   = GetMaximumLogicalsInPackage();
            DataString += L"Max. Logicals/Physical  : ";
            DataString += Format.UInt8( byteValue );
            DataString += PXS_STRING_CRLF;

            Value = PXS_STRING_EMPTY;
            GetName( &Value );
            DataString += L"Name                    : ";
            DataString += Value;
            DataString += PXS_STRING_CRLF;

            GetSignature( &stepping,
                          &model, &family, &type, &modelExt, &familyExt );
            DataString += L"Signature - Stepping    : ";
            DataString += Format.UInt8Hex( stepping, true );;
            DataString += PXS_STRING_CRLF;

            DataString += L"Signature - Model       : ";
            DataString += Format.UInt8Hex( model, true );
            DataString += PXS_STRING_CRLF;

            DataString += L"Signature - Family      : ";
            DataString += Format.UInt8Hex( family, true );
            DataString += PXS_STRING_CRLF;

            DataString += L"Signature - Type        : ";
            DataString += Format.UInt8Hex( type, true );
            DataString += PXS_STRING_CRLF;

            DataString += L"Signature - Model Ext.  : ";
            DataString += Format.UInt8Hex( modelExt, true );
            DataString += PXS_STRING_CRLF;

            DataString += L"Signature - Family Ext. : ";
            DataString += Format.UInt8Hex( familyExt, true );
            DataString += PXS_STRING_CRLF;

            speedMhz    = GetSpeedEstimateMHz();
            DataString += L"Estimated speed [MHz]   : ";
            DataString += Format.UInt32( speedMhz );
            DataString += PXS_STRING_CRLF;

            speedMhz    = GetSpeedRegistryMHz( packageID );
            DataString += L"Registry speed [MHz]    : ";
            DataString += Format.UInt32( speedMhz );
            DataString += PXS_STRING_CRLF;

            DataString += L"Multi-Core/HTT          : ";
            DataString += Format.Int32YesNo( SupportsMultiCore() );
            DataString += PXS_STRING_CRLF;

            Value = PXS_STRING_EMPTY;
            GetType( &Value );
            DataString += L"Type                    : ";
            DataString += Value;
            DataString += PXS_STRING_CRLF;

            Value = PXS_STRING_EMPTY;
            GetVendorString( &Value );
            DataString += L"Vendor String           : ";
            DataString += Value;
            DataString += PXS_STRING_CRLF;
        }
    }
    catch ( const Exception& e )
    {
        DataString += e.GetMessage();
        DataString += PXS_STRING_CRLF;
    }
    DataString += PXS_STRING_CRLF;

    ////////////////////////////////////////////////////////////////////////////
    // Logical Processors

    DataString += PXS_STRING_CRLF;
    DataString += PXS_STRING_CRLF;
    DataString += L"LOGICAL PROCESSORS\r\n";
    DataString += L"==================\r\n\r\n";

    // Need to reset the affinity
    try
    {
        // Loop through the logical processors
        for ( i = 0; i < ( 8 * sizeof ( DWORD_PTR ) ); i++ )
        {
            ThreadMask = ( (DWORD_PTR)1 << i );
            if ( ThreadMask & ProcessAffinityMask )
            {
                DataString += PXS_STRING_CRLF;
                DataString += L"-----------------------------------------";
                DataString += PXS_STRING_CRLF;
                DataString += L"Processor# ";
                DataString += Format.BinarySizeT( ThreadMask );
                DataString += PXS_STRING_CRLF;
                DataString += L"-----------------------------------------";
                DataString += PXS_STRING_CRLF;
                DataString += PXS_STRING_CRLF;
                if ( SetThreadAffinityMask( hThread, ThreadMask ) )
                {
                    Sleep( 0 );     // switch CPU

                    // APIC ID
                    byteValue = GetApicID();
                    DataString += L"APIC Identifier             : ";
                    DataString += Format.UInt8Hex( byteValue, true );
                    DataString += PXS_STRING_CRLF;

                    // Get the maximum standard function
                    Cpuid( 0, &maxStdFunc, nullptr, nullptr, nullptr );
                    DataString += L"Max. Standard CPUID Function: ";
                    DataString += Format.UInt32( maxStdFunc );
                    DataString += PXS_STRING_CRLF;

                    // Get the maximum extended function
                    Cpuid( 0x80000000, &maxExtFunc, nullptr, nullptr, nullptr );
                    DataString += L"Max. Extended CPUID Function: ";
                    DataString += Format.UInt32Hex( maxExtFunc, true );
                    DataString += PXS_STRING_CRLF;
                    DataString += PXS_STRING_CRLF;

                    // Standard registers
                    Registers = PXS_STRING_EMPTY;
                    GetRegistersAsString( true, &Registers );
                    DataString += Registers;
                    DataString += PXS_STRING_CRLF;
                    DataString += PXS_STRING_CRLF;

                    // Extended registers
                    Registers = PXS_STRING_EMPTY;
                    GetRegistersAsString( false, &Registers );
                    DataString += Registers;
                    DataString += PXS_STRING_CRLF;
                    DataString += PXS_STRING_CRLF;
                }
                else
                {
                    // Error setting the mask
                    DataString += L"SetThreadAffinityMask failed, mask = ";
                    DataString += Format.BinarySizeT( ThreadMask );
                    DataString += L", error code = ";
                    DataString += Format.UInt32( GetLastError() );
                    DataString += PXS_STRING_CRLF;
                }
                DataString += PXS_STRING_CRLF;
                DataString += PXS_STRING_CRLF;
            }
        }
    }
    catch ( const Exception& e )
    {
        DataString += e.GetMessage();
        DataString += PXS_STRING_CRLF;
    }
    SetThreadAffinityMask( hThread, ProcessAffinityMask );
    *pDiagnostics = DataString;
}

//===============================================================================================//
//  Description:
//      Get the features for the specified processor/package
//
//  Parameters:
//      packageID - the physical processor/package number
//      pFeatures - receives the features
//
//  Returns:
//      void
//===============================================================================================//
void CpuInformation::GetFeatures( BYTE packageID, StringArray* pFeatures )
{
    DWORD_PTR oldThreadAffinityMask = 0;

    if ( pFeatures == nullptr )
    {
        throw ParameterException( L"pFeatures", __FUNCTION__ );
    }
    pFeatures->RemoveAll();

    // If have not got the APIC identifiers do so now
    if ( m_LogicalApics.GetSize() == 0 )
    {
        FillLogicalApics();
    }
    oldThreadAffinityMask = RunCurrentThreadOnPackageID( packageID );

    try
    {
        if ( PXS_CPU_MANUFACTURER_INTEL == GetManufacturerID() )
        {
            GetFeaturesIntel( pFeatures );
        }
        else
        {
            GetFeaturesAmd( pFeatures );
        }
    }
    catch ( const Exception& )
    {
        SetThreadAffinityMask( GetCurrentThread(), oldThreadAffinityMask );
        throw;
    }
    SetThreadAffinityMask( GetCurrentThread(), oldThreadAffinityMask );
    pFeatures->Sort( true );
}

//===============================================================================================//
//  Description:
//      Get the features of the AMD processor/package the current thread
//      is running on
//
//  Parameters:
//      pFeaturesAmd - receives the features
//
//  Returns:
//      void
//===============================================================================================//
void CpuInformation::GetFeaturesAmd( StringArray* pFeaturesAmd )
{
    DWORD  ECX = 0, EDX = 0;
    size_t i   = 0;

    if ( pFeaturesAmd == nullptr )
    {
        throw ParameterException( L"pFeaturesAmd", __FUNCTION__ );
    }
    pFeaturesAmd->RemoveAll();

    if ( PXS_CPU_MANUFACTURER_AMD != GetManufacturerID() )
    {
        throw SystemException( ERROR_INVALID_FUNCTION, L"PXS_CPU_MANUFACTURER_AMD", __FUNCTION__ );
    }

    // Function ECX.1
    struct _STANDARD_ECX
    {
        DWORD   bit;
        LPCWSTR pszFeature;
    } StdEcx[] = { { 1 <<  0, L"SSE3"      },
                   { 1 <<  1, L"PCLMULQDQ" },
                   { 1 <<  3, L"MONITOR"   },
                   { 1 <<  9, L"SSSE3"     },
                   { 1 << 12, L"FMA"       },
                   { 1 << 13, L"CMPXCHG16B"},
                   { 1 << 19, L"SSE41"     },
                   { 1 << 20, L"SSE42"     },
                   { 1 << 22, L"MOVBE"     },
                   { 1 << 23, L"POPCNT"    },
                   { 1 << 25, L"AES"       },
                   { 1 << 26, L"XSAVE"     },
                   { 1 << 27, L"OSXSAVE"   },
                   { 1 << 28, L"AVX"       },
                   { 1 << 29, L"F16C"      } };
    ECX = 0;
    Cpuid( 1, nullptr, nullptr, &ECX, nullptr );
    for( i = 0; i < ARRAYSIZE( StdEcx ); i++ )
    {
        if ( ECX & StdEcx[ i ].bit )
        {
            pFeaturesAmd->Add( StdEcx[ i ].pszFeature );
        }
    }

    // Function EDX.1
    struct _STANDARD_EDX
    {
        DWORD   bit;
        LPCWSTR pszFeature;
    } StdEdx[] = { { 1 <<  0, L"FPU"            },
                   { 1 <<  1, L"VME"            },
                   { 1 <<  2, L"DE"             },
                   { 1 <<  3, L"PSE"            },
                   { 1 <<  4, L"TSC"            },
                   { 1 <<  5, L"MSR"            },
                   { 1 <<  6, L"PAE"            },
                   { 1 <<  7, L"MCE"            },
                   { 1 <<  8, L"CMPXCHG8B"      },
                   { 1 <<  9, L"APIC"           },
                   { 1 << 11, L"SysEnterSysExit"},
                   { 1 << 12, L"MTRR"           },
                   { 1 << 13, L"PGE"            },
                   { 1 << 14, L"MCA"            },
                   { 1 << 15, L"CMOV"           },
                   { 1 << 16, L"PAT"            },
                   { 1 << 17, L"PSE36"          },
                   { 1 << 19, L"CLFSH"          },
                   { 1 << 23, L"MMX"            },
                   { 1 << 24, L"FXFR"           },
                   { 1 << 25, L"SSE"            },
                   { 1 << 26, L"SSE2"           },
                   { 1 << 28, L"HTT"            } };
    EDX = 0;
    Cpuid( 1, nullptr, nullptr, nullptr, &EDX );
    for( i = 0; i < ARRAYSIZE( StdEdx ); i++ )
    {
        if ( EDX & StdEdx[ i ].bit )
        {
            pFeaturesAmd->Add( StdEdx[ i ].pszFeature );
        }
    }

    // Function ECX.80000001
    struct _EXTENDED_ECX
    {
        DWORD   bit;
        LPCWSTR pszFeature;
    } ExtEcx[] = { { 1 <<  0, L"LahfSahf"               },
                   { 1 <<  1, L"CmpLegacy"              },
                   { 1 <<  2, L"SVM"                    },
                   { 1 <<  3, L"ExtApicSpace"           },
                   { 1 <<  4, L"AltMovCr8"              },
                   { 1 <<  5, L"ABM"                    },
                   { 1 <<  6, L"SSE4A"                  },
                   { 1 <<  7, L"MisAlignSse"            },
                   { 1 <<  8, L"3DNowPrefetch"          },
                   { 1 <<  9, L"OSVW"                   },
                   { 1 << 10, L"IBS"                    },
                   { 1 << 11, L"XOP"                    },
                   { 1 << 12, L"SKINIT"                 },
                   { 1 << 13, L"WDT"                    },
                   { 1 << 15, L"LWP"                    },
                   { 1 << 16, L"FMA4"                   },
                   { 1 << 21, L"TBM"                    },
                   { 1 << 22, L"TopologyExtension"      },
                   { 1 << 23, L"PerfCtrExtCore"         },
                   { 1 << 24, L"PerfCtrExtNB"           },
                   { 1 << 25, L"StreamPerfMon"          },
                   { 1 << 26, L"DataBreakpointExtension"},
                   { 1 << 27, L"PerfTsc"                } };

    ECX = 0;
    Cpuid( 0x80000001, nullptr, nullptr, &ECX, nullptr );
    for( i = 0; i < ARRAYSIZE( ExtEcx ); i++ )
    {
        if ( ECX & ExtEcx[ i ].bit )
        {
            pFeaturesAmd->Add( ExtEcx[ i ].pszFeature );
        }
    }

    // Function EDX.80000001. Will omit duplicates from EDX.1
    struct _EXTENDED_EDX
    {
        DWORD   bit;
        LPCWSTR pszFeature;
    } ExtEdx[] = { {   1 << 11, L"SysCallSysRet" },
                   {   1 << 20, L"NX"            },
                   {   1 << 22, L"MmxExt"        },
                   {   1 << 25, L"FFXSR"         },
                   {   1 << 26, L"Page1GB"       },
                   {   1 << 27, L"RDTSCP"        },
                   {   1 << 29, L"LM"            },
                   {   1 << 30, L"3DNowExt"      },
                   { 1UL << 31, L"3DNow"         } };

    EDX = 0;
    Cpuid( 0x80000001, nullptr, nullptr, nullptr, &EDX );
    for( i = 0; i < ARRAYSIZE( ExtEdx ); i++ )
    {
        if ( EDX & ExtEdx[ i ].bit )
        {
            pFeaturesAmd->Add( ExtEdx[ i ].pszFeature );
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the features of the Intel processor/package the current thread
//      is running on
//
//  Parameters:
//      pFeaturesIntel - receives the features
//
//  Returns:
//      void
//===============================================================================================//
void CpuInformation::GetFeaturesIntel( StringArray* pFeaturesIntel )
{
    size_t i   = 0;
    DWORD  ECX = 0, EDX = 0;

    if ( pFeaturesIntel == nullptr )
    {
        throw ParameterException( L"pFeaturesIntel", __FUNCTION__ );
    }
    pFeaturesIntel->RemoveAll();

    if ( PXS_CPU_MANUFACTURER_INTEL != GetManufacturerID() )
    {
        throw SystemException( ERROR_INVALID_FUNCTION,
                               L"PXS_CPU_MANUFACTURER_INTEL", __FUNCTION__ );
    }

    // Function ECX.1
    struct _STANDARD_ECX
    {
        DWORD   bit;
        LPCWSTR pszFeature;
    } StdEcx[] = { { 1 <<  0, L"SSE3"         },
                   { 1 <<  1, L"PCLMULQDQ"    },
                   { 1 <<  2, L"DTES64"       },
                   { 1 <<  3, L"MONITOR"      },
                   { 1 <<  4, L"DS-CPL"       },
                   { 1 <<  5, L"VMX"          },
                   { 1 <<  6, L"SMX"          },
                   { 1 <<  7, L"EIST"         },
                   { 1 <<  8, L"TM2"          },
                   { 1 <<  9, L"SSSE3"        },
                   { 1 << 10, L"CNXT-ID"      },
                   { 1 << 12, L"FMA"          },
                   { 1 << 13, L"CMPXCHG16B"   },
                   { 1 << 14, L"xTPR"         },
                   { 1 << 15, L"PDCM"         },
                   { 1 << 17, L"PCID"         },
                   { 1 << 18, L"DCA"          },
                   { 1 << 19, L"SSE4.1"       },
                   { 1 << 20, L"SSE4.2"       },
                   { 1 << 21, L"x2APIC"       },
                   { 1 << 22, L"MOVBE"        },
                   { 1 << 23, L"POPCNT"       },
                   { 1 << 24, L"TSC-Deadline" },
                   { 1 << 25, L"AESNI"        },
                   { 1 << 26, L"XSAVE"        },
                   { 1 << 27, L"OSXSAVE"      },
                   { 1 << 28, L"AVX"          },
                   { 1 << 29, L"F16C"         },
                   { 1 << 30, L"RDRAND"       } };
    ECX = 0;
    Cpuid( 1, nullptr, nullptr, &ECX, nullptr );
    for( i = 0; i < ARRAYSIZE( StdEcx ); i++ )
    {
        if ( ECX & StdEcx[ i ].bit )
        {
            pFeaturesIntel->Add( StdEcx[ i ].pszFeature );
        }
    }

    // Function EDX.1
    struct _STANDARD_EDX
    {
        DWORD   bit;
        LPCWSTR pszFeature;
    } StdEdx[] = { {   1 <<  0, L"FPU"   },
                   {   1 <<  1, L"VME"   },
                   {   1 <<  2, L"DE"    },
                   {   1 <<  3, L"PSE"   },
                   {   1 <<  4, L"TSC"   },
                   {   1 <<  5, L"MSR"   },
                   {   1 <<  6, L"PAE"   },
                   {   1 <<  7, L"MCE"   },
                   {   1 <<  8, L"CX8"   },
                   {   1 <<  9, L"APIC"  },
                   {   1 << 11, L"SEP"   },
                   {   1 << 12, L"MTRR"  },
                   {   1 << 13, L"PGE"   },
                   {   1 << 14, L"MCA"   },
                   {   1 << 15, L"CMOV"  },
                   {   1 << 16, L"PAT"   },
                   {   1 << 17, L"PSE-36"},
                   {   1 << 18, L"PSN"   },
                   {   1 << 19, L"CLFSH" },
                   {   1 << 21, L"DS"    },
                   {   1 << 22, L"ACPI"  },
                   {   1 << 23, L"MMX"   },
                   {   1 << 24, L"FXFR"  },
                   {   1 << 25, L"SSE"   },
                   {   1 << 26, L"SSE2"  },
                   {   1 << 27, L"SS"    },
                   {   1 << 28, L"HTT"   },
                   {   1 << 29, L"TM"    },
                   { 1UL << 31, L"PBE"   } };
    EDX = 0;
    Cpuid( 1, nullptr, nullptr, nullptr, &EDX );
    for( i = 0; i < ARRAYSIZE( StdEdx ); i++ )
    {
        if ( EDX & StdEdx[ i ].bit )
        {
            pFeaturesIntel->Add( StdEdx[ i ].pszFeature );
        }
    }

    // Function ECX.80000001
    ECX = 0;
    Cpuid( 0x80000001, nullptr, nullptr, &ECX, nullptr );
    if ( ECX & 1 )
    {
        pFeaturesIntel->Add( L"LAHF" );
    }

    // Function EDX.80000001
    struct _EXTENDED_EDX
    {
        DWORD   bit;
        LPCWSTR pszFeature;
    } ExtEdx[] = { { 1 << 11, L"SYSCALL/SYSRET" },
                   { 1 << 20, L"EDB"            },
                   { 1 << 26, L"1-GByte"        },
                   { 1 << 27, L"RDTSCP"         },
                   { 1 << 29, L"IA64"           } };
    EDX = 0;
    Cpuid( 0x80000001, nullptr, nullptr, nullptr, &EDX );
    for( i = 0; i < ARRAYSIZE( ExtEdx ); i++ )
    {
        if ( EDX & ExtEdx[ i ].bit )
        {
            pFeaturesIntel->Add( ExtEdx[ i ].pszFeature );
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the name and speed of the specified processor/package
//
//  Parameters:
//      packageID - the physical processor/package number
//      pName     - string object to receive name
//      pSpeedMHz - receives the speed in MHz
//
//  Returns:
//      void, "" and zero if cannot get data
//===============================================================================================//
void CpuInformation::GetNameAndSpeedMHz( BYTE packageID, String* pName, DWORD* pSpeedMHz )
{
    DWORD_PTR oldThreadAffinityMask = 0;

    if ( ( pName == nullptr ) || ( pSpeedMHz == nullptr ) )
    {
        throw ParameterException( L"pName/pSpeedMHz", __FUNCTION__ );
    }
    *pName     = PXS_STRING_EMPTY;
    *pSpeedMHz = 0;

    // If have not got the APIC identifiers, do so now
    if ( m_LogicalApics.GetSize() == 0 )
    {
        FillLogicalApics();
    }
    oldThreadAffinityMask = RunCurrentThreadOnPackageID( packageID );

    // Need to reset the thread's mask
    try
    {
        GetName( pName );
        *pSpeedMHz = GetSpeedEstimateMHz();
        if ( *pSpeedMHz == 0 )
        {
            *pSpeedMHz = GetSpeedRegistryMHz( packageID );
        }
     }
    catch ( const Exception& )
    {
        SetThreadAffinityMask( GetCurrentThread(), oldThreadAffinityMask );
        throw;
    }
    SetThreadAffinityMask( GetCurrentThread(), oldThreadAffinityMask );
}

//===============================================================================================//
//  Description:
//      Get the package IDs of the physical processors on the system
//
//  Parameters:
//      pPackageIDs - array to receive the package ids
//
//  Remarks:
//      Package IDs are obtained though APIC physical identifiers which
//      are BYTE's hence package identifiers are bytes as well
//
//  Returns:
//      void
//===============================================================================================//
void CpuInformation::GetPackageIDs( TArray< BYTE >* pPackageIDs )
{
    bool   match = false;
    BYTE   packageID = 0, apicID = 0;
    size_t i = 0, j = 0, numApics = 0, size = 0;

    if ( pPackageIDs == nullptr )
    {
        throw ParameterException( L"pPackageIDs", __FUNCTION__ );
    }
    pPackageIDs->RemoveAll();

    // If have not got the processor(s) APIC's do so now
    if ( m_LogicalApics.GetSize() == 0 )
    {
        FillLogicalApics();
    }

    numApics = m_LogicalApics.GetSize();
    for ( i = 0; i< numApics; i++ )
    {
        apicID    = m_LogicalApics.Get( i );
        packageID = GetPackageIDFromApicID( apicID );

        // Add to output array if not already present
        match = false;
        size  = pPackageIDs->GetSize();
        for ( j = 0; j < size; j++ )
        {
            if ( packageID == pPackageIDs->Get( j ) )
            {
                match = true;
                break;
            }
        }

        if ( match == false )
        {
            pPackageIDs->Add( packageID );
        }
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Execute the CPUID instruction at the specified function number on the
//      logical processor on which the current thread is running
//
//  Parameters:
//      function - the function number
//      pEAX     - receives the contents of the EAX register, NULL is allowed
//      pEBX     - receives the contents of the EBX register, NULL is allowed
//      pECX     - receives the contents of the ECX register, NULL is allowed
//      pEDX     - receives the contents of the EDX register, NULL is allowed
//
//  Remarks:
//      The x64 MSVC compiler no longer supports in line asm so will use the
//      __cpuid intrinsic. However, __cpuid take int's but will work with
//      unsigned values to reduce casts elsewhere in this module.
//
//      No longer checking that the processor supports the CPUID function.
//
//  Returns:
//      void
//===============================================================================================//
void CpuInformation::Cpuid( DWORD  function, DWORD* pEAX, DWORD* pEBX, DWORD* pECX, DWORD* pEDX )
{
    int       InfoType = 0;
    int       CPUInfo[ 4 ] = { 0 };
    String    Error, FunctionNumber;
    Formatter Format;

    // Verify the function is supported
    if ( function > 0x7FFFFFFF )
    {
        __cpuid( CPUInfo, -1 );    // = 0x80000000
    }
    else
    {
        __cpuid( CPUInfo, 0 );
    }

    InfoType = static_cast<int>( function );
    if ( InfoType > CPUInfo[ 0 ]  )
    {
       FunctionNumber = Format.UInt32Hex( function, true );
       Error = Format.String1( L"Function = %%1.", FunctionNumber );
       throw SystemException( ERROR_INVALID_FUNCTION, Error.c_str(), __FUNCTION__ );
    }
    memset( CPUInfo, 0, sizeof ( CPUInfo ) );
    __cpuid( CPUInfo, InfoType );

    if ( pEAX ) *pEAX = static_cast<DWORD>( CPUInfo[ 0 ] );
    if ( pEBX ) *pEBX = static_cast<DWORD>( CPUInfo[ 1 ] );
    if ( pECX ) *pECX = static_cast<DWORD>( CPUInfo[ 2 ] );
    if ( pEDX ) *pEDX = static_cast<DWORD>( CPUInfo[ 3 ] );
}

//===============================================================================================//
//  Description:
//      Fill a class scope array of the APIC identifiers of each logical
//      processor
//
//  Parameters:
//      None
//
//  Remarks:
//      The index of the array is the bit position of the logical processor
//      as given in the dwActiveProcessorMask member of the SYSTEM_INFO
//      structure.
//
//      The physical APIC identifiers are used to map processor "packages"
//      to logical processors as used by Window API functions
//
//  Returns:
//      void
//===============================================================================================//
void CpuInformation::FillLogicalApics()
{
    BYTE    apicID    = 0;
    DWORD   lastError = ERROR_SUCCESS;
    size_t  i = 0;
    DWORD_PTR    ActiveProcessorMask = 0, ThreadAffinityMask = 0;
    DWORD_PTR    ProcessAffinityMask = 0, SystemAffinityMask = 0;
    SYSTEM_INFO  si;
    TArray< BYTE > LogicalApics;

    // Reset the class scope structure
    m_LogicalApics.RemoveAll();

    // Get the processor mask
    memset( &si, 0, sizeof ( si ) );
    GetNativeSystemInfo( &si );
    ActiveProcessorMask = si.dwActiveProcessorMask;

    // Determine if this process can run on all logicals
    if ( GetProcessAffinityMask( GetCurrentProcess(),
                                 &ProcessAffinityMask, &SystemAffinityMask ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetProcessAffinityMask", __FUNCTION__ );
    }

    // Catch exceptions so can restore the thread's processor mask
    try
    {
        i = 0;
        while ( ( lastError == ERROR_SUCCESS ) && (ActiveProcessorMask != 0) )
        {
            // Skip logicals not allowed for the current process
            apicID = PXS_INVALID_APIC_ID;
            if ( ActiveProcessorMask % 2 )
            {
                // Set the thread to run on the desired logical
                ThreadAffinityMask = ( (DWORD_PTR)1 << i );     // **TYPE CAST**
                if ( SetThreadAffinityMask( GetCurrentThread(), ThreadAffinityMask ) )
                {
                    // Yield so on the next time slice the current thread
                    // will run on the desired logical
                    Sleep( 0 );
                    apicID = GetApicID();
                }
                else
                {
                    lastError = GetLastError();     // Store it
                }
            }
            LogicalApics.Add( apicID );

            // Next loop
            ActiveProcessorMask = ( ActiveProcessorMask / 2 );
            i++;
        }
        SetThreadAffinityMask( GetCurrentThread(), si.dwActiveProcessorMask );
    }
    catch ( const Exception& )
    {
        SetThreadAffinityMask( GetCurrentThread(), si.dwActiveProcessorMask );
        throw;
    }

    // Test for error raised by SetThreadAffinityMask
    if ( lastError != ERROR_SUCCESS )
    {
        throw SystemException( lastError, L"SetThreadAffinityMask", __FUNCTION__ );
    }

    // Store at class scope
    m_LogicalApics = LogicalApics;
}

//===============================================================================================//
//  Description:
//      Get the APIC ID of the logical processor on which the current
//      thread is running
//
//  Parameters:
//      None
//
//  Remarks:
//      The APIC physical identifier is in bits 24-31 of EBX, function 1
//
//  Returns:
//      Byte value of the APIC ID
//===============================================================================================//
BYTE CpuInformation::GetApicID()
{
    DWORD EBX = 0;

    Cpuid( 1, nullptr, &EBX, nullptr, nullptr );

    return PXSCastUInt32ToUInt8( EBX >> 24 );   // The byte in bits 24-31
}

//===============================================================================================//
//  Description:
//      Get APID identifiers belonging to the specified package id as a
//      comma separated string of hex formatted bytes.
//
//  Parameters:
//      packageID - the package identifier
//      pApicIDs  - string to receive the APIC identifiers
//
//  Remarks:
//      For a physical processor there may be multiple logical processors,
//      each with its own APIC ID
//
//  Returns:
//      void
//===============================================================================================//
void CpuInformation::GetApicsIDs( BYTE packageID, String* pApicIDs )
{
    BYTE      apicID = 0;
    size_t    i = 0, numApics = 0;
    Formatter Format;

    if ( pApicIDs == nullptr )
    {
        throw ParameterException( L"pApicIDs", __FUNCTION__ );
    }
    *pApicIDs = PXS_STRING_EMPTY;

    // Search for matching APIC's
    numApics = m_LogicalApics.GetSize();
    for ( i = 0; i < numApics; i++ )
    {
        apicID = m_LogicalApics.Get( i );
        if ( packageID == GetPackageIDFromApicID( apicID ) )
        {
            if ( pApicIDs->GetLength() )
            {
                *pApicIDs += L", ";
            }
            *pApicIDs += Format.UInt8Hex( apicID, true );
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the brand id/index of the logical processor that the current
//      thread is running on
//
//  Parameters:
//      none
//
//  Remarks:
//    Zero is a valid brand id, so caller should test boolean returned value
//
//  Returns:
//      The brand ID
//===============================================================================================//
BYTE CpuInformation::GetBrandID()
{
    DWORD EAX = 0;

    // The brand ID is in EAX of function 1
    Cpuid( 1, &EAX, nullptr, nullptr, nullptr );

    return PXSCastUInt32ToUInt8( 0xff & EAX );    // Bits 0-7
}

//===============================================================================================//
//  Description:
//      Get the cache abd TLB information about the logical on which the
//      current thread is running
//
//  Parameters:
//      pCacheInfo - receives the cache information
//      pTLBInfo   - receives the TLB information
//
//  Returns:
//     void
//===============================================================================================//
void CpuInformation::GetCacheAndTLBInfo( String* pCacheInfo, String* pTLBInfo )
{
    size_t i = 0, numElements = 0;
    String Description;
    StringArray CacheDescriptions;

    if ( ( pCacheInfo == nullptr ) || ( pTLBInfo == nullptr ) )
    {
        throw ParameterException( L"pCacheInfo/pTLBInfo", __FUNCTION__ );
    }
    *pCacheInfo = PXS_STRING_EMPTY;
    *pTLBInfo   = PXS_STRING_EMPTY;

    if ( PXS_CPU_MANUFACTURER_INTEL == GetManufacturerID() )
    {
        GetCacheDescriptionsIntel( &CacheDescriptions );
    }
    else
    {
        GetCacheDescriptionsAmd( &CacheDescriptions );
    }

    // Make comma separated string
    CacheDescriptions.Sort( true );
    numElements = CacheDescriptions.GetSize();
    for ( i = 0; i < numElements; i++ )
    {
        Description = CacheDescriptions.Get( i );
        if ( Description.StartsWith( L"TLB", false ) )
        {
            if ( pTLBInfo->GetLength() )
            {
                *pTLBInfo += L", ";
            }
            *pTLBInfo += Description;
        }
        else
        {
            if ( pCacheInfo->GetLength() )
            {
                *pCacheInfo += L", ";
            }
            *pCacheInfo += Description;
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the TLB information about the AMD logical on which the current
//      thread is running
//
//  Parameters:
//      pCacheDescriptions - receives the cache descriptions
//
//  Returns:
//     void
//===============================================================================================//
void CpuInformation::GetCacheDescriptionsAmd(
                                           StringArray* pCacheDescriptionsAmd )
{
    DWORD  EAX = 0, EBX = 0, ECX = 0, EDX = 0, maxExtFunction = 0;
    String Description;
    Formatter   Format;

    if ( pCacheDescriptionsAmd == nullptr )
    {
        throw ParameterException( L"pCacheDescriptionsAmd", __FUNCTION__ );
    }
    pCacheDescriptionsAmd->RemoveAll();

    if ( PXS_CPU_MANUFACTURER_AMD != GetManufacturerID() )
    {
        throw SystemException( ERROR_INVALID_FUNCTION, L"PXS_CPU_MANUFACTURER_AMD", __FUNCTION__ );
    }

    Cpuid( 0x80000000, &maxExtFunction, nullptr, nullptr, nullptr );
    if ( maxExtFunction >= 0x80000005 )
    {
        EAX = 0;
        EBX = 0;
        ECX = 0;
        EDX = 0;
        Cpuid( 0x80000005, &EAX, &EBX, &ECX, &EDX );

        // L1ITlb2and4MSize
        if ( EAX & 0xFF )
        {
            pCacheDescriptionsAmd->Add( L"TLB Instruction 2MB/4MB" );
        }

        // L1DTlb2and4MSize
        if ( ( EAX >> 16 ) & 0xFF )
        {
            pCacheDescriptionsAmd->Add( L"TLB Data 2MB/4MB" );
        }

        // L1ITlb4KSize
        if ( EBX & 0xFF )
        {
            pCacheDescriptionsAmd->Add( L"TLB Instruction 4KB" );
        }

        // L1DTlb4KSize
        if ( ( EBX >> 16 ) & 0xFF )
        {
            pCacheDescriptionsAmd->Add( L"TLB Data 4KB" );
        }

        // L1 Data Cache
        if ( ECX >> 24 )
        {
            Description = Format.StringUInt32( L"L1 Data %%1KB", ECX >> 24 );
            pCacheDescriptionsAmd->Add( Description );
        }

        // L1 Instruction Cache
        if ( EDX >> 24 )
        {
            Description = Format.StringUInt32( L"L1 Instruction %%1KB", EDX >> 24 );
            pCacheDescriptionsAmd->Add( Description );
        }
    }

    // L2 and L3 Cache
    if ( maxExtFunction >= 0x80000006 )
    {
        ECX = 0;
        EDX = 0;
        Cpuid( 0x80000006, nullptr, nullptr, &ECX, &EDX );
        if ( ECX >> 16 )
        {
            Description = Format.StringUInt32( L"L2 %%1KB", ECX >> 16 );
            pCacheDescriptionsAmd->Add( Description );
        }

        if ( EDX >> 18 )
        {
            Description = Format.StringUInt32( L"L3 %%1KB", (EDX >> 18) * 512 );
            pCacheDescriptionsAmd->Add( Description );
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the TLB information about the Intel logical on which the current
//      thread is running
//
//  Parameters:
//      pCacheDescriptionsIntel - receives the cache descriptions
//
//  Returns:
//     void
//===============================================================================================//
void CpuInformation::GetCacheDescriptionsIntel(
                                         StringArray* pCacheDescriptionsIntel )
{
    BYTE   descriptor = 0;
    size_t i = 0, j = 0, k = 0, start = 0;
    DWORD registers[ 4 ] = { 0 };      // EAX to EDX

    if ( pCacheDescriptionsIntel == nullptr )
    {
       throw ParameterException( L"pCacheDescriptionsIntel", __FUNCTION__ );
    }
    pCacheDescriptionsIntel->RemoveAll();

    if ( PXS_CPU_MANUFACTURER_INTEL != GetManufacturerID() )
    {
        throw SystemException( ERROR_INVALID_FUNCTION,
                               L"PXS_CPU_MANUFACTURER_INTEL", __FUNCTION__ );
    }

    struct _CACHE_INFO
    {
        BYTE    descriptor;
        LPCWSTR pszCache;
    } Caches[] =  { { 0x01, L"TLB Instruction 4KB"                },
                    { 0x02, L"TLB Instruction 4MB"                },
                    { 0x03, L"TLB Data 4KB"                       },
                    { 0x04, L"TLB Data 4MB"                       },
                    { 0x05, L"TLB Data 4MB"                       },
                    { 0x06, L"L1 Instruction 8KB"                 },
                    { 0x08, L"L1 Instruction 16KB"                },
                    { 0x09, L"L1 Instruction 32KB"                },
                    { 0x0A, L"L1 Data 8KB"                        },
                    { 0x0B, L"TLB Instruction 4MB"                },
                    { 0x0C, L"L1 Data 16KB"                       },
                    { 0x0D, L"L1 Data 16KB"                       },
                    { 0x0E, L"L1 Data 24KB"                       },
                    { 0x21, L"L2 256KB"                           },
                    { 0x22, L"L3 512KB"                           },
                    { 0x23, L"L3 1MB"                             },
                    { 0x25, L"L3 2MB"                             },
                    { 0x29, L"L3 4MB"                             },
                    { 0x2C, L"L1 Data 32KB"                       },
                    { 0x30, L"L1 Instruction 32KB"                },
                    { 0x41, L"L2 128KB"                           },
                    { 0x42, L"L2 256KB"                           },
                    { 0x43, L"L2 512KB"                           },
                    { 0x44, L"L2 1MB"                             },
                    { 0x45, L"L2 2MB"                             },
                    { 0x46, L"L3 4MB"                             },
                    { 0x47, L"L3 8MB"                             },
                    { 0x48, L"L2 3MB"                             },
                    { 0x49, L"L2 or L3 4MB"                       },
                    { 0x4A, L"L3 6MB"                             },
                    { 0x4B, L"L3 8MB"                             },
                    { 0x4C, L"L3 12MB"                            },
                    { 0x4D, L"L3 16MB"                            },
                    { 0x4E, L"L2 6MB"                             },
                    { 0x4F, L"TLB Instruction 4KB"                },
                    { 0x50, L"TLB Instruction 4KB and 2MB or 4MB" },
                    { 0x51, L"TLB Instruction 4KB and 2MB or 4MB" },
                    { 0x52, L"TLB Instruction 4KB and 2MB or 4MB" },
                    { 0x55, L"TLB Instruction 2MB or 4MB"         },
                    { 0x56, L"TLB Data 4MB"                       },
                    { 0x57, L"TLB Data 4KB"                       },
                    { 0x59, L"TLB Data 4KB"                       },
                    { 0x5A, L"TLB Data 2MB or 4MB"                },
                    { 0x5B, L"TLB Data 4KB and 4MB"               },
                    { 0x5C, L"TLB Data 4KB and 4MB"               },
                    { 0x5D, L"TLB Data 4KB and 4MB"               },
                    { 0x60, L"L1 Data 16KB"                       },
                    { 0x63, L"TLB Data 1GB"                       },
                    { 0x66, L"L1 Data 8KB"                        },
                    { 0x67, L"L1 Data 16KB"                       },
                    { 0x68, L"L1 Data 32KB"                       },
                    { 0x70, L"Trace 12K-uop"                      },
                    { 0x71, L"Trace 16K-uop"                      },
                    { 0x72, L"Trace 32K-uop"                      },
                    { 0x76, L"TLB Instruction 2MB/4MB"            },
                    { 0x78, L"L2 1MB"                             },
                    { 0x79, L"L2 128KB"                           },
                    { 0x7A, L"L2 256KB"                           },
                    { 0x7B, L"L2 512KB"                           },
                    { 0x7C, L"L2 1MB"                             },
                    { 0x7D, L"L2 2MB"                             },
                    { 0x7F, L"L2 512KB"                           },
                    { 0x80, L"L2 512KB"                           },
                    { 0x82, L"L2 256KB"                           },
                    { 0x83, L"L2 512KB"                           },
                    { 0x84, L"L2 1MB"                             },
                    { 0x85, L"L2 2MB"                             },
                    { 0x86, L"L2 512KB"                           },
                    { 0x87, L"L2 1MB"                             },
                    { 0xB0, L"TLB Instruction 4KB"                },
                    { 0xB1, L"TLB Instruction 2MB"                },
                    { 0xB0, L"TLB Instruction 4KB"                },
                    { 0xB3, L"TLB Data 4KB"                       },
                    { 0xB4, L"TLB Data 4KB"                       },
                    { 0xBA, L"TLB Data 4KB"                       },
                    { 0xC0, L"TLB Data 4KB and 4MB"               },
                    { 0xCA, L"STLB Shared 2nd-Level 4KB"          },
                    { 0xD0, L"L3 512KB"                           },
                    { 0xD1, L"L3 1MB"                             },
                    { 0xD2, L"L3 2MB"                             },
                    { 0xD6, L"L3 1MB"                             },
                    { 0xD7, L"L3 2MB"                             },
                    { 0xD8, L"L3 4MB"                             },
                    { 0xDC, L"L3 1.5MB"                           },
                    { 0xDD, L"L3 3MB"                             },
                    { 0xDE, L"L3 6MB"                             },
                    { 0xE2, L"L3 2MB"                             },
                    { 0xE3, L"L3 4MB"                             },
                    { 0xE4, L"L3 8MB"                             },
                    { 0xEA, L"L3 12MB"                            },
                    { 0xEB, L"L3 18MB"                            },
                    { 0xEC, L"L3 24MB"                            },
                    { 0xF0, L"Prefetching 64-Byte"                },
                    { 0xF1, L"Prefetching 128-Byte"               } };

    // Need standard function 2, requires Pentium Pro and Pentium II
    // Processors or newer. One call with EAX = 2, will ignore additional calls.
    Cpuid( 2, &registers[0], &registers[1], &registers[2], &registers[3] );
    for ( i = 0; i < ARRAYSIZE( registers ); i++ )
    {
        // Valid data in descriptor if bit 31 is zero
        if ( ( registers[ i ] >> 31 ) == 0 )
        {
            // Skip the least-significant byte of register EAX as this
            // is the number times to call CPUID
            start = 0;
            if ( i == 1 )
            {
                start = 1;
            }
            for ( j = start; j < 4; j++ )
            {
                // Docs say if the descriptor is 0xFF need to use EAX=4.
                // However, do not yet have 0xFF in the Caches array
                descriptor = 0xFF & ( registers[ i ] >> ( j * 8 ) );
                for ( k = 0; k < ARRAYSIZE( Caches ); k++ )
                {
                    if ( descriptor == Caches[ k ].descriptor )
                    {
                        pCacheDescriptionsIntel->AddUniqueI( Caches[ k ].pszCache );
                    }
                }
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the manufacturer ID of the logical processor on which the
//      current thread is running
//
//  Parameters:
//      None
//
//  Remarks:
//      Only AMD and Intel are recognised.
//
//      "AMDisbetter!" may have been used on older processors, see the CPUID
//      article on Wikipedia
//
//  Returns:
//     Constant representing the manufacturer, otherwise unknown
//===============================================================================================//
DWORD CpuInformation::GetManufacturerID()
{
    DWORD  manufacturerID = PXS_CPU_MANUFACTURER_UNKNOWN;
    String Vendor;

    GetVendorString( &Vendor );
    if ( Vendor.CompareI( L"GenuineIntel" ) == 0 )
    {
        manufacturerID = PXS_CPU_MANUFACTURER_INTEL;
    }
    else if ( Vendor.CompareI( L"AuthenticAMD" ) == 0 )
    {
        manufacturerID = PXS_CPU_MANUFACTURER_AMD;
    }

    return manufacturerID;
}

//===============================================================================================//
//  Description:
//      Get the name of the manufacturer of the specified package/process
//      that the current thread is running on
//
//  Parameters:
//      pManufacturerName - receives the manufacturer's name
//
//  Remarks:
//
//  Returns:
//      void, output has "" if manufacturer is not recognised
//===============================================================================================//
void CpuInformation::GetManufacturerName( String* pManufacturerName )
{
    DWORD  manufacturerID;
    String Insert2;

    if ( pManufacturerName == nullptr )
    {
        throw ParameterException( L"pManufacturerName", __FUNCTION__ );
    }
    *pManufacturerName = PXS_STRING_EMPTY;

    manufacturerID = GetManufacturerID();
    if ( manufacturerID == PXS_CPU_MANUFACTURER_INTEL )
    {
        *pManufacturerName = L"Intel(R) Corporation";
    }
    else if ( manufacturerID == PXS_CPU_MANUFACTURER_AMD )
    {
        *pManufacturerName = L"Advanced Micro Devices";
    }
    else
    {
        // Log it
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogAppError2( L"Unrecognised manufacturer '%%2' in '%%1'.",
                         *pManufacturerName, Insert2 );
    }
}

//===============================================================================================//
//  Description:
//      Get the maximum number of logical processors on the physical
//      processor that the current thread is running on.
//
//  Parameters:
//      None
//
//  Remarks:
//      The returned value is the "maximum" number of logicals the architecture
//      supports. The processor may not implement maximum number. For example,
//      a quad core Xeon(R) X5560 with hyper-threading has 8 logical processors
//      but CPUID will report there are 8 cores hence 16 logical processors.
//
//  Returns:
//      BYTE value of logical per physical
//===============================================================================================//
BYTE CpuInformation::GetMaximumLogicalsInPackage()
{
    DWORD EBX = 0;

    if ( SupportsMultiCore() == false )
    {
        return 1;    // No hyper-threading/multi-core
    }

    // On both Intel and AMD, maximum logicals per processor is in
    // CPUID.1:EBX[23:16]
    Cpuid( 1, nullptr, &EBX, nullptr, nullptr );

    return PXSCastUInt32ToUInt8( 0xFF & ( EBX >> 16 ) );  // Byte at 16-23;
}

//===============================================================================================//
//  Description:
//      Get the package/processor's name that the current thread is
//      running on
//
//  Parameters:
//      pName - receives the processor's name
//
//  Remarks:
//      The processors name is in EAX, EBX, ECX and EDX of extended functions
//      0x80000002 to 0x80000004, i.e 48 bytes
//
//  Returns:
//      void, output has "" if cannot get a name
//===============================================================================================//
void CpuInformation::GetName( String* pName )
{
    char  szName[ 128 ] = { 0 };              // Big enough for a processor name
    DWORD i, EAX, EBX, ECX, EDX;

    if ( pName == nullptr )
    {
        throw ParameterException( L"pName", __FUNCTION__ );
    }
    *pName = PXS_STRING_EMPTY;

    // Execute 0x80000002 to 0x80000004, each one get 16 chars of data
    for ( i = 0; i < 3; i++ )
    {
        EAX = 0;
        EBX = 0;
        ECX = 0;
        EDX = 0;
        Cpuid( 0x80000002 + i, &EAX, &EBX, &ECX, &EDX );
        memcpy( szName +  0 + ( 16 * i ), &EAX, sizeof ( EAX ) );
        memcpy( szName +  4 + ( 16 * i ), &EBX, sizeof ( EBX ) );
        memcpy( szName +  8 + ( 16 * i ), &ECX, sizeof ( ECX ) );
        memcpy( szName + 12 + ( 16 * i ), &EDX, sizeof ( EDX ) );
    }
    szName[ ARRAYSIZE( szName ) - 1 ] = 0x00;
    pName->SetAnsi( szName );
    pName->Trim();
}

//===============================================================================================//
//  Description:
//      Determine the number of cores in the specified physical package on
//      which the current thread is running
//
//  Parameters:
//      packageID - the package id on which the current thread is running
//
//  Returns:
//      The count
//===============================================================================================//
BYTE CpuInformation::GetNumberCoresInPackage()
{
    BYTE  numCores = 0;
    DWORD EBX = 0, ECX = 0, maxExtFunc = 0;

    if ( SupportsMultiCore() == false )
    {
        return 1;    // No hyper-threading/multi-core
    }

    if ( PXS_CPU_MANUFACTURER_UNKNOWN == GetManufacturerID() )
    {
        // FIXME: ECX should be set to zero but MSVC now need asm files.
        // Need to get the maximum number of cores addressable then
        // determine the mask. CPUID.(EAX=4, ECX=0):EAX[31:26] + 1
        throw SystemException( ERROR_CALL_NOT_IMPLEMENTED,
                               L"CPUID.(EAX=4, ECX=0)",  __FUNCTION__ );
    }
    else
    {
        // AMD. Test for legacy mode with CPUID Fn8000_0001_ECX[1]
        Cpuid( 0x80000001, nullptr, nullptr, &ECX, nullptr );
        if ( ECX & 0x00000002 )  // = CmpLegacy
        {
            // Use LogicalProcessorCount CPUID Fn0000_0001_EBX[23:16]
            Cpuid( 0x00000001, nullptr, &EBX, nullptr, nullptr );
            numCores = static_cast< BYTE >( 0xFF & ( EBX >> 16 ) );
        }
        else
        {
            // Extended Method, Fn8000_0008_ECX[0:7] is the actual number of
            // cores implemented in the processor
            Cpuid( 0x80000000, &maxExtFunc, nullptr, nullptr, nullptr );
            if ( maxExtFunc < 0x80000008 )
            {
                throw SystemException( ERROR_BAD_DRIVER_LEVEL,
                                       L"CPUID Fn8000_0008_ECX", __FUNCTION__ );
            }
            Cpuid( 0x80000008, nullptr, nullptr, &ECX, nullptr );
            numCores = static_cast< BYTE >( 0xFF & ECX );
            numCores = PXSAddUInt8( numCores, 1 );
        }
    }

    return numCores;
}

//===============================================================================================//
//  Description:
//      Determine the number of logical processors in the specified physical
//      package.
//
//  Parameters:
//      packageID - the package id
//
//  Returns:
//      The count
//===============================================================================================//
BYTE CpuInformation::GetNumberLogicalsInPackage( BYTE packageID )
{
    BYTE   apicID = 0, count = 0;
    size_t i = 0, numApics = m_LogicalApics.GetSize();

    if ( numApics == 0 )
    {
        throw FunctionException( L"m_LogicalApics", __FUNCTION__ );
    }

    for ( i = 0; i < numApics; i++ )
    {
        apicID = m_LogicalApics.Get( i );
        if ( packageID == GetPackageIDFromApicID( apicID ) )
        {
            count = PXSAddUInt8( count, 1 );
        }
    }

    return count;
}

//===============================================================================================//
//  Description:
//    Determine the Package ID from the specified APIC ID
//
//  Parameters:
//      apicID - the APIC ID
//
//  Remarks:
//      The package id is stored in the APIC ID. Intel and AMD have
//      different formats so first must detect the manufacturer. See Intel's
//      CPUCount.cpp
//
//  Returns:
//      The package ID
//===============================================================================================//
BYTE CpuInformation::GetPackageIDFromApicID( BYTE apicID )
{
    BYTE   mask = 0, packageID = 0, maxLogicals = 0, shift = 0;
    DWORD  manufacturerID = 0;

    if ( apicID == PXS_INVALID_APIC_ID )
    {
        throw ParameterException( L"apicID", __FUNCTION__ );
    }

    // Determine the number of bits required to store the number of logical
    // processors per physical processor, i.e. 1 bit for 2, 2 bits for 4,
    // 3 bits for 8 etc.
    maxLogicals = GetMaximumLogicalsInPackage();
    while ( maxLogicals > 1 )        // i.e. multi-core
    {
        shift++;
        maxLogicals = (BYTE)( maxLogicals / 2 );
    }

    // Must be either Intel or AMD
    manufacturerID = GetManufacturerID();
    if ( manufacturerID == PXS_CPU_MANUFACTURER_INTEL )
    {
        // Intel APIC ID's contain four levels of topology: cluster, package,
        // core and thread. Will assume there are no clusters as per Intel's
        // CPUCount.cpp example
        mask      = (BYTE)( 0xFF & ( 0xFF << shift ) );
        packageID = (BYTE)( apicID & mask );
    }
    else if ( manufacturerID == PXS_CPU_MANUFACTURER_AMD )
    {
        // AMD: APIC ID = (PackageID << Shift) + Number of Logicals, so the
        // package id is in the high bits
        packageID = (BYTE)( apicID >> shift );
    }
    else
    {
        throw SystemException( ERROR_INVALID_DATA, PXS_STRING_EMPTY, __FUNCTION__ );
    }

    return packageID;
}

//===============================================================================================//
//  Description:
//      Get the value of the EAX..EDX registers as a string
//
//  Parameters:
//      standard   - true if want standard functions else extended functions
//      pRegisters - receives the formatted string
//
//  Remarks:
//
//  Returns:
//      true on success, otherwise false
//===============================================================================================//
void CpuInformation::GetRegistersAsString( bool standard, String* pRegisters )
{
    DWORD     EAX = 0, EBX = 0, ECX = 0, EDX = 0;
    DWORD     i = 0, minFunc = 0, maxFunc = 0;
    Formatter Format;

    if ( pRegisters == nullptr )
    {
        throw ParameterException( L"pRegisters", __FUNCTION__ );
    }
    *pRegisters = PXS_STRING_EMPTY;

    if ( standard )
    {
        minFunc = 0;
        Cpuid( 0, &maxFunc, nullptr, nullptr, nullptr );
    }
    else
    {
        minFunc = 0x80000000;
        Cpuid( 0x80000000, &maxFunc, nullptr, nullptr, nullptr );
    }

    for ( i = minFunc; i <= maxFunc; i++ )      // Inclusive
    {
        EAX = 0;
        EBX = 0;
        ECX = 0;
        EDX = 0;
        Cpuid( i, &EAX, &EBX, &ECX, &EDX );

        *pRegisters += L"CPUID.";
        *pRegisters += Format.UInt32Hex( i, true );
        *pRegisters += L".EAX = ";
        *pRegisters += Format.UInt32Hex( EAX, true );
        *pRegisters += PXS_STRING_CRLF;

        *pRegisters += L"CPUID.";
        *pRegisters += Format.UInt32Hex( i, true );
        *pRegisters += L".EBX = ";
        *pRegisters += Format.UInt32Hex( EBX, true );
        *pRegisters += PXS_STRING_CRLF;

        *pRegisters += L"CPUID.";
        *pRegisters += Format.UInt32Hex( i, true );
        *pRegisters += L".ECX = ";
        *pRegisters += Format.UInt32Hex( ECX, true );
        *pRegisters += PXS_STRING_CRLF;

        *pRegisters += L"CPUID.";
        *pRegisters += Format.UInt32Hex( i, true );
        *pRegisters += L".EDX = ";
        *pRegisters += Format.UInt32Hex( EDX, true );
        *pRegisters += PXS_STRING_CRLF;
        *pRegisters += PXS_STRING_CRLF;
    }
}

//===============================================================================================//
//  Description:
//      Get the "signature" of the logical processor that the current
//      thread is running on
//
//  Parameters:
//      pStepping  - receives the stepping value, NULL is allowed
//      pModel     - receives the model value, NULL is allowed
//      pFamily    - receives the family value, NULL is allowed
//      pType      - receives the type value, NULL is allowed
//      pModelExt  - receives the extended model, NULL is allowed
//      pFamilyExt - receives the extended family value, NULL is allowed
//
//  Remarks:
//
//  Returns:
//      true on success, otherwise false
//===============================================================================================//
void CpuInformation::GetSignature( BYTE* pStepping,
                                   BYTE* pModel,
                                   BYTE* pFamily, BYTE* pType, BYTE* pModelExt, BYTE* pFamilyExt )
{
    DWORD EAX = 0, EBX = 0, ECX = 0, EDX = 0;

    // The signature is in EAX of standard function 1
    Cpuid( 1, &EAX, &EBX, &ECX, &EDX );

    if ( pStepping )
    {
        // Bits 0-3
        *pStepping = PXSCastUInt32ToUInt8( 0xf & EAX );
    }

    if ( pModel )
    {
        // Bits 4-7
        *pModel = PXSCastUInt32ToUInt8( 0xf & ( EAX >> 4 ) );
    }

    if ( pFamily )
    {
        // Bits 8-11
        *pFamily = PXSCastUInt32ToUInt8( 0xf & ( EAX >> 8 ) );
    }

    if ( pType )
    {
        // Bits 12-13
        *pType = PXSCastUInt32ToUInt8( 0xf & ( EAX >> 12 ) );
    }

    if ( pModelExt )
    {
        // Bits 16-19
        *pModelExt = PXSCastUInt32ToUInt8( 0xf & ( EAX >> 16 ) );
    }

    if ( pFamilyExt )
    {
        // Bits 20-27
        *pFamilyExt = PXSCastUInt32ToUInt8( 0xf & ( EAX >> 20 ) );
    }
}

//===============================================================================================//
//  Description:
//      Estimate the speed in MHz of the processor/package that the current
//      thread is running on
//
//  Parameters:
//      packageID - the processor/package identifier
//
//  Remarks:
//      Will make the assumption that all the packages/processors are the same.
//      So can just estimate the speed of the logical processor on which the
//      current thread is running.
//
//      Using the RDTSC instruction which is available on family >=5 for
//      Intel. For AMD need to test extended function 0x80000001. RDTSC
//      functionality has been available for several years so will use
//      IsProcessorFeaturePresent as the expectation is the HAL supports it.
//
//      Not accounting for the time taken up by QueryPerformanceCounter itself.
//
//  Returns:
//      Speed in MHz, zero if cannot get a speed
//===============================================================================================//
DWORD CpuInformation::GetSpeedEstimateMHz()
{
    int      threadPriority = 0;
    DWORD    priorityClass = 0, speedMhz = 0, errorCode = 0;
    UINT64   speed = 0, rdtscStart = 0, rdtscEnd = 0;
    UINT64   frequencyQuadPart = 0, counts = 0;
    String   Insert1;
    HANDLE   hCurrentProcess = GetCurrentProcess();
    HANDLE   hCurrentThread  = GetCurrentThread();
    Formatter        Format;
    LARGE_INTEGER    Frequency, CounterStart, CounterEnd;

    // Want a high resolution performance counter
    memset( &Frequency,  0, sizeof ( Frequency ) );
    if ( QueryPerformanceFrequency( &Frequency ) == 0 )
    {
        throw SystemException( GetLastError(), L"QueryPerformanceFrequency", __FUNCTION__ );
    }
    if ( Frequency.QuadPart <= 0 )
    {
        throw SystemException( ERROR_INVALID_DATA, L"Frequency.QuadPart <= 0", __FUNCTION__ );
    }
    frequencyQuadPart = PXSCastInt64ToUInt64( Frequency.QuadPart );

    // Need RDTSC
    if ( IsProcessorFeaturePresent( PF_RDTSC_INSTRUCTION_AVAILABLE ) == 0 )
    {
        throw SystemException( ERROR_INVALID_FUNCTION,
                               L"PF_RDTSC_INSTRUCTION_AVAILABLE", __FUNCTION__);
    }

    // Set thread and process priorities
    priorityClass = GetPriorityClass( hCurrentProcess );
    if ( priorityClass == 0 )
    {
        throw SystemException( GetLastError(), L"GetPriorityClass", __FUNCTION__ );
    }

    threadPriority = GetThreadPriority( hCurrentThread );
    if ( threadPriority == THREAD_PRIORITY_ERROR_RETURN )
    {
        throw SystemException( GetLastError(), L"GetThreadPriority", __FUNCTION__ );
    }

    if ( SetPriorityClass( hCurrentProcess, REALTIME_PRIORITY_CLASS ) == 0 )
    {
        // Log but continue
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogSysError1( GetLastError(), L"SetPriorityClass failed in '%%1'.", Insert1 );
    }

    if ( SetThreadPriority( hCurrentThread, THREAD_PRIORITY_TIME_CRITICAL ) == 0 )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogSysError1( GetLastError(), L"SetThreadPriority failed in '%%1'.", Insert1 );
    }

    memset( &CounterStart, 0, sizeof ( CounterStart ) );
    memset( &CounterEnd  , 0, sizeof ( CounterEnd   ) );
    if ( QueryPerformanceCounter( &CounterStart ) == 0 )
    {
        throw SystemException( errorCode, L"QueryPerformanceCounter", __FUNCTION__ );
    }

    if ( QueryPerformanceCounter( &CounterStart ) )
    {
        rdtscStart = __rdtsc();
        Sleep( 250 );
        rdtscEnd = __rdtsc();
        QueryPerformanceCounter( &CounterEnd );
    }

    // Restore the thread priority
    if ( priorityClass )
    {
        SetPriorityClass( hCurrentProcess, priorityClass );
    }

    if ( THREAD_PRIORITY_ERROR_RETURN != threadPriority )
    {
        SetThreadPriority( hCurrentThread, threadPriority );
    }

    // Elapsed interval is performance_counts / performance_frequency
    // speed = rdtsc_cycles * performance_frequency / performance_counts
    PXSLogAppInfo1( L"RDTSC start: %%1", Format.Int64( CounterStart.QuadPart ) );
    PXSLogAppInfo1( L"RDTSC end  : %%1", Format.Int64( CounterEnd.QuadPart ) );
    if ( ( CounterEnd.QuadPart - CounterStart.QuadPart ) > 0 )
    {
        counts   = PXSCastInt64ToUInt64( CounterEnd.QuadPart - CounterStart.QuadPart );
        speed    = rdtscEnd - rdtscStart;
        speed    = PXSMultiplyUInt64( speed, frequencyQuadPart );
        speed    = speed / counts;
        speedMhz = PXSCastUInt64ToUInt32(  speed / 1000000 );
    }

    return speedMhz;
}

//===============================================================================================//
//  Description:
//      Get processor's speed as recorded in the registry for the specified
//      processor/package id
//
//  Parameters:
//      packageID - the package id
//
//  Remarks:
//      See article KB888282 "Different ways to determine CPU speed in Windows
//      XP or in Windows Server 2003". There are caveats to using the value
//      stored in the registry due to CPU power management features.

//  Returns:
//      The processor's speed in MHz, zero if not found
//===============================================================================================//
DWORD CpuInformation::GetSpeedRegistryMHz( BYTE packageID )
{
    BYTE      apicID = 0;
    DWORD     speedMHz = 0;
    size_t    i = 0;
    String    SubKey;
    Registry  RegObject;
    Formatter Format;

    // Assumption: Each logical in a processor package are the same speed. Will
    // take the speed of the first logical found that belongs to the package
    while ( ( speedMHz == 0 ) && ( i <  m_LogicalApics.GetSize() ) )
    {
        apicID = m_LogicalApics.Get( i );
        if ( packageID == GetPackageIDFromApicID( apicID ) )
        {
            SubKey  = L"HARDWARE\\DESCRIPTION\\System\\CentralProcessor\\";
            SubKey += Format.SizeT( i );
            RegObject.Connect( HKEY_LOCAL_MACHINE );
            RegObject.GetDoubleWordValue( SubKey.c_str(), L"~MHz", 0, &speedMHz );
        }
        i++;
    }

    return speedMHz;
}

//===============================================================================================//
//  Description:
//      Get standard features supported by the logical processor the current
//      thread is running on
//
//  Parameters:
//      None
//
//  Remarks:
//      EDX register, Function 1
//
//  Returns:
//      Mask
//===============================================================================================//
DWORD CpuInformation::GetStandardFeatureInformation()
{
    DWORD EDX = 0;

    Cpuid( 1, nullptr, nullptr, nullptr, &EDX );

    return EDX;
}

//===============================================================================================//
//  Description:
//      Get the type of processor the current thread is running on
//
//  Parameters:
//      pType - string object to receive the processor type
//
//  Remarks:
//      Package IDs are obtained though APIC physical identifiers which
//      are BYTE's hence package identifiers are bytes as well
//
//  Returns:
//      void
//===============================================================================================//
void CpuInformation::GetType( String* pType )
{
    DWORD     EAX = 0, processorType = 0;
    Formatter Format;

    if ( pType == nullptr )
    {
        throw ParameterException( L"pType", __FUNCTION__ );
    }
    *pType = PXS_STRING_EMPTY;

    // The processor type is in EAX of function 1
    Cpuid( 1, &EAX, nullptr, nullptr, nullptr );
    processorType = ( ( EAX >> 12 ) & 0x3 );    // Bits 12-13

    switch( processorType )
    {
        default:
            *pType = Format.UInt32( processorType );
            break;

        case 0:
            *pType = L"OEM Primary";
            break;

        case 1:
            *pType = L"Overdrive";
            break;

        case 2:
            *pType = L"Secondary";
            break;
    }
}

//===============================================================================================//
//  Description:
//      Get the vendor string of the logical processor on which the current
//      thread is running
//
//  Parameters:
//      pVendor - receives the name of the processor vendor
//
//  Returns:
//     void
//===============================================================================================//
void CpuInformation::GetVendorString( String* pVendorString )
{
    char  szBuffer[ 64 ] = { 0 };           // Big enough for a vendor string
    DWORD EBX = 0, ECX = 0, EDX = 0;
    Formatter Format;

    if ( pVendorString == nullptr )
    {
        throw ParameterException( L"pVendorString", __FUNCTION__ );
    }
    *pVendorString = PXS_STRING_EMPTY;

    // The vendor string is in EBX, EDX and EDX of function 0
    Cpuid( 0, nullptr, &EBX, &ECX, &EDX );

    // Fill vendor string - it is in ANSI format
    memcpy( szBuffer + 0, &EBX, sizeof ( EBX ) );
    memcpy( szBuffer + 4, &EDX, sizeof ( EDX ) );
    memcpy( szBuffer + 8, &ECX, sizeof ( ECX ) );
    szBuffer[ ARRAYSIZE( szBuffer ) - 1 ] = 0x00;

    // Convert to the current string encoding
    pVendorString->SetAnsi( szBuffer );
    pVendorString->Trim();
}

//===============================================================================================//
//  Description:
//      Set the current thread to run on the specified package/processor
//
//  Parameters:
//      packageID - the package/processor identifier
//
//  Remarks:
//      Assumption: all of the logicals in a physical are the same
//      Finds the first logical processor belonging to the physical processor
//      and then set the current thread's affinity mask.
//
//  Returns:
//      The threads previous affinity mask
//===============================================================================================//
DWORD_PTR CpuInformation::RunCurrentThreadOnPackageID( BYTE packageID )
{
    size_t    i = 0, idxLogical = 0, numApics = 0;
    BYTE      apicID = 0;
    String    Error;
    Formatter Format;
    DWORD_PTR ThreadAffinityMask = 0, oldMask;

    // Check class scope
    if ( m_LogicalApics.GetSize() == 0 )
    {
        throw FunctionException( L"m_LogicalApics", __FUNCTION__ );
    }

    // Get a logical processor belonging to the package
    idxLogical = PXS_MINUS_ONE;
    numApics   = m_LogicalApics.GetSize();
    for ( i = 0; i < numApics; i++ )
    {
        apicID = m_LogicalApics.Get( i );
        if ( packageID == GetPackageIDFromApicID( apicID ) )
        {
            idxLogical = i;
            break;
        }
    }

    // Make sure found a logical processor
    if ( idxLogical == PXS_MINUS_ONE )
    {
        Error  = L"packageID = ";
        Error += Format.UInt8( packageID );
        throw SystemException( ERROR_NOT_FOUND, Error.c_str(), __FUNCTION__ );
    }

    // Set the current thread to run on the logical
    ThreadAffinityMask = ( (DWORD_PTR)1 << idxLogical );    // **TYPE CAST**
    oldMask = SetThreadAffinityMask( GetCurrentThread(), ThreadAffinityMask );
    if ( oldMask == 0 )
    {
        throw SystemException( GetLastError(), L"SetThreadAffinityMask", __FUNCTION__ );
    }

    return oldMask;
}

//===============================================================================================//
//  Description:
//      Determine if the processor that the current thread is running
//      on supports is multi-core or hyper-threaded
//
//  Parameters:
//      None
//
//  Returns:
//      true if multi-core/hyper-threaded processor, otherwise false
//===============================================================================================//
bool CpuInformation::SupportsMultiCore()
{
    DWORD EDX = 0;

    // On both Intel and AMD, CPUID.1:EDX[28] = 1 if the processor is
    // multi-core or hyper-threaded
    Cpuid( 1, nullptr, nullptr, nullptr, &EDX );

    if ( EDX & ( 1 << 28 ) )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Translate the specified processor architecture to a string
//
//  Parameters:
//      processorArchitecture - defined constant, see SYSTEM_INFO structure
//      Translation           - receives the translation
//
//  Returns:
//      void
//===============================================================================================//
void CpuInformation::TranslateProcessorArchitecture( WORD processorArchitecture,
                                                     String* pTranslation )
{
    Formatter Format;

    if ( pTranslation == nullptr )
    {
        throw ParameterException( L"pTranslation", __FUNCTION__ );
    }

    switch ( processorArchitecture )
    {
        default:
            *pTranslation = Format.UInt16( processorArchitecture );
            break;

        case PROCESSOR_ARCHITECTURE_UNKNOWN:

            *pTranslation = L"PROCESSOR_ARCHITECTURE_UNKNOWN";
            break;

        case PROCESSOR_ARCHITECTURE_INTEL:

            *pTranslation = L"PROCESSOR_ARCHITECTURE_INTEL";
            break;

        case PROCESSOR_ARCHITECTURE_IA64:

            *pTranslation = L"PROCESSOR_ARCHITECTURE_IA64";
            break;

        case PROCESSOR_ARCHITECTURE_AMD64:

            *pTranslation = L"PROCESSOR_ARCHITECTURE_AMD64";
            break;
    }
}
