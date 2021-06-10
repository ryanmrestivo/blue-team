///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Device Information Class Implementation
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
#include "WinAudit/Header Files/DeviceInformation.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/Registry.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/SystemException.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
DeviceInformation::DeviceInformation()
                  :m_uLastDeviceTypeIndex( DWORD_MAX ),     // -1
                   m_DeviceTypes(),
                   m_Devices()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
DeviceInformation::~DeviceInformation()
{
    try
    {
        Reset();
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
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
//      Get the hardware devices as an array of audit records.
//
//  Parameters:
//      pRecords - array to receive the audit records
//
//  Returns:
//      void
//===============================================================================================//
void DeviceInformation::GetAuditRecords( TArray< AuditRecord >* pRecords )
{
    String DeviceType;
    TArray< AuditRecord > DeviceRecords;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();


    FillData();
    while ( GetNextDeviceType( &DeviceType , &DeviceRecords ) )
    {
        pRecords->Append( DeviceRecords );
        DeviceType = PXS_STRING_EMPTY;
        DeviceRecords.RemoveAll();
    };
    PXSSortAuditRecords( pRecords, PXS_HARDWARE_DEVS_DEVICE_TYPE );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Discover all the installed devices and fill the class scope pointer
//      list.
//
//  Parameters:
//      None
//
//  Remarks:
//      Fills an alphabetically sorted array of device types for
//      later enumeration by GetNextDeviceType
//
//      Both Plug-and-Play and Legacy/Hidden devices can be detected
//
//      Will avoid SetupDiClassGuidsFromName as names change and the
//      system may not have all relevant names. SetupDiClassGuidsFromName
//      function retrieves the GUID(s) associated with the specified class
//      name. This list is based on the classes currently installed
//      on the system.
//
//  Returns:
//      Number of devices found
//===============================================================================================//
size_t DeviceInformation::FillData()
{
    GUID       guidZero;
    DWORD      memberIndex = 0, requiredSize = 0;
    HDEVINFO   hDevInfo = nullptr;
    SP_DEVINFO_DATA  DeviceInfoData;
    TYPE_DEVICE_INFO DeviceInfo;

    Reset();

    // Create a HDEVINFO with all present devices.
    hDevInfo = SetupDiGetClassDevs( nullptr,
                                    nullptr,    // Enumerator
                                    nullptr, DIGCF_PRESENT | DIGCF_ALLCLASSES );
    if ( hDevInfo == INVALID_HANDLE_VALUE )
    {
        throw SystemException( GetLastError(), L"SetupDiGetClassDevs", __FUNCTION__ );
    }

    // Catch all errors to ensure release of handle
    try
    {
        // Enumerate through all devices
        memset( &DeviceInfoData, 0, sizeof ( DeviceInfoData ) );
        DeviceInfoData.cbSize = sizeof ( DeviceInfoData );
        while ( SetupDiEnumDeviceInfo( hDevInfo, memberIndex, &DeviceInfoData ) )
        {
            memberIndex++;

            // GUID can be zero, so filter those out
            memset( &guidZero, 0, sizeof ( guidZero ) );    // i.e all zeros
            if ( memcmp( &guidZero,
                         &DeviceInfoData.ClassGuid, sizeof ( guidZero ) ) )
            {
                memset( &DeviceInfo, 0, sizeof ( DeviceInfo ) );

                // GUID
                DeviceInfo.ClassGuid = DeviceInfoData.ClassGuid;

                // Get the status of the device
                GetDeviceStatus( DeviceInfoData.DevInst,
                                 &DeviceInfo.ClassGuid,
                                 &DeviceInfo.statusCode,
                                 DeviceInfo.szStatusMsg, ARRAYSIZE( DeviceInfo.szStatusMsg ) );

                // Class Description / Device Type.
                SetupDiGetClassDescription( &DeviceInfoData.ClassGuid,
                                            DeviceInfo.szDeviceType,
                                            ARRAYSIZE(DeviceInfo.szDeviceType),
                                            nullptr );
                DeviceInfo.szDeviceType[ ARRAYSIZE(DeviceInfo.szDeviceType) - 1 ] = PXS_CHAR_NULL;
                m_DeviceTypes.AddUniqueI( DeviceInfo.szDeviceType );

                // Friendly name
                SetupDiGetDeviceRegistryProperty(hDevInfo,
                                                 &DeviceInfoData,
                                                 SPDRP_FRIENDLYNAME,
                                                 nullptr,
                                                 reinterpret_cast<BYTE*>(&DeviceInfo.szDeviceName),
                                                 sizeof ( DeviceInfo.szDeviceName ),
                                                 nullptr );
                DeviceInfo.szDeviceName[
                      ARRAYSIZE( DeviceInfo.szDeviceName ) -1 ] = PXS_CHAR_NULL;

                // Device description
                SetupDiGetDeviceRegistryProperty(
                                              hDevInfo,
                                              &DeviceInfoData,
                                              SPDRP_DEVICEDESC,
                                              nullptr,
                                              reinterpret_cast<BYTE*>( &DeviceInfo.szDescription ),
                                              sizeof ( DeviceInfo.szDescription ),  // Bytes
                                              nullptr );
                DeviceInfo.szDescription[ ARRAYSIZE(DeviceInfo.szDescription)-1 ] = PXS_CHAR_NULL;

                // Device manufacturer
                SetupDiGetDeviceRegistryProperty(
                                              hDevInfo,
                                              &DeviceInfoData,
                                              SPDRP_MFG,
                                              nullptr,
                                              reinterpret_cast<BYTE*>(&DeviceInfo.szManufacturer),
                                              sizeof ( DeviceInfo.szManufacturer ),  // Bytes
                                              nullptr );
                DeviceInfo.szManufacturer[
                                      ARRAYSIZE( DeviceInfo.szManufacturer ) - 1 ] = PXS_CHAR_NULL;

                // Device location
                SetupDiGetDeviceRegistryProperty( hDevInfo,
                                                  &DeviceInfoData,
                                                  SPDRP_LOCATION_INFORMATION,
                                                  nullptr,
                                                  reinterpret_cast<BYTE*>(&DeviceInfo.szLocation),
                                                  sizeof ( DeviceInfo.szLocation ),  // Bytes
                                                  nullptr );
                DeviceInfo.szLocation[ ARRAYSIZE( DeviceInfo.szLocation ) - 1 ] = PXS_CHAR_NULL;

                // Driver information
                GetDriverInfo( hDevInfo,
                               &DeviceInfoData,
                               DeviceInfo.szDriverProvider,
                               ARRAYSIZE( DeviceInfo.szDriverProvider ),
                               DeviceInfo.szDriverVersion,
                               ARRAYSIZE( DeviceInfo.szDriverVersion ),
                               DeviceInfo.szDriverDate,
                               ARRAYSIZE( DeviceInfo.szDriverDate ) );

                // Device instance ID, this is system defined and is unique for
                // each device, it is persistent across system boots. Note, this
                // function takes number of characters
                requiredSize = 0;
                SetupDiGetDeviceInstanceId( hDevInfo,
                                            &DeviceInfoData,
                                            DeviceInfo.szDeviceID,
                                            ARRAYSIZE( DeviceInfo.szDeviceID ),
                                            &requiredSize );
                DeviceInfo.szDeviceID[ ARRAYSIZE( DeviceInfo.szDeviceID ) - 1 ] = PXS_CHAR_NULL;

                m_Devices.Append( DeviceInfo );
            }
        }
    }
    catch ( const Exception& )
    {
        SetupDiDestroyDeviceInfoList( hDevInfo );
        throw;
    }
    SetupDiDestroyDeviceInfoList( hDevInfo );
    m_DeviceTypes.Sort( true );

    return m_DeviceTypes.GetSize();
}

//===============================================================================================//
//  Description:
//      Get the status of the specified device
//
//  Parameters:
//      dnDevInst        - the device instance
//      pGuid            - the device class GUID
//      pStatusCode      - receives the status code
//      pszStatusMsg     - receives the status message
//      bufferChars      - size of the status message buffer
//
//  Returns:
//      void
//===============================================================================================//
void DeviceInformation::GetDeviceStatus( DEVINST dnDevInst,
                                         const GUID* pGuid,
                                         PULONG pStatusCode,
                                         LPWSTR pszStatusMsg, size_t bufferChars )
{
    ULONG     ulStatus = 0, ulProblemNumber = 0;
    String    GuidString, Status;
    Formatter Format;
    CONFIGRET crResult = 0;

    // Inputs, Note DevInst is a DWORD
    if ( ( pGuid        == nullptr ) ||
         ( pStatusCode  == nullptr ) ||
         ( pszStatusMsg == nullptr )  )
    {
        throw ParameterException( L"Pointer", __FUNCTION__ );
    }
    *pStatusCode = CR_SUCCESS;
    wmemset( pszStatusMsg, 0, bufferChars );

    // Note, the status code is ulProblemNumber
    // if the DN_HAS_PROBLEM bit is set in ulStatus
    crResult = CM_Get_DevNode_Status( &ulStatus,
                                      &ulProblemNumber,
                                      dnDevInst,
                                      0 );      // Must be zero
    if ( crResult != CR_SUCCESS )
    {
        GuidString = Format.GuidToString( *pGuid );
        PXSLogConfigManError1( crResult, L"CM_Get_DevNode_Status failed for '%%1.", GuidString );
        return;
    }

    // Test for a problem
    if ( DN_HAS_PROBLEM & ulStatus )
    {
        *pStatusCode = ulProblemNumber;
        TranslateDeviceCode( ulProblemNumber, &Status );
    }
    else
    {
        Status = L"OK";
    }

    if ( Status.c_str() )
    {
        StringCchCopy( pszStatusMsg, bufferChars, Status.c_str() );
    }
}

//===============================================================================================//
//  Description:
//      Get information about a devices driver.
//
//  Parameters:
//      DeviceInfoSet     - handle to the device instance
//      DeviceInfoData    - pointer to the device data
//      pszDriverProvider - receives driver's provider name
//      numProviderChars  - size of the driver provider buffer in chars
//      pszDriverVersion  - receives driver version
//      numVersionChars   - size of the driver version buffer in chars
//      pszDriverDate     - receives driver date name
//      numDateChars      - size of the driver date buffer in chars
//
//  Returns:
//      void
//===============================================================================================//
void DeviceInformation::GetDriverInfo( HDEVINFO DeviceInfoSet,
                                       PSP_DEVINFO_DATA  DeviceInfoData,
                                       LPWSTR pszDriverProvider,
                                       size_t numProviderChars,
                                       LPWSTR pszDriverVersion,
                                       size_t numVersionChars,
                                       LPWSTR pszDriverDate,
                                       size_t numDateChars )
{
    DWORD    lastError = 0;
    wchar_t  szBuffer[ MAX_PATH + 1 ] = { 0 };
    String   Value, DriverKey, Insert1;
    Registry RegObject;

    if ( ( DeviceInfoSet     == nullptr ) ||
         ( DeviceInfoData    == nullptr ) ||
         ( pszDriverProvider == nullptr ) ||
         ( pszDriverVersion  == nullptr ) ||
         ( pszDriverDate     == nullptr )  )
    {
        throw ParameterException( L"Pointer", __FUNCTION__ );
    }
    *pszDriverProvider = PXS_CHAR_NULL;
    *pszDriverVersion  = PXS_CHAR_NULL;
    *pszDriverDate     = PXS_CHAR_NULL;

    // Get the Driver Key
    if ( SetupDiGetDeviceRegistryProperty( DeviceInfoSet,
                                           DeviceInfoData,
                                           SPDRP_DRIVER,
                                           nullptr,
                                           reinterpret_cast<BYTE*>( &szBuffer ),
                                           sizeof ( szBuffer ),  // In bytes
                                           nullptr ) == 0 )
    {
        Insert1.SetAnsi( __FUNCTION__ );
        lastError = GetLastError();
        if ( lastError == ERROR_INVALID_DATA )
        {
            PXSLogSysInfo1( lastError,
                             L"SPDRP_DRIVER property for a device does not "
                             L"exist or is invalid in '%%1'.", Insert1 );
        }
        else
        {
            PXSLogSysError1( lastError,
                         L"SetupDiGetDeviceRegistryProperty failed for "
                         L"SPDRP_DRIVER in '%%1'.", Insert1 );
        }
        return;
    }
    szBuffer[ ARRAYSIZE( szBuffer ) - 1 ] = PXS_CHAR_NULL;

    // Make the registry path
    DriverKey  = L"System\\CurrentControlSet\\Control\\Class\\";
    DriverKey += szBuffer;
    RegObject.Connect( HKEY_LOCAL_MACHINE );

    // Driver provider name
    Value = PXS_STRING_EMPTY;
    RegObject.GetStringValue( DriverKey.c_str(), L"ProviderName", &Value );
    if ( Value.c_str() )
    {
        StringCchCopy( pszDriverProvider, numProviderChars, Value.c_str() );
    }

    // Driver version
    Value = PXS_STRING_EMPTY;
    RegObject.GetStringValue( DriverKey.c_str(), L"DriverVersion", &Value );
    if ( Value.c_str() )
    {
        StringCchCopy( pszDriverVersion, numVersionChars, Value.c_str() );
    }

    // Driver date
    Value = PXS_STRING_EMPTY;
    RegObject.GetStringValue( DriverKey.c_str(), L"DriverDate", &Value );
    if ( Value.c_str() )
    {
        StringCchCopy( pszDriverDate, numDateChars, Value.c_str() );
    }
}

//===============================================================================================//
//  Description:
//      Get the devices that are of the "next device type".
//
//  Parameters:
//      DeviceType - string object to receive the name of the device type
//      pRecords   - array to receive the audit records
//
//  Remarks:
//      Must have first called FillData.
//
//      Devices are enumerated by device type. Each time this method is called,
//      the device type is advance and DeviceType is set to indicate its name.
//      Records receives formatted audit records of all the devices of the
//      current device type.
//
//  Returns:
//      Number of devices of the named device type. Zero if none found
//===============================================================================================//
size_t DeviceInformation::GetNextDeviceType( String* pDeviceType,
                                             TArray< AuditRecord >* pRecords )
{
    String      Value;
    Formatter   Format;
    AuditRecord Record;
    const TYPE_DEVICE_INFO* pDI = nullptr;

    if ( ( pDeviceType == nullptr ) || ( pRecords == nullptr ) )
    {
        throw ParameterException( L"pDeviceType/pRecords", __FUNCTION__ );
    }
    *pDeviceType = PXS_STRING_EMPTY;
    pRecords->RemoveAll();

    // Get the name of the device type to enumerate in the sorted array.
    if ( m_uLastDeviceTypeIndex == DWORD_MAX )
    {
        // Starting a new enumeration
        m_uLastDeviceTypeIndex = 0;
    }
    else
    {
        // Continuing an existing enumeration
        m_uLastDeviceTypeIndex = PXSAddUInt32( m_uLastDeviceTypeIndex, 1 );
    }

    if ( m_uLastDeviceTypeIndex < m_DeviceTypes.GetSize() )
    {
        *pDeviceType = m_DeviceTypes.Get( m_uLastDeviceTypeIndex );
        pDeviceType->Trim();
    }

    if ( pDeviceType->IsEmpty() )
    {
        return 0;
    }

    if ( m_Devices.IsEmpty() )
    {
        return 0;
    }

    // Scan device the setup list to match the device types
    m_Devices.Rewind();
    do
    {
        pDI = m_Devices.GetPointer();
        if ( lstrcmpi( pDI->szDeviceType, pDeviceType->c_str() ) == 0 )
        {
            // Make the record
            Record.Reset( PXS_CATEGORY_HARDWARE_DEVICES );
            Record.Add( PXS_HARDWARE_DEVS_DEVICE_TYPE , *pDeviceType );
            Record.Add( PXS_HARDWARE_DEVS_DEVICE_NAME , pDI->szDeviceName );
            Record.Add( PXS_HARDWARE_DEVS_DESCRIPTION , pDI->szDescription );
            Record.Add( PXS_HARDWARE_DEVS_MANUFACTURER, pDI->szManufacturer);
            Record.Add( PXS_HARDWARE_DEVS_LOCATION    , pDI->szLocation );
            Record.Add( PXS_HARDWARE_DEVS_DRIVER_PROV , pDI->szDriverProvider);
            Record.Add( PXS_HARDWARE_DEVS_DRIVER_VER  , pDI->szDriverVersion );
            Record.Add( PXS_HARDWARE_DEVS_DRIVER_DATE , pDI->szDriverDate );

            Value = Format.UInt32( pDI->statusCode );
            Record.Add( PXS_HARDWARE_DEVS_STATUS_CODE, Value );

            Record.Add( PXS_HARDWARE_DEVS_STATUS_MSG, pDI->szStatusMsg );

            Value = Format.GuidToString( pDI->ClassGuid );
            Record.Add( PXS_HARDWARE_DEVS_CLASS_GUID, Value );

            Record.Add( PXS_HARDWARE_DEVS_DEVICE_ID, pDI->szDeviceID );
            pRecords->Add( Record );
        }
    } while ( m_Devices.Advance() );

    return pRecords->GetSize();
}

//===============================================================================================//
//  Description:
//      Reset the class scope variables
//
//  Parameters:
//      None
//
//  Remarks:
//      Called by the destructor, so avoid exceptions
//
//  Returns:
//      void
//===============================================================================================//
void DeviceInformation::Reset()
{
    m_uLastDeviceTypeIndex = DWORD_MAX;  // Reset the last device type to none
    m_Devices.RemoveAll();
    m_DeviceTypes.RemoveAll();
}

//===============================================================================================//
//  Description:
//      Get an error message from a device status code
//
//  Parameters:
//      ulProblemNumber - problem as reported by the device driver
//      pStatus         - string object to receive the status description
//
//  Remarks:
//      The API does not seem to specify a translation, neither does there
//      appear to be a string table in the setupapi.dll. So will use the
//      short descriptions a presented by Device Manager. See KB310123
//
//  Returns:
//      void
//===============================================================================================//
void DeviceInformation::TranslateDeviceCode( ULONG ulProblemNumber, String* pStatus)
{
    size_t    i = 0;
    Formatter Format;

    // Device manager codes, there does not seem to be documented method to
    // get these strings so will use the ones in MSDN's documentation
    struct _ERRORS
    {
        ULONG   ulProblem;
        LPCWSTR pszStatus;
    } Errors[] = {
        {  0, L"OK" },
        {  1, L"This device is not configured correctly."                 },
        {  3, L"The driver for this device might be corrupted, or your "
              L"system may be running low on memory or other resources."  },
        { 10, L"This device cannot start."                                },
        { 12, L"This device cannot find enough free resources that it can "
              L"use. If you want to use this device, you will need to "
              L"disable one of the other devices on this system."         },
        { 14, L"This device cannot work properly until you restart your "
              L"computer."                                                },
        { 16, L"Windows cannot identify all the resources this device "
              L"uses."                                                    },
        { 18, L"Reinstall the drivers for this device." },
        { 19, L"Windows cannot start this hardware device because its "
              L"configuration information (in the registry) is incomplete "
              L"or damaged."                                              },
        { 21, L"Windows is removing this device." },
        { 22, L"This device is disabled." },
        { 24, L"This device is not present, is not working properly, or "
              L"does not have all its drivers installed."                 },
        { 28, L"The drivers for this device are not installed." },
        { 29, L"This device is disabled because the firmware of the "
              L"device did not give it the required resources."           },
        { 31, L"This device is not working properly because Windows "
              L"cannot load the drivers required for this device."        },
        { 32, L"A driver (service) for this device has been disabled. An "
              L"alternate driver may be providing this functionality."    },
        { 33, L"Windows cannot determine which resources are required "
              L"for this device."                                         },
        { 34, L"Windows cannot determine the settings for this device. "
              L"Consult the documentation that came with this device "
              L"and use the Resource tab to set the configuration."       },
        { 35, L"Your computer's system firmware does not include enough "
              L"information to properly configure and use this device. "
              L"To use this device, contact your computer manufacturer"
              L" to obtain a firmware or BIOS update."                    },
        { 36, L"This device is requesting a PCI interrupt but is "
              L"configured for an ISA interrupt (or vice versa). Please "
              L"use the computer's system setup program to reconfigure "
              L"the interrupt for this device."                           },
        { 37, L"Windows cannot initialize the device driver for this "
              L"hardware."                                                },
        { 38, L"Windows cannot load the device driver for this hardware "
              L"because a previous instance of the device driver is "
              L"still in memory."                                         },
        { 39, L"Windows cannot load the device driver for this "
              L"hardware. The driver may be corrupted or missing."        },
        { 40, L"Windows cannot access this hardware because its service "
              L"key information in the registry is missing or recorded "
              L"incorrectly."                                             },
        { 41, L"Windows successfully loaded the device driver for this "
              L"hardware but cannot find the hardware device."            },
        { 42, L"Windows cannot load the device driver for this hardware "
              L"because there is a duplicate device already running in "
              L"the system."                                              },
        { 43, L"Windows has stopped this device because it has reported "
              L"problems."                                                },
        { 44, L"An application or service has shut down this hardware "
              L"device."                                                  },
        { 45, L"Currently, this hardware device is not connected to the "
              L"computer."                                                },
        { 46, L"Windows cannot gain access to this hardware device "
              L"because the operating system is in the process of "
              L"shutting down."                                           },
        { 47, L"Windows cannot use this hardware device because it has "
              L"been prepared for \"safe removal\", but it has not been "
              L"removed from the computer."                               },
        { 48, L"The software for this device has been blocked from "
              L"starting because it is known to have problems with "
              L"Windows. Contact the hardware vendor for a new driver."   },
        { 49, L"Windows cannot start new hardware devices because the "
              L"system hive is too large (exceeds the Registry Size "
              L"Limit.)"                                                  } };

    if ( pStatus == nullptr )
    {
        throw ParameterException( L"pStatus", __FUNCTION__ );
    }
    *pStatus = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Errors ); i++ )
    {
        if ( ulProblemNumber == Errors[ i ].ulProblem )
        {
            *pStatus += Errors[ i ].pszStatus;
        }
    }

    if ( pStatus->IsEmpty() )
    {
        *pStatus  = L"Code ";
        *pStatus += Format.UInt32( ulProblemNumber );
        PXSLogAppWarn1( L"Unrecognised ConfigMan error: '%%1'.", *pStatus );
    }
}
