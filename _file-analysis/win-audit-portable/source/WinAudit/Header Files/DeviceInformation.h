///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Device Information Class Header
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

#ifndef WINAUDIT_DEVICE_INFORMATION_H_
#define WINAUDIT_DEVICE_INFORMATION_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/WinAudit.h"

// 2. C System Files
#include <SetupAPI.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/TList.h"

// 5. This Project

// 6. Forwards
class AuditRecord;
template< class T > class TArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class DeviceInformation
{
    public:
        // Default constructor
        DeviceInformation();

        // Destructor
        ~DeviceInformation();

        // Methods
        void    GetAuditRecords( TArray< AuditRecord >* pRecords );

    protected:
        // Data members

        // Methods

    private:
        // Structure to hold device information
        typedef struct _TYPE_DEVICE_INFO
        {
            GUID    ClassGuid;
            ULONG   statusCode;            // See CM_Get_DevNode_Status
            wchar_t szDeviceType[ 64 ];
            wchar_t szDeviceName[ 64 ];
            wchar_t szDescription[ 64 ];
            wchar_t szManufacturer[ 64 ];
            wchar_t szLocation[ 64 ];
            wchar_t szDriverProvider[ 64 ];
            wchar_t szDriverVersion[ 64 ];
            wchar_t szDriverDate[ 64 ];
            wchar_t szStatusMsg[ 128 ];
            wchar_t szDeviceID[ 128 ];
        } TYPE_DEVICE_INFO;

        // Copy constructor - not allowed
        DeviceInformation( const DeviceInformation& oDeviceInfo );

        // Assignment operator - not allowed
        DeviceInformation& operator= ( const DeviceInformation& oDeviceInfo );

        // Methods
        size_t  FillData();
        void    GetDeviceStatus( DEVINST dnDevInst,
                                 const GUID*  pGuid,
                                 PULONG pulProblemNumber,
                                 LPWSTR pszStatusMsg, size_t bufferChars );
        void    GetDriverInfo( HDEVINFO DeviceInfoSet,
                               PSP_DEVINFO_DATA  DeviceInfoData,
                               LPWSTR pszDriverProvider,
                               size_t numProviderChars,
                               LPWSTR pszDriverVersion,
                               size_t numVersionChars,
                               LPWSTR pszDriverDate, size_t numDateChars );
        size_t  GetNextDeviceType( String* pDeviceType, TArray< AuditRecord >* pRecords );
        void    Reset();
        void    TranslateDeviceCode( ULONG ulProblemNumber, String* pStatus );

        // Data members
        DWORD           m_uLastDeviceTypeIndex;
        StringArray     m_DeviceTypes;
        TList< TYPE_DEVICE_INFO > m_Devices;
};

#endif  // WINAUDIT_DEVICE_INFORMATION_H_
