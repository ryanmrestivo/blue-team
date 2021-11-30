///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Windows Management Instrumentation Class Header
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

#ifndef PXSBASE_WMI_H_
#define PXSBASE_WMI_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/PxsBase.h"

// 2. C System Files
#include <Wbemidl.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

// 6. Forwards
class NameValue;
template< class T > class TArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class Wmi
{
    public:
        // Default constructor
        Wmi();

        // Destructor
        ~Wmi();

        // Methods
        void    CloseQuery();
        void    Connect( LPCWSTR pszNameSpace );
        void    Disconnect();
        void    GetPropertyValues( TArray< NameValue >* pNameValues );
        void    ExecQuery( LPCWSTR pszQuery );
        void    Get( LPCWSTR pwszName, String* pValue );
        bool    GetBool( LPCWSTR pwszName, bool* pValue );
        bool    GetInt32( LPCWSTR pwszName, int* pValue );
        bool    GetUInt8( LPCWSTR pwszName, BYTE* pValue );
        DWORD   GetUInt8Array( LPCWSTR pwszName, BYTE* pBuffer, DWORD bufferBytes );
        bool    GetUInt16( LPCWSTR pwszName, WORD* pValue );
        bool    GetUInt32( LPCWSTR pwszName, DWORD* pValue );
        bool    GetUInt64( LPCWSTR pwszName, UINT64* pValue );
        bool    Next();

    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        Wmi( const Wmi& oWmiInformation );

        // Assignment operator - not allowed
        Wmi& operator= ( const Wmi& oWmiInformation );

        // Methods
        void    ApplyProxySecurity( IUnknown* pProxy );

        // Data members
        IEnumWbemClassObject* m_pIEnumWbemClassObject;
        IWbemClassObject*     m_pIWbemClassObject;
        IWbemLocator*         m_pIWbemLocator;
        IWbemServices*        m_pIWbemServices;
};

#endif  // PXSBASE_WMI_H_
