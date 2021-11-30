///////////////////////////////////////////////////////////////////////////////////////////////////
//
// System Registry Class Header
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

#ifndef PXSBASE_REGISTRY_H_
#define PXSBASE_REGISTRY_H_

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
#include "PxsBase/Header Files/StringT.h"
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/TArray.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class Registry
{
    public:
        // Default constructor
        Registry();

        // Destructor
        ~Registry();

        // Methods
        void    Connect( HKEY hKey );
        void    Connect2( HKEY hKey, DWORD samDesired );
        void    Disconnect();
        DWORD   GetBinaryData( LPCWSTR pszSubKey,
                               LPCWSTR pszValueName, BYTE* pBinaryData, size_t bufferBytes );
        void    GetDoubleWordValue( LPCWSTR pszSubKey,
                                    LPCWSTR pszValueName,
                                    DWORD defaultValue, DWORD* pDoubleWordValue );
        void    GetNameValueList( LPCWSTR pszSubKey, TArray< NameValue >* pNameValueList );
        void    GetStringValue( LPCWSTR pszSubKey, LPCWSTR pszValueName, String* pStringValue );
        void    GetSubKeyList( LPCWSTR pszSubKey, StringArray* pSubKeyList ) const;
        void    GetValueAsString( LPCWSTR pszSubKey, LPCWSTR pszValueName, String* pStringValue );

    protected:
        // Data members

        // Data members

    private:
        // Copy constructor - not allowed
        Registry( const Registry& oRegistry );

        // Assignment operator - not allowed
        Registry& operator= ( const Registry& oRegistry );

        // Methods
        void    FormatRegistryBinaryAsString( const BYTE* pData,
                                              DWORD dataBytes, String* pStringValue );
        void    FormatRegistryDataAsString( const BYTE* pData,
                                            DWORD dataBytes, DWORD type, String* pStringValue );
        void    FormatRegistryMultiSzAsString( LPCWSTR pszMultiSz, String* pStringValue );

        // Data members
        HKEY    m_hKey;     // The connected registry key.
};

#endif  // PXSBASE_REGISTRY_H_
