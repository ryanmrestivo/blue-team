////////////////////////////////////////////////////////////////////////////////
//
// Web Origin Settings Class Header
//
////////////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////
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
////////////////////////////////////////////////////////////////////////////////

#ifndef PXSBASE_WEB_ORIGIN_H_
#define PXSBASE_WEB_ORIGIN_H_

////////////////////////////////////////////////////////////////////////////////
// Remarks
////////////////////////////////////////////////////////////////////////////////

// See RFC 6454 "The Web Origin Concept"

// The host name is stored as Unicode Normalization Form C. IPv4 and IPv6
// address are permitted. The AsCii Encoded version is obtained using
// GetAceHost().

////////////////////////////////////////////////////////////////////////////////
// Includes
////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/PxsBase.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Object.h"
#include "PxsBase/Header Files/StringT.h"

// 5. This Project

// 6. Forwards

////////////////////////////////////////////////////////////////////////////////
// Interface
////////////////////////////////////////////////////////////////////////////////

class WebOrigin: public Object
{
    public:
        // Default constructor
        WebOrigin();

        // Destructor
        ~WebOrigin();

        // Copy constructor
        WebOrigin( const WebOrigin& oWebOrigin );

        // Assignment operator
        WebOrigin& operator= ( const WebOrigin& oWebOrigin );

        // Methods
        void  GetAceHost( String* pAceHost ) const;
const String& GetHost() const;
        WORD  GetPort() const;
const String& GetScheme() const;
        bool  IsSame( const WebOrigin& Origin ) const;
        bool  IsValid() const;
        void  Reset();
        void  Set( const String& Scheme, const String& AceHost, const String& Port );
const String& ToDisplayString();

        // Data members

    protected:
        // Methods

        // Data members

    private:
        // Methods

        // Data members
        WORD   m_port;
        String m_Host;
        String m_Scheme;
        String m_DisplayString;
};

#endif  // PXSBASE_WEB_ORIGIN_H_
