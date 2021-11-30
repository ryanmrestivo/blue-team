///////////////////////////////////////////////////////////////////////////////////////////////////
//
// HostInformation Class Header
//
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Copyright 1987-2016 PARMAVEX SERVICES
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

#ifndef PXSBASE_HOST_INFORMATION
#define PXSBASE_HOST_INFORMATION

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/StringT.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class HostInformation : public Object
{
    public:
        // Default constructor
        HostInformation();

        // Copy constructor
        HostInformation( const HostInformation& oHostInformation );

        // Assignment operator
        HostInformation& operator= ( const HostInformation& oHostInformation );

        // Destructor
        ~HostInformation();

        // Methods
const Exception& GetException() const;
const String&    GetHost() const;
const String&    GetPortNumber() const;
const SOCKADDR_STORAGE& GetSockAddr() const;
        DWORD    GetState() const;
        void     Reset();
        void     SetHostPortNumber( const String& Host, const String PortNumber );
        void     SetException( const Exception& e );
        void     SetSockAddr( const SOCKADDR_STORAGE& ss );
        void     SetState( DWORD state );

    protected:
        // Methods

        // Data members

    private:
        // Methods

        // Data members
        DWORD            m_uState;
        SOCKADDR_STORAGE m_ss;
        String           m_Host;
        String           m_PortNumber;
        Exception        m_Exception;

};

#endif  // PXSBASE_HOST_INFORMATION
