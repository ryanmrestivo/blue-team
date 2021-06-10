///////////////////////////////////////////////////////////////////////////////////////////////////
//
// HostInformation Class Implementation
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

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/HostInformation.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StringT.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
HostInformation::HostInformation()
                :m_uState( PXS_STATE_UNKNOWN ),
                 m_ss(),                       // Non-op
                 m_Host(),
                 m_PortNumber(),
                 m_Exception()
{
    memset( &m_ss, 0, sizeof ( m_ss ) );
}

// Copy constructor
HostInformation::HostInformation( const HostInformation& oHostInformation )
                :Object(),
                 m_uState( PXS_STATE_UNKNOWN ),
                 m_ss(),                       // Non-op
                 m_Host(),
                 m_PortNumber(),
                 m_Exception()
{
    memset( &m_ss, 0, sizeof ( m_ss ) );
    *this = oHostInformation;
}

// Destructor
HostInformation::~HostInformation()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
HostInformation& HostInformation::operator= ( const HostInformation& oHostInformation )
{
    if ( this == &oHostInformation ) return *this;

    m_uState     = oHostInformation.m_uState;
    memcpy( &m_ss, &oHostInformation.m_ss, sizeof ( m_ss ) );
    m_Host       = oHostInformation.m_Host;
    m_PortNumber = oHostInformation.m_PortNumber;
    m_Exception  = oHostInformation.m_Exception;

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the exception object associated with this string
//
//  Parameters:
//      None
//
//  Returns:
//      ref to the exception
//===============================================================================================//
const Exception& HostInformation::GetException() const
{
    return m_Exception;
}
//===============================================================================================//
//  Description:
//      Get the host string
//
//  Parameters:
//      None
//
//  Returns:
//      ref to the host string
//===============================================================================================//
const String& HostInformation::GetHost() const
{
    return m_Host;
}

//===============================================================================================//
//  Description:
//      Get the port number string
//
//  Parameters:
//      None
//
//  Returns:
//      ref to the port number string
//===============================================================================================//
const String& HostInformation::GetPortNumber() const
{
    return m_PortNumber;
}

//===============================================================================================//
//  Description:
//      Get the SOCKADDR_STORAGE structure associated with this object
//
//  Parameters:
//      None
//
//  Returns:
//      ref to the SOCKADDR_STORAGE structure
//===============================================================================================//
const SOCKADDR_STORAGE& HostInformation::GetSockAddr() const
{
    return m_ss;
}

//===============================================================================================//
//  Description:
//      Get the state of this object
//
//  Parameters:
//      None
//
//  Returns:
//      DWORD named constant
//===============================================================================================//
DWORD HostInformation::GetState() const
{
    return m_uState;
}

//===============================================================================================//
//  Description:
//      Init this object
//
//  Parameters:
//      None

//
//  Returns:
//      void
//===============================================================================================//
void HostInformation::Reset()
{
    m_uState     = PXS_STATE_UNKNOWN;
    memset( &m_ss, 0, sizeof ( m_ss ) );
    m_Host       = PXS_STRING_EMPTY;
    m_PortNumber = PXS_STRING_EMPTY;
    m_Exception.Reset();
}

//===============================================================================================//
//  Description:
//      Set the host and port number
//
//  Parameters:
//      Host       - the host string
//      PortNumber - the port number string
//
//  Returns:
//      void
//===============================================================================================//
void HostInformation::SetHostPortNumber( const String& Host, const String PortNumber )
{
    if ( ( Host.GetLength()       == 0 ) ||
         ( PortNumber.GetLength() == 0 )  )
    {
        throw ParameterException( L"Host/PortNumber", __FUNCTION__ );
    }
    m_Host       = Host;
    m_PortNumber = PortNumber;
}

//===============================================================================================//
//  Description:
//      Set an exception object
//
//  Parameters:
//      e - the exception
//
//  Returns:
//      void
//===============================================================================================//
void HostInformation::SetException( const Exception& e )
{
    m_Exception = e;
}

//===============================================================================================//
//  Description:
//      Set this object socket address
//
//  Parameters:
//      ss - the socket adfress
//
//  Returns:
//      void
//===============================================================================================//
void HostInformation::SetSockAddr( const SOCKADDR_STORAGE& ss )
{
    memcpy( &m_ss, &ss, sizeof ( m_ss ) );
}

//===============================================================================================//
//  Description:
//     Set the state of this oject
//
//  Parameters:
//      state - name state
//
//  Returns:
//      void
//===============================================================================================//
void HostInformation::SetState( DWORD state )
{
    if ( ( state != PXS_STATE_WAITING   ) &&
         ( state != PXS_STATE_RESOLVING ) &&
         ( state != PXS_STATE_COMPLETED )  )
    {
        throw ParameterException( L"state", __FUNCTION__ );
    }
    m_uState = state;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
