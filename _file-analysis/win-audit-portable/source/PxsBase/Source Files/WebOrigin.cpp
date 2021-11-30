///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Web Origin Class Implementation
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
#include "PxsBase/Header Files/WebOrigin.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
WebOrigin::WebOrigin()
          :m_port( 0 ),
           m_Host(),
           m_Scheme(),
           m_DisplayString()
{
}

// Copy constructor
WebOrigin::WebOrigin( const WebOrigin& oWebOrigin )
          :Object(),
           m_port( 0 ),
           m_Host(),
           m_Scheme(),
           m_DisplayString()
{
    *this = oWebOrigin;
}

// Destructor
WebOrigin::~WebOrigin()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
WebOrigin& WebOrigin::operator= ( const WebOrigin& oWebOrigin )
{
    if ( this == &oWebOrigin ) return *this;

    m_port          = oWebOrigin.m_port;
    m_Host          = oWebOrigin.m_Host;
    m_Scheme        = oWebOrigin.m_Scheme;
    m_DisplayString = oWebOrigin.m_DisplayString;

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the Ascii encoded host name part of this web origin
//
//  Parameters:
//      pAceHost - receives the host in ASCII encoded format
//
//  Returns:
//      the host name
//===============================================================================================//
void WebOrigin::GetAceHost( String* pAceHost ) const
{
    if ( m_Host.GetLength() == 0 )
    {
        throw FunctionException( L"m_Host", __FUNCTION__ );
    }

    if ( pAceHost == nullptr )
    {
        throw NullException( L"pAceHost", __FUNCTION__ );
    }
    pAceHost->Zero();
    PXSIdnToAscii( m_Host, pAceHost );
}

//===============================================================================================//
//  Description:
//      Get the host name part of this web origin
//
//  Parameters:
//      none
//
//  Returns:
//      the host name
//===============================================================================================//
const String& WebOrigin::GetHost() const
{
    return m_Host;
}

//===============================================================================================//
//  Description:
//      Get the port number part of this web origin
//
//  Parameters:
//      none
//
//  Returns:
//      the host name
//===============================================================================================//
WORD WebOrigin::GetPort() const
{
    return m_port;
}

//===============================================================================================//
//  Description:
//      Get the scheme name part of this web origin
//
//  Parameters:
//      none
//
//  Returns:
//      the scheme name
//===============================================================================================//
const String& WebOrigin::GetScheme() const
{
    return m_Scheme;
}

//===============================================================================================//
//  Description:
//      Determine if the input is the same as this web origin
//
//  Parameters:
//      Origin - the web origin to test
//
//  Returns:
//      true if the same otherwise false
//===============================================================================================//
bool WebOrigin::IsSame( const WebOrigin& Origin ) const
{
    // Case sensitive
    if ( ( m_port == Origin.m_port ) &&
         ( m_Host.Compare( Origin.m_Host, true     ) == 0 ) &&
         ( m_Scheme.Compare( Origin.m_Scheme, true ) == 0 )  )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Determine if this web oring has valid parts
//
//  Parameters:
//      none
//
//  Returns:
//      true if the valid otherwise false
//===============================================================================================//
bool WebOrigin::IsValid() const
{
    // Just need to check for a valid port number otherwise would not have
    // been able to set the web origin
    if ( m_port )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Reset the members of this class
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WebOrigin::Reset()
{
    m_port = 0;
    m_Host.Zero();
    m_Scheme.Zero();
    m_DisplayString.Zero();
}

//===============================================================================================//
//  Description:
//      Set the members of this class
//
//  Parameters:
//      Scheme  - the scheme name
//      Host    - the host name, must be in Unicode normal form C
//      Port    - the port number
//
//  Returns:
//      void
//===============================================================================================//
void WebOrigin::Set( const String& Scheme, const String& Host, const String& Port )
{
    WORD   portNumber = 0;
    String HostFormC;

    if ( PXSIsRecognisedIriScheme( Scheme ) == false )
    {
        throw ParameterException( L"Scheme", __FUNCTION__ );
    }

    // The uri-host part of the web origin should be lowercase LDH, see RFC 6454
    // section 4. However, want to handle IPv6 addresses so store the host name in
    // Unicode normal form C. The AsCii Encoded version can be created when required.
    if ( Host.GetLength() == 0 )
    {
        throw ParameterException( L"Host", __FUNCTION__ );
    }
    HostFormC = Host;

    // Do this first before checking the host is normalised
    if ( PXSIsACEHostName( HostFormC ) )
    {
        PXSIdnToUnicode( &HostFormC );
    }

    if ( IsNormalizedString( NormalizationC, HostFormC.c_str(), -1 ) == FALSE )
    {
        throw ParameterException( L"NormalizationC", __FUNCTION__ );
    }

    if ( PXSIsValidPortNumber( Port, Scheme, &portNumber ) == false )
    {
        throw ParameterException( L"Port", __FUNCTION__ );
    }

    m_Scheme = Scheme;
    m_Host   = HostFormC;
    m_port   = portNumber;
    m_DisplayString.Zero();      // Lazy create on demand
}

//===============================================================================================//
//  Description:
//      Convert the web origin to a string, format is scheme://host:port
//
//  Parameters:
//      none
//
//  Returns:
//      reference to the web origin in string format
//===============================================================================================//
const String& WebOrigin::ToDisplayString()
{
    String    Host;
    Formatter Format;

    // Lazy create
    if ( m_DisplayString.GetLength() )
    {
        return m_DisplayString;
    }

    // Must have first set the web origin
    if ( m_Scheme.GetLength() == 0 )
    {
        throw FunctionException( L"m_Scheme", __FUNCTION__ );
    }
    m_DisplayString.Allocate( 256 );        // Usually enough
    m_DisplayString  = m_Scheme;
    m_DisplayString += L"://";
    m_DisplayString += m_Host;
    m_DisplayString += L":";
    m_DisplayString += Format.UInt16( m_port );

    return m_DisplayString;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
