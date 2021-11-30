///////////////////////////////////////////////////////////////////////////////////////////////////
//
// URI Class Implementation
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
#include "PxsBase/Header Files/URI.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateChars.h"
#include "PxsBase/Header Files/CharArray.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StringT.h"
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/UTF8.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
URI::URI()
    :m_Fragment(),
     m_Host(),
     m_Query(),
     m_Path(),
     m_Password(),
     m_Port(),
     m_Scheme(),
     m_User(),
     m_Uri()
{
}

// Copy constructor
URI::URI( const URI& oURI )
    :Object(),
     m_Fragment(),
     m_Host(),
     m_Query(),
     m_Path(),
     m_Password(),
     m_Port(),
     m_Scheme(),
     m_User(),
     m_Uri()
{
    *this = oURI;
}

// Destructor
URI::~URI()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
URI& URI::operator= ( const URI& oURI )
{
    if ( this == &oURI ) return *this;

     m_Fragment = oURI.m_Fragment;
     m_Host     = oURI.m_Host;
     m_Query    = oURI.m_Query;
     m_Path     = oURI.m_Path;
     m_Password = oURI.m_Password;
     m_Port     = oURI.m_Port;
     m_Scheme   = oURI.m_Scheme;
     m_User     = oURI.m_User;
     m_Uri      = oURI.m_Uri;

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      For file:// type Uri's, get the file path as used by the operating
//      system
//
//  Parameters:
//      Path - receives the file path
//
//  Remarks:
//      Caller must have already set the URI using Set()
//
//  Returns:
//      void
//===============================================================================================//
void URI::GetFilePath( String* pFilePath ) const
{
    // Scheme must be file://
    if ( m_Scheme.CompareI( PXS_STRING_SCHEME_FILE ) )
    {
        throw FunctionException( L"m_Scheme", __FUNCTION__ );
    }

    if ( pFilePath == nullptr )
    {
        throw ParameterException( L"pFilePath", __FUNCTION__ );
    }
    *pFilePath = m_Path;
    pFilePath->ReplaceChar( '/' , PXS_PATH_SEPARATOR );
    if ( pFilePath->StartsWith( L"\\\\",false ) == false )
    {
        // Not a UNC path so treat as a regular file path but may need to
        // strip the leading slash, see RFC 3986 section 3.3
        if ( ( pFilePath->GetLength() != 0 ) &&
             ( pFilePath->CharAt( 0 ) == PXS_PATH_SEPARATOR ) )
        {
            pFilePath->SetCharAt( 0, PXS_CHAR_SPACE );
            pFilePath->Trim();
        }
    }
}

//===============================================================================================//
//  Description:
//      Return the URI's fragment
//
//  Parameters:
//      None
//
//  Returns:
//      ref to the fragment
//===============================================================================================//
const String& URI::GetFragment() const
{
    return m_Fragment;
}

//===============================================================================================//
//  Description:
//      Return the URI's host
//
//  Parameters:
//      None
//
//  Returns:
//      ref to the host
//===============================================================================================//
const String& URI::GetHost() const
{
    return m_Host;
}

//===============================================================================================//
//  Description:
//      Get the normalised form of this URI with any special characters percent
//      encoded
//
//  Parameters:
//      pScheme   - optional, receives the normalised scheme
//      pUser     - optional, receives the normalised user
//      pPassword - optional, receives the normalised password
//      pHost     - optional, receives the normalised host
//      pPort     - optional, receives the normalised port
//      pPath     - optional, receives the normalised path
//      pQuery    - optional, receives the normalised query
//      pFragment - optional, receives the normalised fragment
//
//  Returns:
//      void
//===============================================================================================//
void URI::GetNormalised( CharArray* pScheme,
                         CharArray* pUser,
                         CharArray* pPassword,
                         CharArray* pHost,
                         CharArray* pPort,
                         CharArray* pPath, CharArray* pQuery, CharArray* pFragment ) const
{
    UTF8   Utf8Data;
    String Data;

    if ( pScheme )
    {
        Data = PXS_STRING_EMPTY;
        PercentEncode( m_Scheme, PXS_STRING_EMPTY, &Data );
        Utf8Data.Set( Data.c_str() );
        pScheme->Free();
        pScheme->Append( Utf8Data.u_str(), Utf8Data.GetByteLength() );
    }

    if ( pUser )
    {
        Data = PXS_STRING_EMPTY;
        PercentEncode( m_User, PXS_STRING_EMPTY, &Data );
        Utf8Data.Set( Data.c_str() );
        pUser->Free();
        pUser->Append( Utf8Data.u_str(), Utf8Data.GetByteLength() );
    }

    if ( pPassword )
    {
        Data = PXS_STRING_EMPTY;
        PercentEncode( m_Password, PXS_STRING_EMPTY, &Data );
        Utf8Data.Set( Data.c_str() );
        pPassword->Free();
        pPassword->Append( Utf8Data.u_str(), Utf8Data.GetByteLength() );
    }

    if ( pHost )
    {
        Data = PXS_STRING_EMPTY;
        PercentEncode( m_Host, PXS_STRING_EMPTY, &Data );
        Utf8Data.Set( Data.c_str() );
        pHost->Free();
        pHost->Append( Utf8Data.u_str(), Utf8Data.GetByteLength() );
    }

    if ( pPort )
    {
        Data = PXS_STRING_EMPTY;
        PercentEncode( m_Port, PXS_STRING_EMPTY, &Data );
        Utf8Data.Set( Data.c_str() );
        pPort->Free();
        pPort->Append( Utf8Data.u_str(), Utf8Data.GetByteLength() );
    }

    if ( pPath )
    {
        Data = PXS_STRING_EMPTY;
        PercentEncode( m_Path, L"/", &Data );
        Utf8Data.Set( Data.c_str() );
        pPath->Free();
        pPath->Append( Utf8Data.u_str(), Utf8Data.GetByteLength() );
    }

    if ( pQuery )
    {
        // If have a query will pre-prend the "?" delimiter as its possible to
        // have a bare "?". For example "http://example.com/?" is not
        // the same as "http://example.com/". See RFC 3986 section 6.2.3.
        // Scheme-Based Normalization. In other words if there is a bare "?"
        // then that is the query otherwise leave pQuery as zero length.
        pQuery->Free();
        if ( m_Query.c_str() )
        {
            pQuery->Append( "?", 1 );
            Data = PXS_STRING_EMPTY;
            PercentEncode( m_Query, PXS_STRING_EMPTY, &Data );
            Utf8Data.Set( Data.c_str() );
            pQuery->Append( Utf8Data.u_str(), Utf8Data.GetByteLength() );
        }
    }

    if ( pFragment )
    {
        Data = PXS_STRING_EMPTY;
        PercentEncode( m_Fragment, PXS_STRING_EMPTY, &Data );
        Utf8Data.Set( Data.c_str() );
        pFragment->Free();
        pFragment->Append( Utf8Data.u_str(), Utf8Data.GetByteLength() );
    }
}

//===============================================================================================//
//  Description:
//      Return the URI's query part
//
//  Parameters:
//      None
//
//  Returns:
//      ref to the query
//===============================================================================================//
const String& URI::GetQuery() const
{
    return m_Query;
}

//===============================================================================================//
//  Description:
//      Return the URI's path
//
//  Parameters:
//      None
//
//  Returns:
//      ref to the path
//===============================================================================================//
const String& URI::GetPath() const
{
    return m_Path;
}

//===============================================================================================//
//  Description:
//      Return the URI's password
//
//  Parameters:
//      None
//
//  Returns:
//      ref to the password
//===============================================================================================//
const String& URI::GetPassword() const
{
    return m_Password;
}

//===============================================================================================//
//  Description:
//      Return the URI's port
//
//  Parameters:
//      None
//
//  Returns:
//      ref to the port
///===============================================================================================//
const String& URI::GetPort() const
{
    return m_Port;
}

//===============================================================================================//
//  Description:
//      Return the URI's scheme
//
//  Parameters:
//      None
//
//  Returns:
//      ref to the scheme
//===============================================================================================//
const String& URI::GetScheme() const
{
    return m_Scheme;
}

//===============================================================================================//
//  Description:
//      Get the URI
//
//  Parameters:
//      None
//
//  Remarks
//      The is the "fixed-up" strign suitable to display to user
//
//  Returns:
//      ref to the URI
//===============================================================================================//
const String& URI::GetUri()
{
    // Lazy create
    if ( m_Uri.GetLength() )
    {
        return m_Uri;  // Already done it
    }
    MakeUri( m_Scheme, m_User, m_Password, m_Host, m_Port, m_Path, m_Query, &m_Fragment, &m_Uri );

    return m_Uri;
}

//===============================================================================================//
//  Description:
//      Get the normalised form of this URI with any special characters percent
//      encoded
//
//  Parameters:
//      fragment - true if want the fragment part of the URI
//      pUriNormalised - receives the normalised URI
//
//  Remarks
//      The is the "fixed-up" string suitable for byte gy byte comparison
//      with other URI's. See RFC3986 Section 6. Normalization and Comparison
//
//  Returns:
//      void
//===============================================================================================//
void URI::GetUriNormalised( bool fragment, UTF8* pUriNormalised ) const
{
    String Uri, Scheme, User, Password, Host, Port, Path, Query, Fragment;

    if ( pUriNormalised == nullptr )
    {
        throw ParameterException( L"pUriNormalised", __FUNCTION__ );
    }
    PercentEncode( m_Scheme  , PXS_STRING_EMPTY, &Scheme   );
    PercentEncode( m_User    , PXS_STRING_EMPTY, &User     );
    PercentEncode( m_Password, PXS_STRING_EMPTY, &Password );
    PercentEncode( m_Host    , PXS_STRING_EMPTY, &Host     );
    PercentEncode( m_Port    , PXS_STRING_EMPTY, &Port     );
    PercentEncode( m_Path    , L"/",             &Path     );
    PercentEncode( m_Query   , PXS_STRING_EMPTY, &Query    );
    if ( fragment )
    {
        PercentEncode( m_Fragment, PXS_STRING_EMPTY, &Fragment );
    }
    MakeUri( Scheme, User, Password, Host, Port, Path, Query, &Fragment, &Uri );
    pUriNormalised->Set( Uri.c_str() );
}

//===============================================================================================//
//  Description:
//      Return the URI's user name
//
//  Parameters:
//      None
//
//  Returns:
//      ref to the user name
//===============================================================================================//
const String& URI::GetUser() const
{
    return m_User;
}

//===============================================================================================//
//  Description:
//      Determine if the URI's syntax is valid
//
//  Parameters:
//      Uri - the URI to test
//
//  Returns:
//      true if valid otherwise false
//===============================================================================================//
bool URI::IsValidUri( const String& Uri ) const
{
    bool   isValid = true;
    String UriCopy, Fragment, Host, Query, Path, Password, Port, Scheme, User;

    if ( Uri.GetLength() == 0 )
    {
        return false;
    }

    try
    {
        UriCopy = Uri;
        UriCopy.Trim();
        FixupSlashes( &UriCopy );
        ParseUriRfc3986( UriCopy,
                         &Scheme, &User, &Password, &Host, &Port, &Path, &Query, &Fragment );
        
        // If have a scheme it must be recognised
        if ( Scheme.GetLength() )
        {
            if ( PXSIsRecognisedUriScheme( Scheme ) == false )
            {
                return false;
            }
        }

        // Although both are optional, want either a host or a path
        if ( ( Host.GetLength() == 0 ) && ( Path.GetLength() == 0 ) )
        {
            return false;
        }

        // If there is a port, it must be a number
        if ( Port.GetLength() )
        {
            isValid = PXSIsStringOnlyDigits( Port.c_str() );
        }
    }
    catch ( Exception& e )
    {
        isValid = false;
        PXSLogException( e, __FUNCTION__ );
    }

    return isValid;
}

//===============================================================================================//
//  Description:
//      Reset the class members
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void URI::Reset()
{
     m_Fragment = PXS_STRING_EMPTY;
     m_Host     = PXS_STRING_EMPTY;
     m_Query    = PXS_STRING_EMPTY;
     m_Path     = PXS_STRING_EMPTY;
     m_Password = PXS_STRING_EMPTY;
     m_Port     = PXS_STRING_EMPTY;
     m_Scheme   = PXS_STRING_EMPTY;
     m_User     = PXS_STRING_EMPTY;
     m_Uri      = PXS_STRING_EMPTY;
}

//===============================================================================================//
//  Description:
//      Set this Uri propety
//
//  Parameters:
//      Uri - the URI
//
//  Returns:
//      void
//===============================================================================================//
void URI::Set( const String& Uri )
{
    String UriFixedUp, Scheme, User, Password, Host, Port, Path, Query, Fragment;

    Reset();
    UriFixedUp = Uri;
    UriFixedUp.Trim();
    FixupSlashes( &UriFixedUp );
    ParseUriRfc3986( UriFixedUp, &Scheme, &User, &Password, &Host, &Port, &Path, &Query, &Fragment );
    PercentDecode( Scheme  , &m_Scheme   );
    PercentDecode( User    , &m_User     );
    PercentDecode( Password, &m_Password );
    PercentDecode( Host    , &m_Host     );
    PercentDecode( Port    , &m_Port     );
    PercentDecode( Path    , &m_Path     );
    PercentDecode( Query   , &m_Query    );
    PercentDecode( Fragment, &m_Fragment );

    // Set defaults
    if ( m_Scheme.IsEmpty() )
    {
        if ( IsLocalOrUncFilePath( UriFixedUp ) )
        {
            m_Scheme = PXS_STRING_SCHEME_FILE;
        }
        else
        {
            if ( m_Port.CompareI( PXS_STRING_HTTPS_DEFAULT_PORT ) == 0 )
            {
                m_Scheme = PXS_STRING_SCHEME_HTTPS;
            }
            else
            {
                m_Scheme = PXS_STRING_SCHEME_HTTP;
            }
        }
    }

    if ( m_Port.IsEmpty() )
    {
        if ( m_Scheme.CompareI( PXS_STRING_SCHEME_HTTP ) == 0 )
        {
            m_Port = PXS_STRING_HTTP_DEFAULT_PORT;
        }
        else if ( m_Scheme.CompareI( PXS_STRING_SCHEME_HTTPS ) == 0 )
        {
            m_Port = PXS_STRING_HTTPS_DEFAULT_PORT;
        }
    }

    // Set Case, see RFC RFC 3968, section 3.1. etc.
    m_Scheme.ToLowerCase();
    m_Host.ToLowerCase();
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Replace any backslashes with forward slashes
//
//  Parameters:
//      pUri - the string Uri to fixup
//
//  Remarks:
//      Only replaces thos upto the path end, e.g. http:\\ to http://
//
//  Returns:
//      void
//===============================================================================================//
void URI::FixupSlashes( String* pUri ) const
{
    size_t i = 0, length, idxQuery, idxFragment, idxPathEnd;

    if ( pUri == nullptr )
    {
        throw ParameterException( L"pUri", __FUNCTION__ );
    }

    length = pUri->GetLength();
    if ( length == 0 )
    {
        return;     // Nothing to do
    }

    idxQuery     = pUri->IndexOf( '?', 0 );
    idxFragment  = pUri->IndexOf( '#', 0 );
    if ( idxQuery > idxFragment )
    {
        idxQuery = PXS_MINUS_ONE;   // "?" is not a query
    }
    idxPathEnd = PXSMinSizeT( idxQuery, idxFragment );

    idxPathEnd = PXSMinSizeT( idxPathEnd, ( length - 1 ) ); // length > 0 from above
    for ( i = 0; i <= idxPathEnd; i++ )
    {
        if ( pUri->CharAt( i ) == '\\' )
        {
            pUri->ReplaceCharAt( i, '/' );
        }
    }
}

//===============================================================================================//
//  Description:
//      Determine if the specified path is a local or UNC file path
//
//  Parameters:
//
//  Remarks:
//      Do this test AFTER fixing up the back slashes to forward slashes
//      Does not check for the existence of the file.
//
//  Returns:
//      true if the path is valid for local or UNC file path, otherwise false
//===============================================================================================//
bool URI::IsLocalOrUncFilePath( const String& Path ) const
{
    wchar_t ch0 = PXS_CHAR_NULL, ch1 = PXS_CHAR_NULL, ch2 = PXS_CHAR_NULL;
    const wchar_t* pzServer = nullptr;
    const wchar_t* pzShare  = nullptr;
    String      Temp, Drive, Dir, Fname, Ext;
    StringArray Tokens;
    
    // Test the scheme
    Temp  = PXS_STRING_SCHEME_FILE;
    Temp += L":";
    if ( Path.StartsWith( Temp.c_str(), false ) )
    {
        return true;
    }

    // Test for the start of regular file path, e.g. "C:/"
    if ( Path.GetLength() > 3 )
    {
        ch0 = Path.CharAt( 0 );
        ch1 = Path.CharAt( 1 );
        ch2 = Path.CharAt( 2 );
        if ( PXSIsLogicalDriveLetter( ch0 ) &&
             ( ch1 == ':' ) &&
             ( ch2 == '/' )  )
        {
            return true;
        }
    }

    // Test for UNC, i.e. "\\<server name>\<share point name>"
    if ( Path.StartsWith( L"//", false ) )
    {
        Path.ToArray( '/', &Tokens );
        if ( Tokens.GetSize() >= 4 )
        {
            pzServer = Tokens.Get( 2 );    // <server name>
            pzShare  = Tokens.Get( 3 );    // <share point name>
            if ( ( pzServer && *pzServer ) &&
                 ( pzShare  && *pzShare  )  )
            {
                return true;
            }
        }
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Make tthe URI using the supplied components
//
//  Parameters:
//      Scheme    - the scheme
//      User      - the authority user name
//      Password  - the authority password
//      Host      - the host name/address
//      Port      - the port number
//      Path      - the path
//      Query     - the query component
//      pFragment - option fragment component, use NULL if not wanted
//      pUri      - recieve the URI
//
//  Remarks:
//      See RFC 3986 Section 6;
//
//  Returns:
//      void
//===============================================================================================//
void URI::MakeUri( const String& Scheme,
                   const String& User,
                   const String& Password,
                   const String& Host, 
                   const String& Port, 
                   const String& Path, 
                   const String& Query, const String* pFragment, String* pUri ) const
{
    if ( pUri == nullptr )
    {
        throw ParameterException( L"pUri", __FUNCTION__ );
    }
    pUri->Allocate( 512 );      // Minimise repeated new allocs
    *pUri = PXS_STRING_EMPTY;

    // 3.1.  Scheme
    *pUri = Scheme;
    if ( pUri->IsEmpty() )
    {
        *pUri = PXS_STRING_SCHEME_HTTP;
    }
    pUri->ToLowerCase();
    *pUri += L"://";

    // 3.2.1.  User Information
    // Use of the format "user:password" in the userinfo field is deprecated.
    if ( User.GetLength() || Password.GetLength() )
    {
        PXSLogAppWarn( L"User name and password in a URI are deprecated. " \
                       L"These have been removed from the URI." );
    }

    // 3.2.2&3.  Host & Port. Only need the port if its a non-default value
    if ( Host.GetLength() )
    {
        *pUri += Host;
        if ( Port.GetLength() )
        {
            // http port
            if ( ( Scheme.CompareI( PXS_STRING_SCHEME_HTTP     ) == 0 ) &&
                 ( Port.CompareI( PXS_STRING_HTTP_DEFAULT_PORT ) != 0 )  )
            {
                *pUri += L":";
                *pUri += Port;
            }

            // https port
            if ( ( Scheme.CompareI( PXS_STRING_SCHEME_HTTPS     ) == 0 ) &&
                 ( Port.CompareI( PXS_STRING_HTTPS_DEFAULT_PORT ) != 0 )  )
            {
                *pUri += L":";
                *pUri += Port;
            }
        }
    }

    // 6.2.3.  Scheme-Based Normalization: "...In general, a URI that uses the
    // generic syntax for authority with an empty path should be normalized to
    // a path of "/".
    if ( Path.StartsWith( L"/", false  ) == false )
    {
        *pUri += '/';
    }

    // 3.3.  Path
    *pUri += Path;

    // 3.4.  Query
    // If a query is present, it already begins with the leading "?" as according
    // to 6.2.3. Scheme-Based Normalization "...the URI "http://example.com/?"
    // cannot be assumed to be equivalent to any of the examples above...". In
    // other words "?" is the query string
    if ( Query.IsNull() == false )
    {
        *pUri += '?';
        *pUri += Query;
    }

    // 3.5.  Fragment
    if ( pFragment && pFragment->GetLength() )
    {
        *pUri += '#';
        *pUri += *pFragment;
    }
}

//===============================================================================================//
//  Description:
//      Split the Authority part of a URI into its components
//
//  Parameters:
//      Authority - the URI's authority part
//      pUser     - receives the user part of the Authority string
//      pUser     - receives the password part of the Authority string
//      pHost     - receives the host part of the Authority string
//      pPort     - receives the port number part of the Authority string
//
//  Remarks:
//      See RFC3986
//      Authority = user:pass@host:port
//
//  Returns:
//      void
//===============================================================================================//
void URI::ParseAuthorityRfc3986( const String& Authority,
                                 String* pUser,
                                 String* pPassword, String* pHost, String* pPort ) const
{
    size_t  idxUserInfo, idxHost, idxColon; 
    String  UserInfo;

    if ( ( pUser     == nullptr ) ||
         ( pPassword == nullptr ) ||
         ( pHost     == nullptr ) ||
         ( pPort     == nullptr )  )
    {
        throw ParameterException( L"Out parameter", __FUNCTION__ );
    }
    *pUser     = PXS_STRING_EMPTY;
    *pPassword = PXS_STRING_EMPTY;
    *pHost     = PXS_STRING_EMPTY;
    *pPort     = PXS_STRING_EMPTY;

    // UserInfo
    idxHost     = 0;
    idxUserInfo = Authority.IndexOf( '@', 0 );
    if ( idxUserInfo != PXS_MINUS_ONE )
    {
        idxHost = idxUserInfo + 1;
        Authority.SubString( 0, idxUserInfo, &UserInfo );
        UserInfo.Trim();
        idxColon = UserInfo.IndexOf( ':', 0 );
        if ( idxColon != PXS_MINUS_ONE )
        {
            UserInfo.SubString( 0, idxColon, pUser );
            UserInfo.SubString( idxColon + 1, PXS_MINUS_ONE, pPassword );
        }
        else
        {
            // User only
            UserInfo.SubString( 0, idxUserInfo, pUser );
        }
    }

    // Port
    idxColon = Authority.IndexOf( ':', idxHost );
    if ( idxColon != PXS_MINUS_ONE )
    {
        Authority.SubString( idxColon + 1, PXS_MINUS_ONE, pPort );
        pPort->Trim();
    }

    // Host
    if ( ( idxUserInfo != PXS_MINUS_ONE ) && ( idxColon != PXS_MINUS_ONE ) )
    {
        // <userinfo>@<host>:<port>
        if ( idxColon > idxUserInfo )
        {
            Authority.SubString( idxUserInfo + 1, idxColon - idxUserInfo - 1, pHost );
        }
    }
    else if ( ( idxUserInfo != PXS_MINUS_ONE ) && ( idxColon == PXS_MINUS_ONE ) )
    {
        // <userinfo>@<host>
        Authority.SubString( idxUserInfo + 1, PXS_MINUS_ONE, pHost );
    }
    else if ( ( idxUserInfo == PXS_MINUS_ONE ) && ( idxColon != PXS_MINUS_ONE ) )
    {
        // <host>:<port>
        Authority.SubString( 0, idxColon, pHost );
    }
    else
    {
        // <host>
        *pHost = Authority;
    }
    pHost->Trim();
}

//===============================================================================================//
//  Description:
//      Split the specified URI into its components
//
//  Parameters:
//      Uri       - the URI to split
//      pScheme   - receives the scheme
//      pUserinfo - receives the user
//      pPassword - receives the password
//      pHost     - receives the host
//      pPort     - receives the port
//      pPath     - receives the abs_path
//      pQuery    - receives the query
//      pFragment - receives the fragment
//
//  Remarks:
//      See RFC3986
//      This method makes the simplifying assumption that the URI is using
//      a Server-based rather than a Registry-based Naming Authority.
//      scheme:[//<........authority.........>[/]path[?query][#fragment]
//      scheme:[//[user:password@]host[:port]][/]path[?query][#fragment]
//
//      The input must already have fixed up the slashes otherwise the
//      parsing will be unpredictable.
//
//      "#" is not allowed anywhere except escaped so it is always the start
//      of the fragment. "?" is not allowed anywhere before the query part
//      so if found before "#", it is the start of the query part.
//
//  Returns:
//      void
//===============================================================================================//
void URI::ParseUriRfc3986( const String& Uri,
                           String* pScheme,
                           String* pUser,
                           String* pPassword,
                           String* pHost,
                           String* pPort, String* pPath, String* pQuery, String* pFragment ) const
{
    size_t idxQuery, idxSchemeEnd, idxAuthority;
    size_t idxAuthorityEnd, idxPath, idxPathEnd, idxFragment;
    String Authority;

    if ( ( pScheme   == nullptr ) ||
         ( pUser     == nullptr ) ||
         ( pPassword == nullptr ) ||
         ( pHost     == nullptr ) ||
         ( pPort     == nullptr ) ||
         ( pPath     == nullptr ) ||
         ( pQuery    == nullptr ) ||
         ( pFragment == nullptr )  )
    {
        throw ParameterException( L"out parameter", __FUNCTION__ );
    }
    *pScheme   = PXS_STRING_EMPTY;
    *pUser     = PXS_STRING_EMPTY;
    *pPassword = PXS_STRING_EMPTY;
    *pHost     = PXS_STRING_EMPTY;
    *pPort     = PXS_STRING_EMPTY;
    *pPath     = PXS_STRING_EMPTY;
    *pQuery    = PXS_STRING_EMPTY;
    *pFragment = PXS_STRING_EMPTY;

    // See RFC3986 5.2.2.  Transform References. Will first parse into
    // R.scheme, R.authority, R.path, R.query, R.fragment then further
    // parse R.authority

    if ( Uri.GetLength() == 0 )
    {
        return;     // Nothing to do
    }
    idxQuery     = Uri.IndexOf( '?', 0 );
    idxFragment  = Uri.IndexOf( '#', 0 );
    if ( idxQuery > idxFragment )
    {
        idxQuery = PXS_MINUS_ONE;   // "?" is not a query
    }
    idxPathEnd = PXSMinSizeT( idxQuery, idxFragment );

    // "://" is a valid string in the fragment or query so
    // if found must be before the path end
    idxSchemeEnd = Uri.IndexOf( L"://" );
    if ( ( idxSchemeEnd == PXS_MINUS_ONE ) ||
         ( idxSchemeEnd >  idxPathEnd    )  )
    {
        idxSchemeEnd = 0;   // i.e. no scheme
    }

    // Authority and Path
    // RFC 3986 section 3.2: "The authority component is preceded by a double"
    // "slash ("//") and is terminated by the next slash ("/"), question mark"
    // "("?"), or number sign ("#") character, or by the end of the URI.
    // Note, will assume file paths do not have an Authority part, just a path.
    // For UNC paths, the server name could be said to equate to a URI host
    // but as windows file API's take UNC paths will assume there is no host.
    idxAuthority    = 0;
    idxAuthorityEnd = 0;
    idxPath         = 0;
    if ( IsLocalOrUncFilePath( Uri ) )
    {
        if ( idxSchemeEnd )
        {
            idxAuthority    = idxSchemeEnd + 3;  // Length of "://"
            idxAuthorityEnd = idxSchemeEnd + 3;
            idxPath         = idxSchemeEnd + 3;
        }
    }
    else
    {
        // Authority
        if ( idxSchemeEnd )
        {
            idxAuthority = idxSchemeEnd + 3;  // Length of "://"
        }
        idxAuthorityEnd = Uri.IndexOf( '/', idxAuthority );
        idxAuthorityEnd = PXSMinSizeT( idxAuthorityEnd, idxPathEnd );
        
        // Path
        idxPath = Uri.IndexOf( '/', idxAuthorityEnd );
        if ( idxPath > idxPathEnd )
        {
            idxPath = PXS_MINUS_ONE;    // No path
        }
    }

    // Extract the strings
    Uri.SubString( 0, idxSchemeEnd, pScheme );
    pScheme->Trim();

    if ( idxAuthorityEnd > idxAuthority )
    {
        Uri.SubString( idxAuthority, idxAuthorityEnd - idxAuthority, &Authority );
        Authority.Trim();
        ParseAuthorityRfc3986( Authority, pUser, pPassword, pHost, pPort );
    }

    if ( idxPathEnd > idxPath )
    {
        Uri.SubString( idxPath, idxPathEnd - idxPath, pPath );
        pPath->Trim();
    }

    // The "?" should be treated as part of the query even if the rest
    // of the query is an empty string as "http://example.com/?" is not
    // the same as "http://example.com/". See RFC 3986 section 6.2.3.
    // Scheme-Based Normalization. In other words if there is a bare "?"
    // then that is the query
    if ( idxQuery == PXS_MINUS_ONE )
    {
        pQuery->Empty();
    }
    else
    {
       if ( idxFragment > idxQuery )
        {
            Uri.SubString( idxQuery + 1, idxFragment - idxQuery - 1, pQuery );
        }
        else
        {
            Uri.SubString( idxQuery + 1, PXS_MINUS_ONE, pQuery );
        }
        pQuery->Trim();
    }

    if ( idxFragment != PXS_MINUS_ONE )
    {
        Uri.SubString( idxFragment + 1, PXS_MINUS_ONE, pFragment );
        pFragment->Trim();
    }
}

//===============================================================================================//
//  Description:
//      Percent decode the input, e.g. %20 = space char
//
//  Parameters:
//      Encoded  - the encoded string
//      pDecoded - receives the decoded string
//
//  Returns:
//      character
//===============================================================================================//
void URI::PercentDecode( const String& Encoded, String* pDecoded )
{
    char*   pszUtf8 = nullptr;
    size_t  i  = 0, length;
    wchar_t ch = 0, ch1 = 0, ch2 = 0;
    String  Temp;
    Formatter     Format;
    AllocateChars Alloc;

    if ( pDecoded == nullptr )
    {
        throw ParameterException( L"pDecoded", __FUNCTION__ );
    }
    *pDecoded = PXS_STRING_EMPTY;

    length = Encoded.GetLength();
    if ( ( length <= 2 ) ||
         ( Encoded.IndexOf ( '%', 0 ) == PXS_MINUS_ONE ) )
    {
        // No percent encoded characters
        *pDecoded = Encoded;
        return;
    }

    // See 2.4.  When to Encode or Decode
    // Replace any percent encoded chars with their characters. The result
    // should be a string og UTF-8 characters
    Temp.Allocate( PXSAddSizeT( length, 1 ) );
    for ( i = 0; i < length; i++ )
    {
        ch = Encoded.CharAt( i );
        if ( ( ch == '%' ) && i < ( length - 2 ) )
        {
            ch1 = Encoded.CharAt( i + 1 );
            ch2 = Encoded.CharAt( i + 2 );
            if ( PXSIsHexit( ch1 ) && PXSIsHexit( ch2 ) )
            {
                ch = PXSHexDigitsToChar( ch1, ch2 );
                i +=2;
            }
        }
        Temp += ch;
    }

    // After decoding the string should normally be composed of single byte
    // characters in UTF-8 format otherwise leave as is because do not alawys
    // have contol of the input. See 2.5.  Identifying Data "...data should
    // first be encoded as octets according to the UTF-8 character encoding
    // [STD63]; then only those octets that do not correspond to characters
    // in the unreserved set should be percent-encoded."
    if ( Temp.IsUtf8() )
    {
        length  = Temp.GetLength();
        pszUtf8 = Alloc.New( PXSAddSizeT( length, 1 ) );
        for ( i = 0; i < length; i++ )
        {
            pszUtf8[ i ] = static_cast<char>( 0xff & Temp.CharAt( i ) );
        }
        pszUtf8[ length ] = '\0';
        *pDecoded = Format.UTF8ToWide( pszUtf8 );
    }
    else
    {
        *pDecoded = Temp;
    }
}

//===============================================================================================//
//  Description:
//      Percent encode the input URI, e.g. space char = %20
//
//  Parameters:
//      Decoded    - the decoded string
//      pwzExclude - exclude encoding any characters in the string
//      pEncoded   - receives the encoded string
//
//  Returns:
//      character
//===============================================================================================//
void URI::PercentEncode( const String& Decoded,
                         const wchar_t* pwzExclude, String* pEncoded ) const
{
    UTF8    Utf8;
    size_t  i   = 0, length;
    wchar_t wch = 0;
    wchar_t hexit[ 16 ] = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9',
                            'A', 'B', 'C', 'D', 'E', 'F' };

    if ( ( pwzExclude == nullptr ) || ( pEncoded == nullptr ) )
    {
        throw ParameterException( L"pwzExclude/pEncoded", __FUNCTION__ );
    }

    // If the input is NULL preserve it
    if ( Decoded.c_str() == nullptr )
    {
        pEncoded->Empty();
        return;
    }
    pEncoded->Allocate( 512 );      // Mininise new allocations
    *pEncoded = PXS_STRING_EMPTY;

    // Convert to UTF-8
    Utf8.Set( Decoded.c_str() );

    // Percent Encode
    length = Utf8.GetByteLength();
    for ( i = 0; i < length; i++ )
    {
        wch = PXSCastCharToWChar( Utf8.GetAt( i ) );
        if ( wcschr( pwzExclude, wch ) ||
             wcschr( PXS_STRING_URI_UNRESERVED, wch ) )
        {
            *pEncoded += wch;
        }
        else
        {
            *pEncoded += '%';
            *pEncoded += hexit[ wch / 16 ];
            *pEncoded += hexit[ wch % 16 ];
        }
    }
}
