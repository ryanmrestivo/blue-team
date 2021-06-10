///////////////////////////////////////////////////////////////////////////////////////////////////
//
// URI Class Header
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

#ifndef PXSBASE__URI_H_
#define PXSBASE__URI_H_

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
#include "PxsBase/Header Files/Object.h"
#include "PxsBase/Header Files/StringT.h"

// 6. Forwards
class UTF8;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class URI : public Object
{
    public:
        // Default constructor
        URI();

        // Copy constructor
        URI( const URI& oURI );

        // Assignment operator
        URI& operator= ( const URI& oURI );

        // Destructor
        ~URI();

        // Methods
        void          GetFilePath( String* pFilePath ) const;
        const String& GetFragment() const;
        const String& GetHost() const;
        const String& GetQuery() const;
        void          GetNormalised( CharArray* pScheme,
                                     CharArray* pUser,
                                     CharArray* pPassword,
                                     CharArray* pHost,
                                     CharArray* pPort,
                                     CharArray* pPath,
                                     CharArray* pQuery, CharArray* pFragment ) const;
        const String& GetPath() const;
        const String& GetPassword() const;
        const String& GetPort() const;
        const String& GetScheme() const;
        const String& GetUri();
        void          GetUriNormalised( bool fragment, UTF8* pUriNormalised ) const;
        const String& GetUser() const;
        bool          IsValidUri( const String& Uri ) const;
        void          Reset();
        void          Set( const String& Uri );

    protected:
        // Methods

        // Data members

    private:
        // Methods
        void FixupSlashes( String* pUri ) const;
        bool IsLocalOrUncFilePath( const String& Path ) const;
        void MakeUri( const String& Scheme,
                      const String& User,
                      const String& Password,
                      const String& Host, 
                      const String& Port, 
                      const String& Path, 
                      const String& Query, const String* pFragment, String* pUri ) const;
        void ParseAuthorityRfc3986( const String& Authority,
                                    String* pUser,
                                    String* pPassword, String* pHost, String* pPort ) const;
        void ParseUriRfc3986( const String& URI,
                              String* pScheme,
                              String* pUser,
                              String* pPassword,
                              String* pHost,
                              String* pPort,
                              String* pPath, String* pQuery, String* pFragment) const;
        void PercentDecode( const String& Encoded, String* pDecoded );
        void PercentEncode( const String& Decoded,
                            const wchar_t* pwzExclude, String* pEncoded ) const;

        // Data members
        String  m_Fragment;
        String  m_Host;
        String  m_Query;
        String  m_Path;
        String  m_Password;
        String  m_Port;
        String  m_Scheme;
        String  m_User;
        String  m_Uri;
};

#endif  // PXSBASE_URI_H_
