///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Font Class Header
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

#ifndef PXSBASE_FONT_H_
#define PXSBASE_FONT_H_

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

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class Font
{
    public:
        // Default constructor
        Font();

        // Copy constructor
        Font( const Font& oFont );

        // Destructor
        ~Font();

        // Assignment operator
        Font& operator= ( const Font& oFont);

        // Methods
        void    Create();
        void    DecrementLogicalHeight();
        BYTE    GetCharacterSet();
        HFONT   GetHandle() const;
        void    GetLogFont( LOGFONT* pLogfFont ) const;
        int     GetSize( HDC hdc );
        bool    IsBold() const;
        void    IncrementLogicalHeight();
        bool    IsItalic() const;
        bool    IsUnderlined() const;
        void    SetBold( bool bold );
        void    SetCharacterSet( BYTE characterSet );
        void    SetItalic( bool italic );
        void    SetLogFont( const LOGFONT* pLogFont );
        void    SetFaceName( LPCWSTR pszFaceName );
        void    SetPointSize( int pointSize, HDC hdc );
        void    SetStockFont( int stockFont );
        void    SetUnderlined( bool underlined );

    protected:
        // Methods

        // Data members

    private:
        // Methods

        // Data members
        bool     m_bIsStockFont;
        HFONT    m_hFont;
        LOGFONTW m_LogFont;
};

#endif  // PXSBASE_FONT_H_
