///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Label Class Implementation
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
#include "PxsBase/Header Files/Label.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
Label::Label()
{
    // Class registration
    m_WndClassEx.style         = CS_PARENTDC | CS_DBLCLKS;
    m_WndClassEx.lpszClassName = L"STATIC";

    // Creation parameters
    m_CreateStruct.cy          = 25;
    m_CreateStruct.cx          = 80;
    m_CreateStruct.style       = WS_VISIBLE | WS_CHILD | SS_LEFT;
    m_CreateStruct.lpszClass   = m_WndClassEx.lpszClassName;
    m_CreateStruct.dwExStyle   = WS_EX_TRANSPARENT;
}

// Copy constructor - not allowed so no implementation

// Destructor
Label::~Label()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator - not allowed so no implementation

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get this window's preferred size
//
//  Parameters:
//      pSize - receives the size
//
//  Returns:
//      void
//===============================================================================================//
void Label::GetPreferredSize( SIZE* pSize ) const
{
    int     cbString   = 0;
    HDC     hdc        = nullptr;
    HFONT   hFont      = nullptr;
    HGDIOBJ oldFont    = nullptr;
    LPCWSTR pszCaption = nullptr;
    String  Text;

    if ( m_hWindow == nullptr )
    {
        return;
    }

    if ( pSize == nullptr )
    {
        return;
    }
    pSize->cx = 16;
    pSize->cy = 16;

    GetText( &Text );
    pszCaption = Text.c_str();
    if ( pszCaption )
    {
        cbString = lstrlen( pszCaption );
        hdc = GetDC( m_hWindow );
        if ( hdc )
        {
            hFont   = (HFONT)SendMessage( m_hWindow, WM_GETFONT, 0, 0 );
            oldFont = SelectObject( hdc, hFont );
            GetTextExtentPoint32( hdc, pszCaption, cbString, pSize );
            if ( oldFont )
            {
                SelectObject( hdc, oldFont );
            }
            ReleaseDC( m_hWindow, hdc );
        }
        pSize->cx += 2;
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
