///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Menu Item Class Header
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

#ifndef PXSBASE_MENU_ITEM_H_
#define PXSBASE_MENU_ITEM_H_

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
#include "PxsBase/Header Files/Menu.h"
#include "PxsBase/Header Files/StringT.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class MenuItem : public Menu
{
    public:
        // Default constructor
        MenuItem();

        // Destructor
        ~MenuItem();

        // Methods
        HBITMAP GetBitmap() const;
        UINT    GetCommandID() const;
  const String& GetLabel() const;
        HMENU   GetParent() const;
        bool    IsChecked() const;
        bool    IsEnabled() const;
        bool    IsHyperLink() const;
        bool    IsMenuBarItem() const;
        void    SetBitmap( WORD resourceID );
        void    SetChecked( bool checked );
        void    SetCommandID( WORD wommandID );
        void    SetEnabled( bool enabled );
        void    SetFilledBitmap( COLORREF fill, COLORREF border );
        void    SetHyperLink( bool hyperLink );
        void    SetLabel( const String& Label );
        void    SetMenuBarItem( bool menuBarItem );
        void    SetParent( HMENU hMenuParent );

    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        MenuItem( const MenuItem& oMenuItem );

        // Assignment operator - not allowed
        MenuItem& operator= ( const MenuItem& oMenuItem );

        // Methods

        // Data members
        bool        m_bHyperLink;
        bool        m_bMenuBarItem;
        UINT        m_uCommandID;
        HMENU       m_hMenuParent;
        HBITMAP     m_hBitmap;
        String      m_Label;
};

#endif  // PXSBASE_MENU_ITEM_H_
