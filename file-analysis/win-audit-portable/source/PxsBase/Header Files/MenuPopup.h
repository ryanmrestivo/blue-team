///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Menu Popup Class Header
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

#ifndef PXSBASE_MENU_POPUP_H_
#define PXSBASE_MENU_POPUP_H_

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
#include "PxsBase/Header Files/Menu.h"

// 5. This Project

// 6. Forwards
class MenuItem;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class MenuPopup : public Menu
{
    public:
        // Default constructor
        MenuPopup();

        // Destructor
        ~MenuPopup();

        // Methods
        void    AddMenuItem( MenuItem* pMenuItem );
        void    AddPopUpMenu( HMENU hMenu, MenuItem* pMenuItem );
        void    AddSeparator();
        void    Destroy();
        void    RemoveAllMenuItems();
        void    Repaint();
        void    Show( HMENU hMenuParent, UINT position );

    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        MenuPopup( const MenuPopup& oMenuPopup );

        // Assignment operator - not allowed
        MenuPopup& operator= ( const MenuPopup& oMenuPopup );

        // Methods

        // Data members
        UINT    m_uPosition;
        HMENU   m_hMenuParent;
};

#endif  // PXSBASE_MENU_POPUP_H_
