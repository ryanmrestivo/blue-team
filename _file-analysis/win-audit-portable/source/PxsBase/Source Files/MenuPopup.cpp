///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Menu Pop-up Class Implementation
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
#include "PxsBase/Header Files/MenuPopup.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/MenuItem.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
MenuPopup::MenuPopup()
          :m_uPosition( UINT_MAX ),  // -1
           m_hMenuParent( nullptr )
{
}

// Copy constructor - not allowed so no implementation

// Destructor
MenuPopup::~MenuPopup()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator - Not allowed so no implementation

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Add a menu item to this pop up menu.
//
//  Parameters:
//      pMenuItem - the menu item to be added. The menu item must remain
//                  valid for the lifetime of this popup. On output the item's
//                  parent is set to this pop-up
//
//  Returns:
//      void
//===============================================================================================//
void MenuPopup::AddMenuItem( MenuItem* pMenuItem )
{
    int   itemCount = 0;
    UINT  itemIndex = 0;
    HWND  hWndMainFrame;
    MENUITEMINFO mii;

    if ( pMenuItem == nullptr )
    {
        throw ParameterException( L"pMenuItem", __FUNCTION__ );
    }

    if ( m_hMenu == nullptr )
    {
        throw FunctionException( L"m_hMenu", __FUNCTION__ );
    }

    // Add it then set tell it who its parent is
    if ( AppendMenu( m_hMenu,
                     MF_OWNERDRAW | MF_ENABLED,
                     pMenuItem->GetCommandID(), reinterpret_cast<LPCWSTR>( pMenuItem ) ) == 0 )
    {
        throw SystemException( GetLastError(), L"AppendMenu", __FUNCTION__ );
    }
    pMenuItem->SetParent( m_hMenu );

    // Account for Right To Left Reading, pop-up menus can be nested or even
    // free floating so its not easy to identify the owning window. So get
    // the handle to the application's main frame and test that for
    // directionality. This feature is only required for child pop-ups.
    if ( g_pApplication == nullptr )
    {
        return;
    }
    hWndMainFrame = g_pApplication->GetHwndMainFrame();
    if ( hWndMainFrame == nullptr )
    {
        return;
    }
    if ( WS_EX_RTLREADING & GetWindowLongPtr( hWndMainFrame, GWL_EXSTYLE ) )
    {
        itemCount = GetMenuItemCount( m_hMenu );
        if ( itemCount > 0 )
        {
            // Get the added item's info and set the RTL style
            itemIndex = PXSCastInt32ToUInt32( itemCount - 1 );
            memset( &mii, 0, sizeof ( mii ) );
            mii.cbSize = sizeof ( mii );
            mii.fMask = MIIM_TYPE;
            if ( GetMenuItemInfo( m_hMenu, itemIndex, TRUE, &mii ) )
            {
                mii.fType |= MFT_RIGHTORDER;
                SetMenuItemInfo( m_hMenu, itemIndex, TRUE, &mii );
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Add a pop-up menu to this pop up menu.
//
//  Parameters:
//      hMenu     - handle to the pop-up menu that will be shown when the
//                  this item is selected
///      pMenuItem - the menu item object that owns the popup. It must remain
//                   valid for the lifetime of this popup. On output the item's
//                   parent is set to this popup
//
//  Returns:
//      void
//===============================================================================================//
void MenuPopup::AddPopUpMenu( HMENU hMenu, MenuItem* pMenuItem )
{
    if ( pMenuItem == nullptr )
    {
       throw ParameterException( L"pMenuItem", __FUNCTION__ );
    }

    if ( hMenu == nullptr )
    {
       throw ParameterException( L"hMenu", __FUNCTION__ );
    }

    if ( m_hMenu == nullptr )
    {
        throw FunctionException( L"m_hMenu", __FUNCTION__ );
    }

    if ( AppendMenu( m_hMenu,
                     MF_OWNERDRAW | MF_ENABLED | MF_POPUP,
                     reinterpret_cast<UINT_PTR>( hMenu ),
                     reinterpret_cast<LPCWSTR>( pMenuItem ) ) == 0 )
    {
        throw SystemException( GetLastError(), L"AppendMenu", __FUNCTION__ );
    }
    pMenuItem->SetParent( m_hMenu );
}

//===============================================================================================//
//  Description:
//      Add a menu separator
//
//  Parameters:
//      None
//
//  Remarks:
//      Must have created the pop-up menu first
//
//  Returns:
//      void
//===============================================================================================//
void MenuPopup::AddSeparator()
{
    int  itemCount = 0;
    UINT itemIndex = 0;

    // Check class scope variables
    if ( m_hMenu == nullptr )
    {
        throw FunctionException( L"m_hMenu", __FUNCTION__ );
    }

    itemCount = GetMenuItemCount( m_hMenu );
    if ( itemCount == -1 )
    {
        throw SystemException( GetLastError(), L"GetMenuItemCount", __FUNCTION__ );
    }
    itemIndex = PXSCastInt32ToUInt32( itemCount );

    if ( InsertMenu( m_hMenu,
                     itemIndex, MF_OWNERDRAW | MF_SEPARATOR | MF_BYPOSITION, 0, nullptr ) == 0 )
    {
        throw SystemException( GetLastError(), L"InsertMenu", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Destroy the pop-up menu, no effect if not yet created
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void MenuPopup::Destroy()
{
    if ( m_hMenu )
    {
        RemoveAllMenuItems();
        DestroyMenu( m_hMenu );
        m_hMenu = nullptr;
    }
}

//===============================================================================================//
//  Description:
//      Remove all the menu items associated with this pop-up menu
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void MenuPopup::RemoveAllMenuItems()
{
    int i, itemCount;

    if ( m_hMenu == nullptr )
    {
        return;     // Nothing to do
    }

    itemCount = GetMenuItemCount( m_hMenu );
    if ( itemCount == -1 )
    {
        throw SystemException( GetLastError(), L"GetMenuItemCount", __FUNCTION__ );
    }

    for ( i = 0; i < itemCount; i++ )
    {
        // Keep removing the first one
        if ( RemoveMenu( m_hMenu, 0, MF_BYPOSITION ) == 0 )
        {
            throw SystemException( GetLastError(), L"RemoveMenu", __FUNCTION__);
        }
    }
}

//===============================================================================================//
//  Description:
//      Repaint the pop-up menu
//
//  Parameters:
//      None
//
//  Remarks:
//      Ignore errors, do not want recursive exceptions
//
//  Returns:
//      void
//===============================================================================================//
void MenuPopup::Repaint()
{
    RECT  rcItem = { 0, 0, 0, 0 };
    POINT centre = { 0, 0 };
    HWND  hWnd   = nullptr;
    HMENU hMenu  = GetMenuHandle();

    if ( hMenu == nullptr )
    {
        return;     // Nothing to do
    }

    // Get the rectangle of the zero-th item
    if ( GetMenuItemRect( nullptr, hMenu, 0, &rcItem ) == 0 )
    {
        return;     // Nothing to draw
    }

    if ( rcItem.right != rcItem.left )
    {
        // Get the window handle of the pop-up menu
        centre.x = ( rcItem.left + rcItem.right  ) / 2;
        centre.y = ( rcItem.top  + rcItem.bottom ) / 2;
        hWnd     = WindowFromPoint( centre );
        if ( hWnd )
        {
            // Want the entire client area redrawn
            InvalidateRect( hWnd, nullptr, FALSE );
            UpdateWindow( hWnd );
        }
    }
}

//===============================================================================================//
//  Description:
//      Create the pop-up menu
//
//  Parameters:
//      hMenuParent - handle to the parent of this pop-up menu usually a menu
//                    bar or another pop-up menu, but can be null if this menu
//                    is floating.
//      position    - position of the pop-up menu, zero-based
//
//  Returns:
//      void
//===============================================================================================//
void MenuPopup::Show( HMENU hMenuParent, UINT position )
{
    // Only allow creation only once
    if ( m_hMenu )
    {
        throw FunctionException( L"m_hMenu", __FUNCTION__ );
    }

    // Create it
    m_hMenu = CreatePopupMenu();
    if ( m_hMenu == nullptr )
    {
        throw SystemException( GetLastError(), L"CreatePopupMenu", __FUNCTION__ );
    }

    // Set the parent
    m_uPosition   = position;
    m_hMenuParent = hMenuParent;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
