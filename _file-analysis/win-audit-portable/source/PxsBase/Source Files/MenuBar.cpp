///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Menu Bar Class Implementation
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
#include "PxsBase/Header Files/MenuBar.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/MenuItem.h"
#include "PxsBase/Header Files/MenuPopup.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/Window.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
MenuBar::MenuBar()
        :m_hWndOwner( nullptr )
{
}

// Copy constructor - not allowed so no implementation

// Destructor
MenuBar::~MenuBar()
{
    try
    {
        Destroy();
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
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
//      Add a pop up menu to this menu bar
//
//  Parameters:
//      MenuBarItem - the menu bar item to which the pop-up belongs
//      Popup       - the pop up to add
//
//  Returns:
//      void
//===============================================================================================//
void MenuBar::AddPopup( const MenuItem& MenuBarItem, const MenuPopup& Popup )
{
    int      itemCount  = 0;
    UINT     itemIndex  = 0;
    LPCWSTR  lpNewItem  = nullptr;
    UINT_PTR uIDNewItem = 0;
    MENUITEMINFO mii;

    // Need a menu
    if ( m_hMenu == nullptr )
    {
        throw FunctionException( L"m_hMenu", __FUNCTION__ );
    }

    if ( MenuBarItem.IsMenuBarItem() == false )
    {
        throw FunctionException( L"IsMenuBarItem", __FUNCTION__ );
    }

    // The menu bar item must belong to this menu bar
    if ( MenuBarItem.GetParent() != m_hMenu )
    {
        throw SystemException( ERROR_INVALID_MENU_HANDLE,
                               L"GetParent() != m_hMenu", __FUNCTION__ );
    }

    uIDNewItem = reinterpret_cast<UINT_PTR>( Popup.GetMenuHandle() );
    lpNewItem  = reinterpret_cast<LPCWSTR>( &MenuBarItem );
    if ( AppendMenu( m_hMenu, MF_OWNERDRAW | MF_POPUP, uIDNewItem, lpNewItem ) == 0 )
    {
        throw SystemException ( GetLastError(), L"AppendMenu", __FUNCTION__ );
    }

    // Account for Right To Left Reading
    if ( WS_EX_RTLREADING & GetWindowLongPtr( m_hWndOwner, GWL_EXSTYLE ) )
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
                mii.fType |= ( MFT_RIGHTJUSTIFY | MFT_RIGHTORDER );
                SetMenuItemInfo( m_hMenu, itemIndex, TRUE, &mii );
            }
        }
    }
    Draw();     // Need to do this any time the menu bar changes
}

//===============================================================================================//
//  Description:
//      Create the menu bar
//
//  Parameters:
//      ownerWindow - the window that owns this menu bar
//
//  Remarks:
//      The owner window must have been created first.
//
//  Returns:
//      void
//===============================================================================================//
void MenuBar::Create( const Window& ownerWindow )
{
    DWORD lastError = 0;

    // Only allow creation once
    if ( m_hMenu )
    {
        throw FunctionException( L"m_hMenu", __FUNCTION__ );
    }

    if ( ownerWindow.GetHwnd() == nullptr )
    {
        throw ParameterException( L"ownerWindow", __FUNCTION__ );
    }

    // Create the menu
    m_hMenu = CreateMenu();
    if ( m_hMenu == nullptr )
    {
        throw SystemException( GetLastError(), L"CreateMenu", __FUNCTION__ );
    }

    // Attach to owner
    if ( SetMenu( ownerWindow.GetHwnd(), m_hMenu ) == 0 )
    {
        lastError = GetLastError();
        DestroyMenu( m_hMenu );
        m_hMenu = nullptr;
        throw SystemException( lastError, L"SetMenu", __FUNCTION__ );
    }
    m_hWndOwner = ownerWindow.GetHwnd();
}

//===============================================================================================//
//  Description:
//      Destroy the menu bar, no effect if not yet created
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void MenuBar::Destroy()
{
    if ( m_hMenu )
    {
        RemoveAllPopupMenus();
        DestroyMenu( m_hMenu );
        m_hMenu = nullptr;
    }
}

//===============================================================================================//
//  Description:
//         Draw the menu bar.
//
//  Parameters:
//      None
//
//  Remarks:
//      Call this method after the menu bar has been created or altered.
//
//  Returns:
//      void
//===============================================================================================//
void MenuBar::Draw() const
{
    // Check owner window
    if ( m_hWndOwner == nullptr )
    {
        throw FunctionException( L"m_hWndOwner", __FUNCTION__ );
    }

    if ( DrawMenuBar( m_hWndOwner ) == 0 )
    {
        throw SystemException( GetLastError(), L"DrawMenuBar", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Remove all the pop-up menus associated with this menu bar
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void MenuBar::RemoveAllPopupMenus()
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
    Draw();     // Need to do this any time the menu bar changes
}

//===============================================================================================//
//  Description:
//      Show/hide the menu bar.
//
//  Parameters:
//      visible - flag to indicate if the menu bar is visible
//
//  Remarks:
//      Only effective if the menu bar has been created and has an
//      owner window
//
//  Returns:
//      void
//===============================================================================================//
void MenuBar::SetVisible( bool visible )
{
    // Check owner window and menu handle
    if ( m_hWndOwner == nullptr || m_hMenu == nullptr )
    {
        return;     // Nothing to do
    }

    if ( GetMenu( m_hWndOwner ) )
    {
        return;     // Nothing to do
    }

    if ( visible )
    {
        if ( SetMenu( m_hWndOwner, m_hMenu ) == 0 )
        {
            throw SystemException( GetLastError(), L"SetMenu", __FUNCTION__ );
        }
    }
    else
    {
        // Set the window's menu to NULL
        if ( SetMenu( m_hWndOwner, nullptr ) == 0 )
        {
            throw SystemException( GetLastError(), L"SetMenu", __FUNCTION__ );
        }
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
