///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Menu Item Class Implementation
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

// Will identify menu items by MF_BYCOMMAND, i.e. by their command identifiers.

///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/MenuItem.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/StringT.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
MenuItem::MenuItem()
         :m_bHyperLink( false ),
          m_bMenuBarItem( false ),
          m_uCommandID( 0 ),
          m_hMenuParent( nullptr ),
          m_hBitmap( nullptr ),
          m_Label()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
MenuItem::~MenuItem()
{
    // Clean up
    if ( m_hBitmap )
    {
        DeleteObject( m_hBitmap );
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
//      Get the handle of the bitmap image for this menu item.
//
//  Parameters:
//      void
//
//  Returns:
//      HBITMAP of image
//===============================================================================================//
HBITMAP MenuItem::GetBitmap() const
{
    return m_hBitmap;
}

//===============================================================================================//
//  Description:
//      Get this menu item's command identifier
//
//  Parameters:
//      None
//
//  Returns:
//      UINT
//===============================================================================================//
UINT MenuItem::GetCommandID() const
{
    return m_uCommandID;
}

//===============================================================================================//
//  Description:
//      Get the label of this menu item
//
//  Parameters:
//      None
//
//  Returns:
//      A reference to the label
//===============================================================================================//
const String& MenuItem::GetLabel() const
{
    return m_Label;
}

//===============================================================================================//
//  Description:
//      Get the parent menu of this menu item
//
//  Parameters:
//      None
//
//  Returns:
//      Handle to parent menu
//===============================================================================================//
HMENU MenuItem::GetParent() const
{
    return m_hMenuParent;
}

//===============================================================================================//
//  Description:
//      Get if the menu item is checked.
//
//  Parameters:
//      None
//
//  Remarks:
//      Item must exist and the parent set, else will return false
//
//  Returns:
//      true if checked, else false
//===============================================================================================//
bool MenuItem::IsChecked() const
{
    bool  checked   = false;
    UINT  menuState = 0;

    // Check for a parent menu
    if ( m_hMenuParent == nullptr )
    {
        return false;
    }

    menuState = GetMenuState( m_hMenuParent, m_uCommandID, MF_BYCOMMAND );
    if ( menuState == UINT_MAX )   // = -1
    {
        throw SystemException( ERROR_MENU_ITEM_NOT_FOUND, L"GetMenuState", __FUNCTION__ );
    }

    if ( menuState & MF_CHECKED )
    {
        checked = true;
    }

    return checked;
}

//===============================================================================================//
//  Description:
//        Get if this the menu item is enabled.
//
//  Parameters:
//      None
//
//  Remarks:
//      Returns false if the parent is not set or the menu item does not exist
//
//  Returns:
//      true if enabled, else false
//===============================================================================================//
bool MenuItem::IsEnabled() const
{
    UINT menuState;

    // Must have a parent
    if ( m_hMenuParent == nullptr )
    {
        return false;
    }

    menuState = GetMenuState( m_hMenuParent, m_uCommandID, MF_BYCOMMAND );
    if ( menuState == UINT_MAX )   // = -1
    {
        throw SystemException( ERROR_MENU_ITEM_NOT_FOUND, L"GetMenuState", __FUNCTION__ );
    }

    // Test for MF_DISABLED as MF_ENABLED = 0
    if ( menuState & MF_DISABLED )
    {
        return false;
    }

    return true;
}

//===============================================================================================//
//  Description:
//      Get the menu item is a hyper-link.
//
//  Parameters:
//      None
//
//  Returns:
//      true is hyper-link, else false
//===============================================================================================//
bool MenuItem::IsHyperLink() const
{
    return m_bHyperLink;
}

//===============================================================================================//
//  Description:
//        Get if this the menu item is on the menu bar.
//
//  Parameters:
//      None
//
//  Remarks:
//      Menu items are either on the windows menu bar or the pop-up menu
//
//  Returns:
//      true if item is on the menu bar
//===============================================================================================//
bool MenuItem::IsMenuBarItem() const
{
    return m_bMenuBarItem;
}

//===============================================================================================//
//  Description:
//      Set the bitmap image for this menu item
//
//  Parameters:
//      resourceID  - name of the bitmap resource ID
//
//  Returns:
//      void
//===============================================================================================//
void MenuItem::SetBitmap( WORD resourceID )
{
    HBITMAP   hBitmap = nullptr;
    String    ErrorDetails;
    Formatter Format;

    // Load the bitmap
    if ( resourceID )
    {
        hBitmap = LoadBitmap( GetModuleHandle( nullptr ), MAKEINTRESOURCE( resourceID ) );
        if ( hBitmap == nullptr )
        {
            ErrorDetails = Format.StringUInt32( L"LoadBitmap, resourceID=%%1", resourceID );
            throw SystemException( GetLastError(), ErrorDetails.c_str(), __FUNCTION__ );
        }
    }

    // Replace
    if ( m_hBitmap )
    {
        DeleteObject( m_hBitmap );
    }
    m_hBitmap = hBitmap;
}

//===============================================================================================//
//  Description:
//      Set if the menu item is checked
//
//  Parameters:
//      checked - flag to indicate whether there is a check mark
//
//  Remarks:
//      Parent must be set first
//
//  Returns:
//      void
//===============================================================================================//
void MenuItem::SetChecked( bool checked )
{
    UINT flags = 0;

    // Check parent menu
    if ( m_hMenuParent == nullptr )
    {
        throw FunctionException( L"m_hMenuParent", __FUNCTION__ );
    }

    // Ignore errors
    if ( checked )
    {
        flags = MF_BYCOMMAND | MF_CHECKED;
    }
    else
    {
        flags = MF_BYCOMMAND | MF_UNCHECKED;
    }
    CheckMenuItem( m_hMenuParent, m_uCommandID, flags );
}

//===============================================================================================//
//  Description:
//      Set the command ID associated with this menu item, i.e. the message
//      posted when the menu item is selected.
//
//  Parameters:
//      commandID - the command ID
//
//  Remarks:
//      Command ID's are of type WORD to keep consistent with any keyboard
//      accelerators and application messages which are in the range
//      WM_APP(0x8000) through 0xBFFF i.e. unsigned short.
//
//  Returns:
//      void
//===============================================================================================//
void MenuItem::SetCommandID( WORD commandID )
{
    m_uCommandID = commandID;
}

//===============================================================================================//
//  Description:
//        Set if this the menu item is enabled.
//
//  Parameters:
//      enabled - flag to indicate the state of the menu item
//
//  Remarks:
//      Menu item must already have been attached to the parent menu
//
//  Returns:
//      void
//===============================================================================================//
void MenuItem::SetEnabled( bool enabled )
{
    UINT flags = 0;

    // Check parent menu
    if ( m_hMenuParent == nullptr )
    {
        throw FunctionException( L"m_hMenuParent", __FUNCTION__ );
    }

    // Ignore errors
    if ( enabled )
    {
        flags = MF_BYCOMMAND | MF_ENABLED;
    }
    else
    {
        flags = MF_BYCOMMAND | MF_GRAYED;
    }
    EnableMenuItem( m_hMenuParent, m_uCommandID, flags );
}

//===============================================================================================//
//  Description:
//      Set the bitmap for this menu item as a filled rectangle of the
//      specified fill colour and border.
//
//  Parameters:
//      fill        - the colour to fill the interior of the border
//      border      - the border colour, 1 pixel wide
//
//  Returns:
//      void
//===============================================================================================//
void MenuItem::SetFilledBitmap( COLORREF fill, COLORREF border )
{
    HDC     hDC     = nullptr;
    HBITMAP hBitmap = nullptr;

    // Use the desktop's device context
    hDC = GetDC( nullptr );
    if ( hDC == nullptr )
    {
        throw SystemException( GetLastError(), L"hDC", __FUNCTION__ );
    }

    // Catch exceptions to release the DC
    try
    {
        // There does not appear to a metric for menu item images so
        // will use 16 x 16 pixels
        hBitmap = PXSCreateFilledBitmap( hDC, 16, 16, fill, border, PXS_SHAPE_RECTANGLE );
    }
    catch ( const Exception& )
    {
        DeleteObject( hBitmap );
        ReleaseDC( nullptr, hDC );
        throw;
    }
    ReleaseDC( nullptr, hDC );

    // Replace
    if ( m_hBitmap )
    {
        DeleteObject( m_hBitmap );
    }
    m_hBitmap = hBitmap;
}

//===============================================================================================//
//  Description:
//      Set if the menu item is a hyper-link
//
//  Parameters:
//      hyperLink - true if is hyper-link, else false
//
//  Returns:
//     void
//===============================================================================================//
void MenuItem::SetHyperLink( bool hyperLink )
{
    m_bHyperLink = hyperLink;
}

//===============================================================================================//
//  Description:
//      Set the label for this menu item
//
//  Parameters:
//      Label - the label
//
//  Remarks:
//      Do not need to call ModifyMenu because it is owner drawn.
//
//  Returns:
//      void
//===============================================================================================//
void MenuItem::SetLabel( const String& Label )
{
    m_Label = Label;
}

//===============================================================================================//
//  Description:
//      Set if this the menu item is on the menu bar.
//
//  Parameters:
//      menuBarItem - flag to indicate the state of the menu item
//
//  Remarks:
//      Menu items are either on the windows menu bar or the pop-up menu
//
//  Returns:
//      void
//===============================================================================================//
void MenuItem::SetMenuBarItem( bool menuBarItem )
{
    m_bMenuBarItem = menuBarItem;
}

//===============================================================================================//
//  Description:
//      Set the parent menu of this menu item
//
//  Parameters:
//      hMenuParent - handle to parent menu
//
//  Returns:
//      void
//===============================================================================================//
void MenuItem::SetParent( const HMENU menuParent )
{
    m_hMenuParent = menuParent;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
