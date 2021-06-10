///////////////////////////////////////////////////////////////////////////////////////////////////
//
/// Tool Bar Class Implementation
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
#include "PxsBase/Header Files/ToolBar.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/ImageButton.h"
#include "PxsBase/Header Files/MemoryException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
ToolBar::ToolBar()
{
    // Creation parameters
    m_CreateStruct.style  = WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN;
    m_CreateStruct.cy     = 52;

    // Properties
    m_crBackground = GetSysColor( COLOR_BTNFACE );
    try
    {
        SetLayoutProperties( PXS_LAYOUT_STYLE_ROW_MIDDLE, 3, 3 );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Copy constructor - not allowed so no implementation

// Destructor
ToolBar::~ToolBar()
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
//      Get the preferred height of the tool bar
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
int ToolBar::GetPreferredHeight()
{
    int     maxHeight   = 0, horizontalGap = 0, verticalGap = 0;
    RECT    childRect   = { 0, 0, 0, 0 };
    HWND    hWndChild   = nullptr;
    DWORD   layoutStyle = 0;

    // Checks, return a nominal value
    if ( m_hWindow == nullptr )
    {
        return 16;
    }

    // Get the maximum height of the children
    hWndChild = GetWindow( m_hWindow, GW_CHILD );
    while ( hWndChild )
    {
        memset( &childRect, 0, sizeof ( childRect ) );
        if ( GetWindowRect( hWndChild, &childRect ) )
        {
            maxHeight = PXSMaxInt( maxHeight, childRect.bottom - childRect.top );
        }
        hWndChild = GetWindow( hWndChild, GW_HWNDNEXT );
    }

    // Acount for vertical gaps
    GetLayoutProperties( &layoutStyle, &horizontalGap, &verticalGap );
    maxHeight += ( 2* verticalGap );

    return maxHeight;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Handle WM_COMMAND events.
//
//  Parameters:
//      as per WM_COMMAND
//
//  Returns:
//      0 if handled, else non-zero.
//===============================================================================================//
LRESULT ToolBar::CommandEvent( WPARAM wParam, LPARAM lParam )
{
    LRESULT result = 0;

    // Pass commands to its AppMessageListener if the command came from a child
    if ( m_hWndAppMessageListener && m_hWindow && lParam )
    {
        if ( IsChild( GetParent( m_hWindow ), reinterpret_cast<HWND>(lParam) ) )
        {
            // Send to listener
            result = SendMessage( m_hWndAppMessageListener, WM_COMMAND, wParam, lParam );
        }
    }

    return result;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

