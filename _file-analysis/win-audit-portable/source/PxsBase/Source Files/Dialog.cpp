///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Dialog Class Implementation
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
#include "PxsBase/Header Files/Dialog.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
Dialog::Dialog()
       :m_Title()
{
    COLORREF crGradient1 = CLR_INVALID, crGradient2 = CLR_INVALID;

    // Set properties
    m_crBackground = GetSysColor( COLOR_BTNFACE );
    try
    {
        if ( g_pApplication )
        {
            g_pApplication->GetGradientColours( &crGradient1, &crGradient2 );
            SetBackgroundGradient( crGradient1, crGradient2, false );
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Copy constructor - not allowed so no implementation

// Destructor
Dialog::~Dialog()
{
    // The user data is a pointer to 'this'
    if ( m_hWindow )
    {
        SetWindowLongPtr( m_hWindow, GWLP_USERDATA, 0 );
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
//      Create/Show this modal dialogue box
//
//  Parameters:
//      hWndParent - window that owners the dialogue box
//
//  Remarks:
//      If the owner is NULL will use the desktop
//
//  Returns:
//      signed integer, the result specified in the call to the EndDialog
//      when the dialogue box was closed
//===============================================================================================//
INT_PTR Dialog::Create( HWND hWndParent )
{
    RECT    recParent = { 0, 0, 0, 0 };
    LONG    baseUnits = 0;
    WORD    baseUnitX = 0, baseUnitY = 0;
    DWORD   style;
    wchar_t wTemplate[ 256 ] = { 0 };
    POINT   dialogLocation   = { 0, 0 };
    INT_PTR resultPtr        = 0;
    Formatter Format;

    // Check inputs, use desktop if no owner
    if ( hWndParent == nullptr )
    {
        hWndParent = GetDesktopWindow();
    }

    // Dialog box style. Note, must not use the WS_CHILD style
    style = WS_POPUPWINDOW  |
            WS_VISIBLE      |
            WS_CLIPSIBLINGS |
            WS_DLGFRAME     |
            WS_OVERLAPPED   | DS_3DLOOK | WS_CAPTION | WS_SYSMENU | DS_MODALFRAME | DS_SETFONT;

    // Position the dialogue, if none position centre it over the parent
    GetLocation( &dialogLocation );
    if ( ( dialogLocation.x == 0 ) && ( dialogLocation.y == 0 ) )
    {
        GetClientRect( hWndParent, &recParent );
        dialogLocation.x = ( recParent.right  - m_CreateStruct.cx ) / 2;
        dialogLocation.y = ( recParent.bottom - m_CreateStruct.cy ) / 2;
    }
    baseUnits = GetDialogBaseUnits();
    baseUnitX = LOWORD( baseUnits );
    baseUnitY = HIWORD( baseUnits );

    // DLGTEMPLATE in wchar_t/WORD
    wTemplate[  0 ] = LOWORD( style );
    wTemplate[  1 ] = HIWORD( style );
    wTemplate[  2 ] = 0;                // ExtendedStyle
    wTemplate[  3 ] = 0;                // ExtendedStyle
    wTemplate[  4 ] = 0;                // NumberOfItems
    wTemplate[  5 ] = LOWORD( (dialogLocation.x  * 4) / baseUnitX );  // x
    wTemplate[  6 ] = LOWORD( (dialogLocation.y  * 8) / baseUnitY );  // y
    wTemplate[  7 ] = LOWORD( (m_CreateStruct.cx * 4) / baseUnitX );  // cx
    wTemplate[  8 ] = LOWORD( (m_CreateStruct.cy * 8) / baseUnitY );  // cy
    wTemplate[  9 ] = 0;                                              // Menu
    wTemplate[ 10 ] = 0;                                              // Class

    // Append the title
    if ( m_Title.c_str() )
    {
        StringCchCopy( wTemplate + 11, ARRAYSIZE( wTemplate ) - 12, m_Title.c_str() );
    }
    wTemplate[ ARRAYSIZE( wTemplate ) - 1 ] = PXS_CHAR_NULL;
    resultPtr = DialogBoxIndirectParam( GetModuleHandle( nullptr ),
                                        reinterpret_cast<LPDLGTEMPLATE>( &wTemplate ),
                                        hWndParent,
                                        DialogProc, reinterpret_cast<LPARAM>( this ) );
    if ( resultPtr == -1 )
    {
        throw SystemException( GetLastError(), L"DialogBoxIndirectParam", __FUNCTION__ );
    }

    return resultPtr;
}

//===============================================================================================//
//  Description:
//      Set the dialogue's  title
//
//  Parameters:
//      Title - pointer to title
//
//  Returns:
//      void
//===============================================================================================//
void Dialog::SetTitle( const String& Title )
{
    m_Title = Title;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Static window callback procedure.
//
//  Parameters:
//      hwndDlg - Handle to the window that generated the message.
//      uMsg    - Specifies the message.
//      wParam  - Specifies additional message information.
//      lParam  - Specifies additional message information.
//
//  Remarks:
//      The WM_INITDIALOG message has a pointer to the dialog window
//      in lParam. This was passed in using DialogBoxIndirectParam.
//
//      Returns:
//      INT_PTR, which is usually 0 if the message was processed.
//===============================================================================================//
INT_PTR Dialog::DialogProc( HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
    bool showException = true;
    Window* pWindow    = nullptr;
    Dialog* pDialog    = reinterpret_cast<Dialog*>( GetWindowLongPtr( hwndDlg, GWLP_USERDATA ) );
    INT_PTR resultPtr  = 0;
    Formatter Format;

    // Only allow null dialog if its being created
    if ( ( pDialog == nullptr ) && ( uMsg != WM_INITDIALOG ) )
    {
        return 0;
    }

    try
    {
        switch ( uMsg )
        {
            default:

                if ( uMsg >= WM_APP )
                {
                    pDialog->AppMessageEvent( uMsg, wParam, lParam );
                }
                break;

            case WM_INITDIALOG:

                if ( lParam )   // = pointer to this
                {
                    Dialog* pInit = reinterpret_cast<Dialog*>( lParam );
                    resultPtr = pInit->DoInitDialog( hwndDlg, lParam );
                }
                break;

            case WM_COMMAND:

                resultPtr = pDialog->DoWmCommand( wParam, lParam );
                break;

            case WM_DRAWITEM:

                // Do not show exceptions, used in painting
                showException = false;
                pDialog->DoDrawItem( reinterpret_cast<const DRAWITEMSTRUCT*>( lParam ) );
                resultPtr = TRUE;     // Documentation says should return TRUE
                break;

            case WM_CTLCOLORSTATIC:
            case WM_ERASEBKGND:
            case WM_HELP:
            case WM_KILLFOCUS:
            case WM_SETFOCUS:
            case WM_LBUTTONUP:
            case WM_MEASUREITEM:
            case WM_MOUSEMOVE:
            case WM_NOTIFY:
            case WM_PAINT:
            case WM_SIZE:
            case WM_TIMER:

                // These can be handled by WindowProc
                pWindow   = static_cast< Window* >( pDialog );
                resultPtr = (INT_PTR)pWindow->WindowProc( hwndDlg, uMsg, wParam, lParam );
                break;
        }
    }
    catch ( const Exception& e )
    {
        if ( showException )
        {
            HWND hWndOwner = hwndDlg;
            if ( uMsg == WM_INITDIALOG )
            {
                EndDialog( hwndDlg, IDCANCEL );
                if ( g_pApplication )
                {
                    hWndOwner = g_pApplication->GetHwndMainFrame();
                }
                MessageBox( hWndOwner, e.GetMessage().c_str(), nullptr, MB_OK | MB_ICONERROR );
            }
            else
            {
                PXSShowExceptionDialog( e, hWndOwner );
            }
        }
    }

    return resultPtr;
}

//===============================================================================================//
//  Description:
//      Handle WM_DRAWITEM event for button items
//
//  Parameters:
//      pDraw - pointer to item data structure
//
//  Returns:
//      void
//===============================================================================================//
void Dialog::DrawButtonItemEvent( const DRAWITEMSTRUCT* pDraw )
{
    RECT    windowRect = { 0, 0, 0, 0 }, innerRect= { 0, 0, 0, 0 };
    POINT   cursorPos  = { 0, 0 };
    String  Caption;
    wchar_t szCaption[ 256 ] = { 0 };
    StaticControl Static;

    if ( ( pDraw      == nullptr ) ||
         ( pDraw->hDC == nullptr ) ||
         ( pDraw->hwndItem == nullptr ) )
    {
        return;
    }

    // Set the background and border
    GetCursorPos( &cursorPos );
    GetWindowRect( pDraw->hwndItem, &windowRect );
    if ( ( pDraw->hwndItem == GetCapture()  )  &&
         ( PtInRect( &windowRect, cursorPos ) ) )
    {
        Static.SetShape( PXS_SHAPE_3D_SUNK );
    }
    else
    {
        Static.SetShape( PXS_SHAPE_RAISED );
    }
    Static.SetBackground( GetSysColor( COLOR_BTNFACE ) );
    Static.SetBackgroundGradient( m_crGradient1, m_crGradient2, true );

    // Set the caption depending on the control's state
    GetWindowText( pDraw->hwndItem, szCaption, ARRAYSIZE( szCaption ) );
    szCaption[ ARRAYSIZE( szCaption ) - 1 ] = PXS_CHAR_NULL;
    Caption = szCaption;
    Static.SetSingleLineText( Caption );
    if ( ODS_DISABLED & pDraw->itemState )
    {
        Static.SetShapeColour( PXS_COLOUR_LITEGREY );
        Static.SetForeground( PXS_COLOUR_LITEGREY );
    }
    else
    {
        Static.SetShapeColour( PXS_COLOUR_BLACK );
        Static.SetForeground( PXS_COLOUR_BLACK );
    }
    Static.SetBounds( pDraw->rcItem );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
    Static.Draw( pDraw->hDC );
    Static.Reset();

    // Highlight the button with the focus otherwise the default button
    memset( &innerRect, 0, sizeof ( innerRect ) );
    if ( ODS_FOCUS & pDraw->itemState )
    {
        innerRect.left   = pDraw->rcItem.left   + 2;
        innerRect.top    = pDraw->rcItem.top    + 2;
        innerRect.bottom = pDraw->rcItem.bottom - 3;
        innerRect.right  = pDraw->rcItem.right  - 3;
        Static.SetBounds( innerRect );
        Static.SetShape( PXS_SHAPE_RECTANGLE_DOTTED );
        Static.Draw( pDraw->hDC );

        // Draw an enter arrow to the top right
        MoveToEx(pDraw->hDC, pDraw->rcItem.right -  8, pDraw->rcItem.top + 4,
                 nullptr );
        LineTo(  pDraw->hDC, pDraw->rcItem.right -  8, pDraw->rcItem.top + 7 );
        LineTo(  pDraw->hDC, pDraw->rcItem.right - 16, pDraw->rcItem.top + 7 );
        LineTo(  pDraw->hDC, pDraw->rcItem.right - 14, pDraw->rcItem.top + 5 );
        LineTo(  pDraw->hDC, pDraw->rcItem.right - 14, pDraw->rcItem.top + 9 );
        LineTo(  pDraw->hDC, pDraw->rcItem.right - 16, pDraw->rcItem.top + 7 );
    }
}

//===============================================================================================//
//  Description:
//      Handle the WM_INITDIALOG message
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Dialog::InitDialogEvent()
{
}

//===============================================================================================//
//  Description:
//      Determine if the parameters are a click from the specified button
//
//  Parameters:
//      ButtonControl - the button
//      wParam -
//      lParam -
//
//  Returns:
//      true if the focus was changed otherwise false
//===============================================================================================//
bool Dialog::IsClickFromButton( const Button& ButtonControl, WPARAM wParam, LPARAM lParam )
{
    if ( ( lParam != 0 ) &&
         ( HIWORD( wParam) == BN_CLICKED ) &&
         ( ButtonControl.GetHwnd() == reinterpret_cast<HWND>( lParam ) ) )
    {
        return true;
    }

    return false;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Dispatch the WM_DRAWITEM message.
//
//  Parameters:
//      lParam - LPARAM from DialogProc
//
//  Returns:
//      void
//===============================================================================================//
void Dialog::DoDrawItem( const DRAWITEMSTRUCT* pDraw )
{
    if ( pDraw == nullptr )
    {
        return;
    }

    if ( pDraw->CtlType == ODT_MENU )
    {
        DrawMenuItemEvent( pDraw );
    }
    else if ( pDraw->CtlType == ODT_BUTTON )
    {
        DrawButtonItemEvent( pDraw );
    }
}

//===============================================================================================//
//  Description:
//      Process and Dispatch the WM_INITDIALOG message.
//
//  Parameters:
//      hwndDlg - handle to the dialog window
//      lParam  - LPARAM from DialogProc
//
//  Returns:
//      1 if processed otherwise 0
//===============================================================================================//
INT_PTR Dialog::DoInitDialog( HWND hwndDlg, LPARAM lParam )
{
    HWND    hWndParent;
    LONG    exStyle = 0;
    POINT   location= { 0, 0 };
    Dialog* pDialog = reinterpret_cast<Dialog*>( lParam );

    if ( ( hwndDlg == nullptr ) || ( pDialog == nullptr ) )
    {
        return 0;
    }

    // Save pointer to this dialogue which was passed in as
    // the initialization value of DialogBoxIndirectParam
    SetWindowLongPtr( hwndDlg, GWLP_USERDATA, lParam );  // LPARAM = LONG_PTR

    // Propagate the Right to Left style
    hWndParent = GetParent( hwndDlg );
    exStyle    = GetWindowLong( hWndParent, GWL_EXSTYLE );
    if ( WS_EX_RTLREADING & exStyle )
    {
        exStyle |= ( WS_EX_RIGHT | WS_EX_RTLREADING );
        SetWindowLong( hwndDlg, GWL_EXSTYLE, exStyle );
    }
    SetWindowPos( hwndDlg,
                  nullptr,
                  0,
                  0,
                  0,
                  0,
                  SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOREDRAW );

    // Position the window, if one was specified
    pDialog->GetLocation( &location );
    if ( location.x || location.y )
    {
        SetWindowPos( hwndDlg, GetDesktopWindow(), location.x, location.y, 0, 0, SWP_NOSIZE );
    }
    pDialog->SetHwnd( hwndDlg );
    pDialog->InitDialogEvent();

    return 1;  // Return non-zero to set focus on the first control
}

//===============================================================================================//
//  Description:
//      Process and Dispatch the WM_COMMAND message.
//
//  Parameters:
//      hwndDlg - handle to the dialog window
//      lParam  - LPARAM from DialogProc
//
//  Returns:
//      1 if processed otherwise 0
//===============================================================================================//
INT_PTR Dialog::DoWmCommand( WPARAM wParam, LPARAM lParam )
{
    HWND hWndFocus = GetFocus();
    INT_PTR resultPtr = 0;

    // Handle this first
    if ( LOWORD( wParam ) == IDCANCEL )
    {
        return CommandEvent( wParam, lParam );
    }

    // Handle button clicks. Note, BN_CLICKED = 0 as well as the HIWORD
    // of wParam for menu commands
    if ( ( HIWORD( wParam ) == BN_CLICKED ) && ( LOWORD( wParam ) == IDOK ) )
    {
        if ( IsButton( hWndFocus ) )
        {
            resultPtr  = CommandEvent( MAKEWPARAM( 0, BN_CLICKED ),
                                       reinterpret_cast<LPARAM>( hWndFocus ) );
        }
    }
    else
    {
        resultPtr = CommandEvent( wParam, lParam );
    }

    return resultPtr;
}

//===============================================================================================//
//  Description:
//      Determine if specified window is a button
//
//  Parameters:
//      hWnd - window handle
//
//  Returns:
//      true if the default button should be fired otherwise false
//===============================================================================================//
bool Dialog::IsButton( HWND hWnd )
{
    bool     isButton = false;
    wchar_t  szClassName[ MAX_PATH ] = { 0 };

    if ( hWnd == nullptr )
    {
        return false;
    }

    if ( GetClassName( hWnd, szClassName, ARRAYSIZE( szClassName ) ) )
    {
        szClassName[ ARRAYSIZE( szClassName ) - 1 ] = PXS_CHAR_NULL;
        if ( PXSCompareString( szClassName, L"BUTTON", false ) == 0 )
        {
            isButton = true;
        }
    }

    return isButton;
}

//===============================================================================================//
//  Description:
//      Set this dialog's HWND. This method must be called when processing
//      the WM_INITDIALOG message.
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Dialog::SetHwnd( HWND hWindow )
{
    m_hWindow = hWindow;
}
