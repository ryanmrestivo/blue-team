///////////////////////////////////////////////////////////////////////////////////////////////////
//
// ComboBox Class Implementation
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
#include "PxsBase/Header Files/ComboBox.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateWChars.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
ComboBox::ComboBox()
{
    // Class registration
    m_WndClassEx.lpszClassName = L"COMBOBOX";

    // Creation parameters, Win2000 does show list if use CW_USEDEFAULT
    m_CreateStruct.cy        = 200;
    m_CreateStruct.cx        = 110;
    m_CreateStruct.style     = WS_VISIBLE |
                               WS_CHILD   |
                               WS_TABSTOP |
                               WS_VSCROLL | WS_HSCROLL | CBS_AUTOHSCROLL | CBS_DROPDOWNLIST;
    m_CreateStruct.lpszClass = m_WndClassEx.lpszClassName;
}

// Copy constructor - not allowed so no implementation


// Destructor
ComboBox::~ComboBox()
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
//      Add a string to the combo box
//
//  Parameters:
//      pszString - pointer to null terminated string
//
//  Remarks:
//      Window must have been created first
//
//  Returns:
//      void
//===============================================================================================//
void ComboBox::Add( LPCWSTR pszString )
{
    LPARAM lParam = reinterpret_cast<LPARAM>( pszString );

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWnd", __FUNCTION__ );
    }

    // The error result depends on the type of message so will just
    // test for memory exhaustion
    if ( SendMessage( m_hWindow, CB_ADDSTRING, 0, lParam ) == CB_ERRSPACE )
    {
        throw MemoryException( __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Get the currently selected index
//
//  Parameters:
//      None
//
//  Remarks:
//      Window must have been created first
//
//  Returns:
//      zero-based index of selected item, else -1
//===============================================================================================//
size_t ComboBox::GetSelectedIndex() const
{
    size_t  index  = PXS_MINUS_ONE;
    LRESULT result = 0;

    // Check window handle
    if ( m_hWindow == nullptr )
    {
        return index;
    }

    result = SendMessage( m_hWindow, CB_GETCURSEL, 0, 0 );
    if ( result > 0 )   // CB_ERR = -1
    {
        index = PXSCastLongPtrToSizeT( result );
    }

    return index;
}

//===============================================================================================//
//  Description:
//      Get the currently selected string
//
//  Parameters:
//      pSelected - string object to receive the selected string
//
//  Remarks:
//      Window must have been created first
//
//  Returns:
//      void
//===============================================================================================//
void ComboBox::GetSelectedString( String* pSelected ) const
{
    WPARAM   bufLen   = 0, index   = 0;
    LRESULT  curSel   = 0, textLen = 0;
    wchar_t* lpBuffer = nullptr;
    AllocateWChars StringBuffer;

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    if ( pSelected == nullptr )
    {
        throw ParameterException( L"pSelected", __FUNCTION__ );
    }
    *pSelected = PXS_STRING_EMPTY;

    // Get the selected index, CB_ERR means no item selected
    curSel = SendMessage( m_hWindow, CB_GETCURSEL, 0, 0 );
    if ( ( curSel != CB_ERR ) && ( curSel >= 0 ) )
    {
        // Must use CB_GETLBTEXTLEN to get the required buffer length
        // as CB_GETLBTEXT does not make provision for it.
        index   = PXSCastLResultToWParam( curSel );
        textLen = SendMessage( m_hWindow, CB_GETLBTEXTLEN, index, 0 );
        if ( ( textLen != CB_ERR ) && ( textLen >= 0 ) )
        {
            bufLen = PXSCastLResultToWParam( textLen );
            bufLen++;       // Terminator
            lpBuffer = StringBuffer.New( bufLen );
            if ( CB_ERR == SendMessage( m_hWindow,
                                        CB_GETLBTEXT,
                                        index, reinterpret_cast<LPARAM>( lpBuffer ) ) )
            {
                throw SystemException( ERROR_INVALID_INDEX, L"SendMessage", __FUNCTION__ );
            }
            lpBuffer[ bufLen - 1 ] = PXS_CHAR_NULL;
            *pSelected = lpBuffer;
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the index of a specified string in the combo box
//
//  Parameters:
//      pszFind - pointer to the string to find
//
//  Returns:
//      -1 if the string is not found
//===============================================================================================//
size_t ComboBox::IndexOf( LPCWSTR pszFind ) const
{
    size_t  index  = PXS_MINUS_ONE;
    LRESULT result = 0;

    if ( m_hWindow == nullptr )
    {
        return index;
    }

    if ( ( pszFind == nullptr ) || ( *pszFind == PXS_CHAR_NULL ) )
    {
        return index;
    }

    result = SendMessage( m_hWindow,
                          CB_FINDSTRINGEXACT,
                          static_cast<WPARAM>( -1 ), reinterpret_cast<LPARAM>( pszFind ) );
    if ( result > 0 )   // CB_ERR = -1
    {
        index = PXSCastInt64ToSizeT( result );
    }

    return index;
}

//===============================================================================================//
//  Description:
//      Remove all the items from the list
//
//  Parameters:
//      None
//
//  Remarks:
//      According to MSDN this message always returns CB_OKAY
//
//  Returns:
//===============================================================================================//
void ComboBox::RemoveAll()
{
    if ( m_hWindow )
    {
        SendMessage( m_hWindow, CB_RESETCONTENT, 0, 0 );
    }
}

//===============================================================================================//
//  Description:
//      Allow the editing of the text field
//
//  Parameters:
//      allowEdit - indicates if can allow edit
//
//  Remarks:
//      Window must not have been created yet
//
//  Returns:
//      void
//===============================================================================================//
void ComboBox::SetAllowEdit( bool allowEdit )
{
    if ( m_hWindow )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    if ( allowEdit )
    {
        SetStyle( CBS_DROPDOWNLIST, false );
        SetStyle( CBS_DROPDOWN, true );
    }
    else
    {
        SetStyle( CBS_DROPDOWNLIST, true );
        SetStyle( CBS_DROPDOWN, false );
    }
}

//===============================================================================================//
//  Description:
//      Set the selected index
//
//  Parameters:
//      selectedIndex - zero-based index of selected item
//
//  Remarks:
//      Window must have been created first
//
//  Returns:
//      void
//
//===============================================================================================//
void ComboBox::SetSelectedIndex( WPARAM selectedIndex ) const
{
    if ( m_hWindow == nullptr  )
    {
        return;     // Nothing to set
    }

    // CB_ERR is a valid response, means selection cleared
    SendMessage( m_hWindow, CB_SETCURSEL, selectedIndex, 0 );
}

//===============================================================================================//
//  Description:
//      Set the selected item as the specified string
//
//  Parameters:
//      SelectedString - string object to match in list
//
//  Remarks:
//      Window must have been created first
//
//  Returns:
//
//===============================================================================================//
void ComboBox::SetSelectedString( const String& SelectedString ) const
{
    WPARAM  selectedIndex = 0;
    LRESULT findString    = 0;

    if ( m_hWindow == nullptr  )
    {
        return;     // Nothing to set or error
    }

    if ( SelectedString.c_str() == nullptr )
    {
        return;     // Nothing to do
    }

    findString = SendMessage( m_hWindow,
                             CB_FINDSTRING,
                             static_cast<WPARAM>( -1 ),
                             reinterpret_cast<LPARAM>(SelectedString.c_str()) );
    if ( findString != CB_ERR )
    {
        selectedIndex = PXSCastLResultToWParam( findString );
        SetSelectedIndex( selectedIndex );
    }
}

//===============================================================================================//
//  Description:
//      Set if want list sorted
//
//  Parameters:
//      sorted - if true list is sorted
//
//  Returns:
//      void
//===============================================================================================//
void ComboBox::SetSorted(bool sorted)
{
    SetStyle( CBS_SORT, sorted );
}

//===============================================================================================//
//  Description:
//      Subclass this ComboBox
//
//  Parameters:
//      none
//
//  Remarks:
//      Window must have been created first
//
//  Returns:
//
//===============================================================================================//
void ComboBox::SubClass()
{
    DWORD   lastError;
    WNDPROC pOldWndProc;

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    // One shot
    if ( GetWindowLongPtr( m_hWindow, DWLP_USER ) )
    {
        return;     // Nothing to do
    }
    pOldWndProc = reinterpret_cast<WNDPROC>( SetWindowLongPtr(
                       m_hWindow,
                       GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(ComboBox::ComboBoxProc ) ) );
    SetLastError( 0 );
    SetWindowLongPtr( m_hWindow,
                      GWLP_USERDATA, reinterpret_cast<LONG_PTR>( pOldWndProc ) );
    lastError = GetLastError();
    if ( lastError )
    {
        throw SystemException( lastError, L"SetWindowLongPtr", __FUNCTION__ );
    }

    // Refresh the window's data
    SetWindowPos( m_hWindow,
                  nullptr,
                  0,
                  0,
                  0,
                  0,
                  SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
LRESULT ComboBox::ComboBoxProc( HWND hWnd,
                                UINT uMsg, WPARAM wParam, LPARAM lParam )
{
    HDC     hdc  = nullptr;
    RECT    rect = { 0, 0, 0, 0 };
    LRESULT lResult = 0;
    HGDIOBJ hOldPen, hOldBrush;
    WNDPROC pOldWndProc;

    // Call default procedure stored in the user data
    pOldWndProc = reinterpret_cast<WNDPROC>( GetWindowLongPtr( hWnd, GWLP_USERDATA ) );
    if ( pOldWndProc )
    {
        lResult = pOldWndProc( hWnd, uMsg, wParam, lParam );
    }

    if ( hWnd == nullptr )
    {
        return lResult;
    }

    // Paint over the 3-D border
    if ( uMsg == WM_PAINT )
    {
        hdc = GetDC( hWnd );
        if ( hdc )
        {
            GetClientRect( hWnd, &rect );
            hOldPen   = SelectObject( hdc, GetStockObject( BLACK_PEN ) );
            hOldBrush = SelectObject( hdc, GetStockObject( HOLLOW_BRUSH ) );
            Rectangle( hdc, 0, 0, rect.right, rect.bottom );

            SelectObject( hdc, GetStockObject( WHITE_PEN ) );
            Rectangle( hdc, 1, 1, rect.right-1, rect.bottom-1 );

            if ( hOldBrush ) SelectObject( hdc, hOldBrush );
            if ( hOldPen   ) SelectObject( hdc, hOldPen );
            ReleaseDC( hWnd, hdc );
        }
    }
    else if ( ( uMsg == WM_COMMAND ) && ( HIWORD( wParam ) == EN_CHANGE ) )
    {
        int     length = 0;
        LRESULT index  = -1;
        WCHAR   szText[ 1024 ] = { 0 };       // Should be enough
        COMBOBOXINFO cbi;

        memset( &cbi, 0, sizeof ( cbi ) );
        cbi.cbSize = sizeof ( cbi );
        GetComboBoxInfo( hWnd, &cbi );

        // Show/hide the drop down. For some reason, CB_SHOWDROPDOWN selects an
        // item and puts it in the edit control, then selects that. Will put
        // back the text then position the caret at the end in the edit
        // control. Also, when the cursor changes to the IDC_IBEAM for editing,
        // then the drop down list is shown, the cursor is either invisible or
        // remains as an IDC_IBEAM, so will reset to the default
        SendMessage( hWnd, WM_GETTEXT, ARRAYSIZE( szText ), (LPARAM)szText );
        szText[ ARRAYSIZE( szText ) - 1 ] = PXS_CHAR_NULL;
        if ( ( szText[ 0 ] ) &&
             ( SendMessage( hWnd, CB_GETCOUNT, 0, 0 ) > 0 ) )   // CB_ERR = -1
        {
            length = lstrlen( szText );
            SendMessage( hWnd, CB_SHOWDROPDOWN, TRUE, 0 );
            SendMessage( hWnd, WM_SETTEXT, 0, (LPARAM)szText );
            SendMessage( cbi.hwndItem, EM_SETSEL, (WPARAM)length, (LPARAM)length );
            SetCursor( LoadCursor( nullptr, IDC_ARROW ) );

            // Select the item in the list box
            index = SendMessage( hWnd, CB_FINDSTRING, (WPARAM)-1, (LPARAM)szText );
            if ( index != -1 )
            {
                SendMessage( cbi.hwndList, LB_SETCURSEL, (WPARAM)index, 0 );
            }
        }
        else
        {
            SendMessage( hWnd, CB_SHOWDROPDOWN, FALSE, 0 );
        }
    }

    return lResult;
}
