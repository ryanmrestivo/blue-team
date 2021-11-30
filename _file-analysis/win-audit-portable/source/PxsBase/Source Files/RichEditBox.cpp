///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Rich Edit Box Class Implementation
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
#include "PxsBase/Header Files/RichEditBox.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateChars.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Functions
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Callback for the EM_STREAMIN message, see EditStreamCallback Function
//
//  Parameters:
//      dwCookie - pointer to an instance of this class
//
//  Returns:
//      DWORD
//===============================================================================================//
DWORD CALLBACK PXSRichEditStreamInCallback( DWORD_PTR dwCookie,
                                            LPBYTE pbBuff, LONG cb, LONG* pcb )
{
    DWORD error = 0;
    RichEditBox* pRichEditBox;

    if ( ( dwCookie == 0       ) ||
         ( pbBuff   == nullptr ) ||
         ( cb       <= 0       ) ||
         ( pcb      == nullptr )  )
    {
        return ERROR_INVALID_PARAMETER;
    }
    *pcb = 0;

    // Trap exceptions in this callback
    try
    {
        pRichEditBox = reinterpret_cast<RichEditBox*>( dwCookie );
        *pcb = pRichEditBox->StreamInCallback( pbBuff, cb );
    }
    catch ( const Exception& e )
    {
        error = e.GetErrorCode();
    }

    return error;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
RichEditBox::RichEditBox()
            :m_uStreamInOffset( 0 ),
             m_pszStreamInAnsi( nullptr )
{
    // Class registration
    m_WndClassEx.lpszClassName = MSFTEDIT_CLASS;    // Available with XP + SP1

    m_CreateStruct.style     = WS_CHILD     |
                               WS_VISIBLE   |
                               ES_MULTILINE |
                               ES_SAVESEL   | WS_VSCROLL | WS_HSCROLL | WS_BORDER | ES_NOHIDESEL;
    m_CreateStruct.lpszClass = m_WndClassEx.lpszClassName;
}

// Copy constructor - not allowed so no implementation

// Destructor
RichEditBox::~RichEditBox()
{
    // ANSI EM_STREAMIN string
    if ( m_pszStreamInAnsi )
    {
        delete [] m_pszStreamInAnsi;
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
//      Add the specified events to the rich edit's event mask
//
//  Parameters:
//      events - the events to be added to the event mask
//
//  Returns:
//      void
//===============================================================================================//
void RichEditBox::AddToEventMask( LPARAM events )
{
    LPARAM mask;

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }
    mask = SendMessage( m_hWindow, EM_GETEVENTMASK, 0, 0 );
    mask = mask | events;
    SendMessage( m_hWindow, EM_SETEVENTMASK, 0, mask );
}

//===============================================================================================//
//  Description:
//      Append rich text to the end of the window's content's
//
//  Parameters:
//      RichText - the rich text
//
//  Remarks:
//
//  Returns:
//      void
//===============================================================================================//
void RichEditBox::AppendRichText( const String& RichText )
{
    size_t     numChars = 0;
    Formatter  Format;
    EDITSTREAM es;

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    if ( RichText.IsEmpty() )
    {
        return;     // nothing to do
    }

    // Reset the stream in buffer string
    if ( m_pszStreamInAnsi )
    {
        delete [] m_pszStreamInAnsi;
        m_pszStreamInAnsi = nullptr;
    }
    m_uStreamInOffset = 0;

    // Allocate new string
    numChars = RichText.GetAnsiMultiByteLength();
    m_pszStreamInAnsi = new char[ numChars ];
    if ( m_pszStreamInAnsi == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    memset( m_pszStreamInAnsi, 0, numChars );
    Format.StringToAnsi( RichText, m_pszStreamInAnsi, numChars );

    // dwCookie is pointer to instance of this rich edit class
    es.dwCookie    = reinterpret_cast<DWORD_PTR>( this );
    es.dwError     = 0;
    es.pfnCallback = PXSRichEditStreamInCallback;
    SendMessage( m_hWindow, EM_SETSEL, (WPARAM)-1, -1);  // Set selection at end
    SendMessage( m_hWindow, EM_STREAMIN, SF_RTF | SFF_SELECTION, (LPARAM)&es );
    if ( es.dwError )
    {
        // Not necessarily a system error but will assume it is
        throw SystemException( es.dwError, L"EM_STREAMIN", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Find the specified text by searching forward or backward
//
//  Parameters:
//      Text          - the text to search for
//      caseSensitive - true if want a case sensitive search
//      forward       - true to search forward otherwise backward
//
//  Returns:
//      true if found the text otherwise false
//===============================================================================================//
bool RichEditBox::FindText( const String& Text,
                            bool caseSensitive, bool forward, bool fromSelectionStart )
{
    bool      found   = false;
    LPARAM    lParam  = 0;
    WPARAM    wParam  = 0;
    LRESULT   lResult = 0;
    FINDTEXT  ft;
    CHARRANGE cr;

    if ( m_hWindow == nullptr )
    {
        return false;   // Nothing to do or error
    }

    if ( Text.IsEmpty() )
    {
        return false;   // Nothing to do
    }

    if ( caseSensitive )
    {
        wParam = FR_MATCHCASE;
    }

    // Search from the current selection
    memset( &cr, 0, sizeof ( cr ) );
    SendMessage( m_hWindow, EM_EXGETSEL, 0, (LPARAM)&cr );
    if ( forward )
    {
        // Search down to the end starting after current selection
        wParam  |= FR_DOWN;
        if ( fromSelectionStart )
        {
            ft.chrg.cpMin = cr.cpMin;
        }
        else
        {
            ft.chrg.cpMin = cr.cpMax;
        }
        ft.chrg.cpMax = -1;
        ft.lpstrText  = Text.c_str();
        lResult = SendMessage( m_hWindow, EM_FINDTEXT, wParam, (LPARAM)&ft );
        if ( lResult == -1 )
        {
            // Search from top downwards
            ft.chrg.cpMin = 0;
            ft.chrg.cpMax = -1;
            ft.lpstrText  = Text.c_str();
            lResult = SendMessage( m_hWindow, EM_FINDTEXT, wParam, (LPARAM)&ft );
        }
    }
    else
    {
        // Search backward to top
        if ( fromSelectionStart )
        {
            ft.chrg.cpMin = cr.cpMax;
        }
        else
        {
            ft.chrg.cpMin = cr.cpMin;
        }
        ft.chrg.cpMax = -1;
        ft.lpstrText  = Text.c_str();
        lResult = SendMessage( m_hWindow, EM_FINDTEXT, wParam, (LPARAM)&ft );
        if ( lResult == -1 )
        {
            // Set caret at the end and search upwards
            cr.cpMin = -1;
            cr.cpMax = -1;
            SendMessage( m_hWindow, EM_EXSETSEL, 0, (LPARAM)&cr );
            SendMessage( m_hWindow, EM_EXGETSEL, 0, (LPARAM)&cr );
            ft.chrg.cpMin = cr.cpMax;
            ft.chrg.cpMax = 0;
            ft.lpstrText  = Text.c_str();
            lResult = SendMessage( m_hWindow,
                                   EM_FINDTEXT, wParam, (LPARAM)&ft );
        }
    }

    if ( lResult != -1 )
    {
        // Scroll to the text, fist set selection at end so that when select the
        // desired text the caret is at the top of the window.
        lParam = lResult + lstrlen( Text.c_str() );
        SendMessage( m_hWindow, EM_SETSEL, (WPARAM)-1, -1 );  // Set at end
        SendMessage( m_hWindow, EM_SETSEL, (WPARAM)lResult, lParam );
        found = true;
    }

    return found;
}

//===============================================================================================//
//  Description:
//      Determine if any the text is selected
//
//  Parameters:
//      None
//
//  Returns:
//      true if any the text is selected, otherwise false
//===============================================================================================//
bool RichEditBox::IsAnyTextSelected()
{
    CHARRANGE cr;

    if ( m_hWindow == nullptr )
    {
        return false;     // Nothing to do or function error
    }

    memset( &cr, 0, sizeof ( cr ) );
    SendMessage( m_hWindow, EM_EXGETSEL, 0, (LPARAM)&cr );
    if ( cr.cpMin != cr.cpMax )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Set the rich text for this window
//
//  Parameters:
//      RichText - the rich text
//
//  Returns:
//      void
//===============================================================================================//
void RichEditBox::SetRichText( const String& RichText )
{
    char*     pszAnsi  = nullptr;
    size_t    numChars = 0;
    SETTEXTEX set      = { ST_DEFAULT, CP_ACP };  // System default code page
    Formatter Format;
    AllocateChars AnsiString;

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    numChars = RichText.GetAnsiMultiByteLength();
    if ( numChars )
    {
        pszAnsi  = AnsiString.New( numChars );
        Format.StringToAnsi( RichText, pszAnsi, numChars );
        pszAnsi[ numChars - 1 ] = 0x00;
        if ( SendMessage( m_hWindow, EM_SETTEXTEX, (WPARAM)&set, (LPARAM)pszAnsi ) == 0 )
        {
            PXSLogSysError( GetLastError(), L"EM_SETTEXTEX failed." );
        }
    }
}

//===============================================================================================//
//  Description:
//      Entry point for EM_STREAMIN message, see EditStreamCallback Function
//
//  Parameters:
//      pbBuff - buffer to receive the data
//      cb     - size of the buffer
//
//  Returns:
//      number of bytes copied to the buffer
//===============================================================================================//
LONG RichEditBox::StreamInCallback( LPBYTE pbBuff, LONG cb )
{
    size_t bytesCopied, bufferBytes, lenChars;

    if ( pbBuff == nullptr || cb <= 0 )
    {
        return 0;
    }
    bufferBytes = PXSCastLongToSizeT( cb );

    if ( m_pszStreamInAnsi == nullptr )
    {
        return 0;
    }
    lenChars = strlen( m_pszStreamInAnsi );

    // If have copied the entire string, return zero to tell the system to stop
    // calling EditStreamCallback
    if ( m_uStreamInOffset >= lenChars )
    {
        return 0;   // No more data
    }
    bytesCopied = lenChars - m_uStreamInOffset;

    // Copy as much as can fit
    if ( bytesCopied > bufferBytes )
    {
        bytesCopied = bufferBytes;
    }
    memcpy( pbBuff, m_pszStreamInAnsi + m_uStreamInOffset, bytesCopied );
    m_uStreamInOffset = PXSAddSizeT( m_uStreamInOffset, bytesCopied );

    return PXSCastSizeTToLong( bytesCopied );
}

//===============================================================================================//
//  Description:
//      Set if want the text wrapped
//
//  Parameters:
//      wrap - true for wrapped text otherwise false
//
//  Returns:
//      void
//===============================================================================================//
void RichEditBox::WrapText( bool wrap )
{
    // The window must not have been created yet
    if ( m_hWindow )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }
    SetStyle( WS_HSCROLL, !wrap );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
