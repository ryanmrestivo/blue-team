///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Text Field Class Implementation
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

//  According to the documentation only the following edit control styles can
//  be changed after the window has been created:
//    ES_LOWERCASE
//    ES_NUMBER
//    ES_OEMCONVERT
//    ES_UPPERCASE
//    ES_WANTRETURN
//

///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/TextField.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/FunctionException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
TextField::TextField()
{
    // Class registration
    m_WndClassEx.lpszClassName = L"EDIT";

    // Creation parameters
    m_CreateStruct.cy          = 20;
    m_CreateStruct.cx          = 200;
    m_CreateStruct.style       = WS_VISIBLE | WS_CHILD | ES_LEFT |
                                 ES_AUTOHSCROLL | WS_TABSTOP | WS_BORDER;
    m_CreateStruct.lpszClass   = m_WndClassEx.lpszClassName;
}

// Copy constructor - not allowed so no implementation

// Destructor
TextField::~TextField()
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
//      Copy the selected text
//
//  Parameters:
//      None
//
//  Remarks:
//      Despite the documentation WM_COPY seems to handle rich text as well
//
//  Returns:
//      void
//===============================================================================================//
void TextField::CopySelection()
{
    if ( m_hWindow == nullptr )
    {
        return;     // Nothing to do
    }
    SendMessage( m_hWindow, WM_COPY, 0, 0 );
}

//===============================================================================================//
//  Description:
//      Select all of the text
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void TextField::RemoveSelection()
{
    if ( m_hWindow == nullptr )
    {
        return;     // Nothing to do
    }
    SendMessage( m_hWindow, EM_SETSEL, 0, 0 );
}

//===============================================================================================//
//  Description:
//      Select all of the text
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void TextField::SelectAll()
{
    if ( m_hWindow == nullptr )
    {
        return;     // Nothing to do
    }
    SendMessage( m_hWindow, EM_SETSEL, 0, -1 );
}

//===============================================================================================//
//  Description:
//      Set if this edit box is to accept digits only
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void TextField::SetDigitsOnly()
{
    SetStyle( ES_NUMBER, true );
}

//===============================================================================================//
//  Description:
//      Set if this edit box is to be read only
//
//  Parameters:
//      readOnly - flag to indicate edit box is read only
//
//  Returns:
//      void
//===============================================================================================//
void TextField::SetReadOnly( bool readOnly )
{
    // Set the style
    SetStyle( ES_READONLY, readOnly );
    if ( m_hWindow )
    {
        SendMessage( m_hWindow, EM_SETREADONLY, (WPARAM)readOnly, 0 );
    }
}

//===============================================================================================//
//  Description:
//      Set the maximum length of the window's text
//
//  Parameters:
//      textMaxLength - maximum character length
//
//  Remarks:
//      Window must have been created
//
//  Returns:
//      void
//===============================================================================================//
void TextField::SetTextMaxLength( size_t textMaxLength )
{
    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    // Maximum is 0x7FFFFFFE characters for single-line edit controls
    if ( textMaxLength > 0x7FFFFFFE )
    {
        textMaxLength = 0x7FFFFFFE;
    }
    SendMessage( m_hWindow, EM_LIMITTEXT, textMaxLength, 0 );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
