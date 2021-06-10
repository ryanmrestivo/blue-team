///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Radio Button Class Implementation
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
#include "PxsBase/Header Files/RadioButton.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/FunctionException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////


// Default constructor
RadioButton::RadioButton()
{
    // Class registration
    m_WndClassEx.lpszClassName = L"BUTTON";

    // Creation parameters, WS_TABSTOP is not required,
    // it is assigned automatically when the button is checked
    m_CreateStruct.cy          = 25;
    m_CreateStruct.cx          = 200;
    m_CreateStruct.style       = WS_VISIBLE | WS_CHILD | BS_AUTORADIOBUTTON | BS_NOTIFY | BS_FLAT;
    m_CreateStruct.lpszClass   = m_WndClassEx.lpszClassName;
    m_CreateStruct.dwExStyle   = WS_EX_TRANSPARENT;
}

// Copy constructor - not allowed so no implementation

// Destructor - do not throw any exceptions
RadioButton::~RadioButton()
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
//      Get the state of this radio button
//
//  Parameters:
//      None
//
//  Returns:
//      true if checked, else false
//===============================================================================================//
bool RadioButton::GetState() const
{
    if ( m_hWindow )
    {
        if ( BST_CHECKED == SendMessage( m_hWindow, BM_GETCHECK, 0, 0 ) )
        {
            return true;
        }
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Set the state of this radio button
//
//  Parameters:
//      state - flag to indicate if radio button is checked
//
//  Returns:
//      void
//===============================================================================================//
void RadioButton::SetState( bool state )
{
    WPARAM wParam = BST_UNCHECKED;

    // Window must have been created
    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    if ( state )
    {
        wParam = BST_CHECKED;
    }
    SendMessage( m_hWindow, BM_SETCHECK, wParam, 0 );
}

//===============================================================================================//
//  Description:
//      Set that this radio button is the start of a new group
//
//  Parameters:
//      None
//
//  Remarks:
//      Must be set before the radio button is created
//
//  Returns:
//      void
//===============================================================================================//
void RadioButton::StartNewGroup()
{
    if ( m_hWindow )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }
    SetStyle( WS_GROUP, true );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
