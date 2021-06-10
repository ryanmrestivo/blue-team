///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Check Box Class Implementation
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
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/CheckBox.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
CheckBox::CheckBox()
{
    // Class registration
    m_WndClassEx.lpszClassName = L"BUTTON";

    // Creation parameters
    m_CreateStruct.cy        = 26;
    m_CreateStruct.cx        = 80;
    m_CreateStruct.style     = WS_VISIBLE |
                               WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX | BS_NOTIFY | BS_FLAT;
    m_CreateStruct.lpszClass = m_WndClassEx.lpszClassName;
    m_CreateStruct.dwExStyle = WS_EX_TRANSPARENT;
}

// Copy constructor - not allowed so no implementation

// Destructor
CheckBox::~CheckBox()
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
//      Get the state of this check box
//
//  Parameters:
//      None
//
//  Returns:
//      true if checked, else false
//===============================================================================================//
bool CheckBox::GetState() const
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
//      Set the state of this check box
//
//  Parameters:
//      state - flag to indicate if check box is checked
//
//  Returns:
//      void
//===============================================================================================//
void CheckBox::SetState( bool state )
{
    if ( m_hWindow )
    {
        if ( state )
        {
            SendMessage( m_hWindow, BM_SETCHECK, BST_CHECKED, 0 );
        }
        else
        {
            SendMessage( m_hWindow, BM_SETCHECK, BST_UNCHECKED, 0 );
        }
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
