///////////////////////////////////////////////////////////////////////////////////////////////////
//
//  Hand Cursor Class Implementation
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
#include "PxsBase/Header Files/HandCursor.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Exception.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
HandCursor::HandCursor()
           :m_hPreviousCursor( nullptr )
{
}

// Copy constructor - not allowed so no implementation

// Destructor
HandCursor::~HandCursor()
{
    // Must not destroy a cursor that was loaded as a resource
    // from an exe in use.
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
//      Restore the cursor to its previous state
//
//  Parameters:
//      None
//
//  Returns:
//      Void
//===============================================================================================//
void HandCursor::Restore()
{
    // Must not set a NULL cursor
    if ( m_hPreviousCursor )
    {
        SetCursor( m_hPreviousCursor );
    }
    else
    {
        // Set the default cursor
        SetCursor( LoadCursor( nullptr, IDC_ARROW ) );
    }
}

//===============================================================================================//
//  Description:
//      Set the cursor to a hand
//
//  Parameters:
//      None
//
//  Remarks:
//      Ignore failures
//
//  Returns:
//      Void
//===============================================================================================//
void HandCursor::Set()
{
    HCURSOR hand = LoadCursor ( nullptr, IDC_HAND );

    // Must never set a NULL cursor
    if ( hand )
    {
        m_hPreviousCursor = SetCursor( hand );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
