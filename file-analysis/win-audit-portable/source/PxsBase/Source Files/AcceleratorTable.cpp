///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Accelerator Table Class Implementation
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

// Usage: First use the Add method for each desired accelerator then call
// the Create method to create the table.

///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/AcceleratorTable.h"

// 2. C System Files
#include <WinUser.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
AcceleratorTable::AcceleratorTable()
                 :m_bEnabled( true ),       // Default is enabled
                  m_uTableSize( 0 ),
                  m_pAcceleratorTable( nullptr ),
                  m_hAccel( nullptr )
{
}

// Copy constructor - not allowed so no implementation

// Destructor
AcceleratorTable::~AcceleratorTable()
{
    try
    {
        Destroy();
    }
    catch( const Exception& e )
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
//      Add an accelerator to the table, see ACCEL
//
//  Parameters:
//      wKey          - WORD representing the key
//      wCommand      - WORD representing the command
//
//  Remarks:
//      Uses virtual flags FCONTROL | FNOINVERT | FVIRTKEY
//
//  Returns:
//      void
//===============================================================================================//
void AcceleratorTable::Add( WORD key, WORD cmd )
{
    size_t newSize = 0;
    ACCEL  accel;
    ACCEL* pNewAccel = nullptr;

    // Fill an ACCEL structure
    memset( &accel, 0, sizeof ( accel ) );
    accel.key   = key;
    accel.cmd   = cmd;
    accel.fVirt = FCONTROL | FNOINVERT | FVIRTKEY;

    // Allocate memory for existing and the new accelerators
    newSize   = PXSAddSizeT( m_uTableSize, 1 );
    pNewAccel = new ACCEL[ newSize ];
    if ( pNewAccel == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }

    // Fill the new array
    if ( m_pAcceleratorTable )
    {
        memcpy( pNewAccel, m_pAcceleratorTable, m_uTableSize * sizeof ( ACCEL ) );
    }
    memcpy( pNewAccel + m_uTableSize, &accel, sizeof ( accel ) );

    // Replace
    delete[] m_pAcceleratorTable;
    m_pAcceleratorTable = pNewAccel;
    m_uTableSize = newSize;
}

//===============================================================================================//
//  Description:
//      Create the accelerator table
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void AcceleratorTable::Create()
{
    int    tableSize = 0;
    HACCEL hAccel;

    if ( m_pAcceleratorTable == nullptr )
    {
        throw FunctionException( L"m_pAcceleratorTable", __FUNCTION__ );
    }
    tableSize = PXSCastSizeTToInt32( m_uTableSize );
    hAccel    = CreateAcceleratorTable( m_pAcceleratorTable, tableSize );
    if ( hAccel == nullptr )
    {
        throw SystemException( GetLastError(), L"CreateAcceleratorTable", __FUNCTION__ );
    }

    // Replace old
    if ( m_hAccel )
    {
        DestroyAcceleratorTable( m_hAccel );
    }
    m_hAccel = hAccel;
}

//===============================================================================================//
//  Description:
//      Destroy the accelerator table
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void AcceleratorTable::Destroy()
{
    if ( m_hAccel )
    {
        DestroyAcceleratorTable( m_hAccel );
    }
    m_hAccel = nullptr;

    if ( m_pAcceleratorTable )
    {
        delete[] m_pAcceleratorTable;
        m_pAcceleratorTable = nullptr;
    }
    m_uTableSize = 0;
}

//===============================================================================================//
//  Description:
//      Get the handle to the accelerator table
//
//  Parameters:
//      None
//
//  Returns:
//      Handle to the created accelerator table, NULL otherwise
//===============================================================================================//
HACCEL AcceleratorTable::GetHaccel()
{
    return m_hAccel;
}

//===============================================================================================//
//  Description:
//      Get if the accelerator table is enabled
//
//  Parameters:
//      None
//
//  Remarks:
//      Message translators can use this method to determine if it
//      should process keyboard accelerator strokes.
//
//  Returns:
//      true if accelerator table is enabled, else false
//===============================================================================================//
bool AcceleratorTable::IsEnabled()
{
    return m_bEnabled;
}

//===============================================================================================//
//  Description:
//      Set if the accelerator table is enabled
//
//  Parameters:
//      enabled - set to false to inform message translators that
//                do not want to process keyboard accelerator strokes
//
//  Returns:
//      void
//===============================================================================================//
void AcceleratorTable::SetEnabled( bool enabled )
{
    m_bEnabled = enabled;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
