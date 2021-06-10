///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Accelerator Table Class Header
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

#ifndef PXSBASE_ACCELERATROR_TABLE_H_
#define PXSBASE_ACCELERATROR_TABLE_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

// Accelerator tables are maintained by the system not the application
// that created them, so before the application closes it must call
// DestroyAcceleratorTable to release system resources.

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/PxsBase.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class AcceleratorTable
{
    public:
        // Default constructor
        AcceleratorTable();

        // Destructor
        ~AcceleratorTable();

        // Operators

        // Methods
        void    Add( WORD key, WORD cmd );
        void    Create();
        void    Destroy();
        HACCEL  GetHaccel();
        bool    IsEnabled();
        void    SetEnabled( bool enabled );

    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        AcceleratorTable( const AcceleratorTable& oAcceleratorTable );

        // Assignment operator - not allowed
        AcceleratorTable& operator= ( const AcceleratorTable& oTable );

        // Methods

        // Data members
        bool    m_bEnabled;           // Indicates the table is enabled for
        size_t  m_uTableSize;         // Number of accelerators in the table
        ACCEL*  m_pAcceleratorTable;  // Array containing the accelerators
        HACCEL  m_hAccel;             // Handle to accelerator table
};

#endif  // PXSBASE_ACCELERATROR_TABLE_H_
