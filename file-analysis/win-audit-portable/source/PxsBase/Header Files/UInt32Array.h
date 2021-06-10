///////////////////////////////////////////////////////////////////////////////////////////////////
//
// UInt32/DWORD Array Class Header
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

#ifndef PXSBASE_UINT32_ARRAY_H_
#define PXSBASE_UINT32_ARRAY_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

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
// Defines
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class UInt32Array
{
    public:
        // Default constructor
        UInt32Array();

        // Copy constructor
        UInt32Array( const UInt32Array& oArray );

        // Destructor
        ~UInt32Array();

       // Assignment operator
        UInt32Array& operator= ( const UInt32Array& oArray );

        // Methods
        void    Add( DWORD un32 );
        void    AddArray( const UInt32Array& oArray );
        bool    AddUnique( DWORD un32 );
        DWORD   Get( size_t index )const;
        size_t  GetSize() const;
        void    RemoveAll();
        void    Set( size_t index, DWORD un32 );
        void    SetSize( size_t newSize );

    protected:
        // Methods

        // Data members

    private:
        // Methods

        // Data members
        size_t  REALLOC_SIZE;     // Pseudo constant
        size_t  m_uSize;
        size_t  m_uAllocated;
        DWORD*  m_pArray;
};

#endif  // PXSBASE_UINT32_ARRAY_H_
