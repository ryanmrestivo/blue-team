///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Unlock Mutex Class Header
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

#ifndef PXSBASE_UNLOCK_MUTEX_H_
#define PXSBASE_UNLOCK_MUTEX_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

// Calls Mutex::Unlock when the class goes out of scope

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
class Mutex;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class AutoUnlockMutex
{
    public:
        // Default constructor
        AutoUnlockMutex();

        // Parameter constructor
        explicit AutoUnlockMutex( Mutex* pMutex );

        // Destructor
        ~AutoUnlockMutex();

        // Methods
        void Unlock();

    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        AutoUnlockMutex( const AutoUnlockMutex& oAutoUnlockMutex );

        // Assignment operator - not allowed
        AutoUnlockMutex& operator= ( const AutoUnlockMutex& oAutoUnlockMutex );

        // Methods

        // Data members
        Mutex*  m_pMutex;
};

#endif  // PXSBASE_UNLOCK_MUTEX_H_
