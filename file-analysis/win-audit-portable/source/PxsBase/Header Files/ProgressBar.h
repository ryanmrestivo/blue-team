///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Progress Bar Class Header
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

#ifndef PXSBASE_PROGRESS_BAR_H_
#define PXSBASE_PROGRESS_BAR_H_

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
#include "PxsBase/Header Files/Window.h"

// 6. Forwards
class String;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class ProgressBar : public Window
{
    public:
        // Default constructor
        ProgressBar();

        // Destructor
        ~ProgressBar();

        // Methods
        void    RedrawNow();
        void    SetNote( const String& Note );
        void    SetPercentage( DWORD percentage );
        void    Start( UINT timerInterval );
        void    Stop();

    protected:
        // Methods
        void    PaintEvent( HDC hdc );
        void    TimerEvent( UINT_PTR timerID );

        // Data members

    private:
        // Copy constructor - not allowed
        ProgressBar( const ProgressBar& oProgressBar );

        // Assignment operator - not allowed
        ProgressBar& operator= ( const ProgressBar& oProgressBar);

        // Methods

        // Data members
        const UINT_PTR  TIMER_ID;   // timer ID
        DWORD           m_uMaximum;
        DWORD           m_uStep;
        DWORD           m_uValue;
        String          m_Note;
};

#endif  // PXSBASE_PROGRESS_BAR_H_
