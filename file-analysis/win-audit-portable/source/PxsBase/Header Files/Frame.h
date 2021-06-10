///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Frame Window Class Header
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

#ifndef PXSBASE_FRAME_H_
#define PXSBASE_FRAME_H_

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
#include "PxsBase/Header Files/AcceleratorTable.h"
#include "PxsBase/Header Files/Window.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class Frame : public Window
{
    public:
        // Default constructor
        Frame();

        // Destructor
        ~Frame();

        // Methods
        void    Maximize();
        void    SetFrameBounds( const String& FrameBounds );
        void    SetIconImage( WORD resourceID, bool bigIcon );
        void    SetTitle( LPCWSTR pszTitle );
        void    SetVisible( bool visible );
        DWORD   StartMessageLoop();

    protected:
        // Methods
        void    ActivateEvent( bool activated );
        void    DestroyEvent();
        void    NonClientPaintEvent( HDC hdc );

        // Data members
        AcceleratorTable m_AccelTable;

    private:
        // Copy constructor - not allowed
        Frame( const Frame& oFrame );

        // Assignment operator  - not allowed
        Frame& operator= ( const Frame& oFrame );

        // Methods

        // Data members
        String  m_Title;
        HICON   m_hIcon;                // Window icon - big
        HICON   m_hIconSmall;           // Window icon - small
        HWND    m_hWndChildLastFocus;   // Child window that last had focus
};

#endif  // PXSBASE_FRAME_H_
