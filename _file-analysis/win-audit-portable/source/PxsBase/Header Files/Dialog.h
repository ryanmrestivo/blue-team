///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Dialog Class Header
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

#ifndef PXSBASE_DIALOG_H_
#define PXSBASE_DIALOG_H_

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
#include "PxsBase/Header Files/Button.h"
#include "PxsBase/Header Files/StringT.h"
#include "PxsBase/Header Files/Window.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class Dialog : public Window
{
    public:
        // Default constructor
        Dialog();

        // Destructor
        ~Dialog();

        // Operators

        // Methods
        INT_PTR         Create( HWND hWndParent );
        void            SetTitle( const String& Title);

    protected:
        // Methods
        static INT_PTR CALLBACK DialogProc( HWND hwndDlg,
                                            UINT uMsg, WPARAM wParam, LPARAM lParam );
        virtual void    DrawButtonItemEvent( const DRAWITEMSTRUCT* pDraw );
        virtual void    InitDialogEvent();
        bool            IsClickFromButton( const Button& ButtonControl,
                                           WPARAM wParam, LPARAM lParam );

        // Data members

    private:
        // Copy constructor - not allowed
        Dialog( const Dialog& oDialog );

        // Assignment operator - not allowed
        Dialog& operator= ( const Dialog& oDialog );

        // Methods
        void    DoDrawItem( const DRAWITEMSTRUCT* pDraw );
        INT_PTR DoInitDialog( HWND hwndDlg, LPARAM lParam );
        INT_PTR DoWmCommand( WPARAM wParam, LPARAM lParam );
        bool    IsButton( HWND hWnd );
        void    SetHwnd( HWND hWindow );

        // Data members
        String  m_Title;
};

#endif  // PXSBASE_DIALOG_H_
