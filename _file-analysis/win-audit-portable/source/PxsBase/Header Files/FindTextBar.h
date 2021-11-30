///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Find Text Class Header
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

#ifndef PXSBASE_FIND_TEXT_BAR_H_
#define PXSBASE_FIND_TEXT_BAR_H_

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
#include "PxsBase/Header Files/CheckBox.h"
#include "PxsBase/Header Files/ImageButton.h"
#include "PxsBase/Header Files/Label.h"
#include "PxsBase/Header Files/TextField.h"
#include "PxsBase/Header Files/Window.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class FindTextBar : public Window
{
    public:
        // Default constructor
        FindTextBar();

        // Destructor
        ~FindTextBar();

        // Operators

        // Methods
        void    CreateControls( HWND hWndAppMessageListener );
        void    FocusOnTextBox();
        HWND    GetFindBackwardHwnd() const;
        HWND    GetFindFowardHwnd() const;
        int     GetPreferredHeight() const;
        void    GetSearchParameters( String* pText,
                                     bool* pCaseSensitive ) const;
        HWND    GetTextFieldHwnd() const;
        void    SetCaptions( const String& Find,
                             const String& FindBackwards,
                             const String& FindForwards,
                             const String& MatchCase, const String& TextNotFound );
        void    SetControlColours( COLORREF buttonFillColour );
        void    SetSearchParameters( const String& Text, bool caseSensitive );
        void    ShowTextNotFoundLabel( bool visible );

    protected:
        // Methods
        LRESULT CommandEvent( WPARAM wParam, LPARAM lParam );

        // Data members

    private:
        // Copy constructor - not allowed
        FindTextBar( const FindTextBar& oFindTextBar );

        // Assignment operator - not allowed
        FindTextBar& operator= ( const FindTextBar& oFindTextBar );

        // Methods

        // Data members
        int         FIND_TEXT_BAR_HEIGHT;   // Pseudo constant
        Label       m_Find;
        Label       m_TextNotFound;
        TextField   m_TextField;
        CheckBox    m_CaseSensitive;
        ImageButton m_FindBackward;
        ImageButton m_FindForward;
};

#endif  // PXSBASE_FIND_TEXT_BAR_H_
