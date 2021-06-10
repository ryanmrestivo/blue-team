///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Message Dialog Class Header
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

#ifndef PXSBASE_MESSAGE_DIALOG_H_
#define PXSBASE_MESSAGE_DIALOG_H_

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
#include "PxsBase/Header Files/Button.h"
#include "PxsBase/Header Files/Dialog.h"
#include "PxsBase/Header Files/MenuItem.h"
#include "PxsBase/Header Files/MenuPopup.h"
#include "PxsBase/Header Files/RichEditBox.h"

// 5. This Project

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class MessageDialog : public Dialog
{
    public:
        // Default constructor
        MessageDialog();

        // Destructor
        ~MessageDialog();

        // Methods
        void    SetIsError( bool isError );
        void    SetMessage( const String& Message );

    protected:
        // Methods
        LRESULT CommandEvent( WPARAM wParam, LPARAM lParam );
        void    InitDialogEvent();
        void    NotifyEvent( const NMHDR* pNmhdr );

        // Data members

    private:
        // Copy constructor - not allowed
        MessageDialog( const MessageDialog& oMessageDialog );

        // Assignment operator - not allowed
        MessageDialog& operator= ( const MessageDialog& oMessageDialog );

        // Method

        // Data members
        bool        m_bIsError;
        String      m_Message;
        Button      m_Copy;
        Button      m_OK;
        MenuPopup   m_MenuPopup;
        MenuItem    m_MenuSelectAll;
        MenuItem    m_MenuCopy;
        RichEditBox m_RichEditBox;
};

#endif  // PXSBASE_MESSAGE_DIALOG_H_
