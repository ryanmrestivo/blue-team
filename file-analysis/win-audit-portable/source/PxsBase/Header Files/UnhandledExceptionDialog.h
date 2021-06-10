///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Unhandled Exception Dialog Class Header
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

#ifndef PXSBASE_UNHANDLED_EXCEPTION_DIALOG_H_
#define PXSBASE_UNHANDLED_EXCEPTION_DIALOG_H_

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
#include "PxsBase/Header Files/Dialog.h"
#include "PxsBase/Header Files/TextArea.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class UnhandledExceptionDialog : public Dialog
{
    public:
        // Default constructor
        UnhandledExceptionDialog();

        // Destructor
        ~UnhandledExceptionDialog();

        // Methods
        void    SetExceptionMessage( const String& ExceptionMessage );

    protected:
        // Methods
        LRESULT CommandEvent( WPARAM wParam, LPARAM lParam );
        void    InitDialogEvent();

        // Data members

    private:
        // Copy constructor - not allowed
        UnhandledExceptionDialog( const UnhandledExceptionDialog& oUnhandled );

        // Assignment operator - not allowed
        UnhandledExceptionDialog& operator= ( const UnhandledExceptionDialog& oUnhandled );

        // Methods

        // Data members
        Button   m_MailButton;
        Button   m_ExitButton;
        TextArea m_DetailsText;
        String   m_ExceptionMessage;
};

#endif  // PXSBASE_UNHANDLED_EXCEPTION_DIALOG_H_
