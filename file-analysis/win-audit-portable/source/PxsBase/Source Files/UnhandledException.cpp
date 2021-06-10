///////////////////////////////////////////////////////////////////////////////////////////////////
//
// UNhandled Exception Class Implementation
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

///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/UnhandledException.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
UnhandledException::UnhandledException()
                   :Exception( PXS_ERROR_TYPE_EXCEPTION,
                               PXS_ERROR_UNHANDLED_EXCEPTION,
                               nullptr,
                               nullptr )
{
}

// Constructor based on an exception record
UnhandledException::UnhandledException( const EXCEPTION_POINTERS* pException )
{
    // Fill the error codes
    if ( pException && pException->ExceptionRecord )
    {
        m_uErrorType = PXS_ERROR_TYPE_EXCEPTION;
        m_uErrorCode = pException->ExceptionRecord->ExceptionCode;
    }

    try
    {
        FillInDescription();
        FillCrashString( pException );
        FillInMessage();
    }
    catch ( const Exception& )
    { }     // Ignore, no body to tell
}

// Copy constructor
UnhandledException::UnhandledException( const UnhandledException& oUnhandled )
                   :Exception( PXS_ERROR_TYPE_EXCEPTION,
                               PXS_ERROR_UNHANDLED_EXCEPTION,
                               nullptr,
                               nullptr )
{
    *this = oUnhandled;
}

// Destructor
UnhandledException::~UnhandledException()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
UnhandledException& UnhandledException::operator= (
                                         const UnhandledException& oUnhandled )
{
    if ( this == &oUnhandled ) return *this;

    // Base class
    Exception::operator= ( oUnhandled );

    m_uErrorCode  = oUnhandled.m_uErrorCode;
    m_uErrorType  = oUnhandled.m_uErrorType;
    m_CrashString = oUnhandled.m_CrashString;
    m_Description = oUnhandled.m_Description;
    m_Details     = oUnhandled.m_Details;
    m_Message     = oUnhandled.m_Message;

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Method
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
