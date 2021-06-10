///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Invalid Function/Method Parameter Exception Class Implementation
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
#include "PxsBase/Header Files/ParameterException.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
ParameterException::ParameterException()
                   :Exception( PXS_ERROR_TYPE_SYSTEM,
                               ERROR_INVALID_PARAMETER,
                               nullptr,
                               nullptr )
{
}

// Constructor with additional details and the throwing method/function
ParameterException::ParameterException( LPCWSTR pszDetails, const char* pszFunction )
                   :Exception( PXS_ERROR_TYPE_SYSTEM,
                               ERROR_INVALID_PARAMETER,
                               pszDetails,
                               pszFunction )
{
}

// Copy constructor
ParameterException::ParameterException( const ParameterException& oParameter )
                   :Exception( PXS_ERROR_TYPE_SYSTEM,
                               ERROR_INVALID_PARAMETER,
                               nullptr,
                               nullptr )
{
    *this = oParameter;
}

// Destructor
ParameterException::~ParameterException()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
ParameterException& ParameterException::operator= ( const ParameterException& oParameter )
{
    if ( this == &oParameter ) return *this;

    // Base class
    Exception::operator= ( oParameter );

    m_uErrorCode  = oParameter.m_uErrorCode;
    m_uErrorType  = oParameter.m_uErrorType;
    m_CrashString = oParameter.m_CrashString;
    m_Description = oParameter.m_Description;
    m_Details     = oParameter.m_Details;
    m_Message     = oParameter.m_Message;

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
