///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Name-Value Class Implementation
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
#include "PxsBase/Header Files/NameValue.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Exception.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor, an empty name value pair
NameValue::NameValue()
          :m_Name(),
           m_Value()
{
}

// Copy constructor
NameValue::NameValue( const NameValue& oNameValue )
          :m_Name(),
           m_Value()
{
    *this = oNameValue;
}

// Destructor
NameValue::~NameValue()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
NameValue& NameValue::operator= ( const NameValue& oNameValue )
{
    if ( this == &oNameValue ) return *this;

    SetNameValue( oNameValue.GetName(), oNameValue.GetValue() );

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Case insensitive comparison of this class to the input
//
//  Parameters:
//      Input - input name/value pair
//
//  Remarks:
//      For the comparison Name takes precedence over Value
//
//  Returns:
//      > 0  if this class greater than Input
//      = 0  if this class equivalent to Input
//      < 0  if this class less than Input
//===============================================================================================//
int NameValue::CompareI( const NameValue& Input ) const
{
    int nResult = 0;

    if ( m_Name.CompareI( Input.GetName() ) > 0 )
    {
        nResult = 1;
    }
    else if ( m_Name.CompareI( Input.GetName() ) < 0 )
    {
        nResult = -1;
    }
    else
    {
        // Names are the same, compare the value
        if ( m_Value.CompareI( Input.GetValue() ) > 0 )
        {
            nResult = 1;
        }
        else if ( m_Value.CompareI( Input.GetValue() ) < 0 )
        {
            nResult = -1;
        }
    }

    return nResult;
}

//===============================================================================================//
//  Description:
//      Get the name stored in this class
//
//  Parameters:
//      void
//
//  Returns:
//      Constant reference to the name
//===============================================================================================//
const String& NameValue::GetName() const
{
    return m_Name;
}

//===============================================================================================//
//  Description:
//      Get the value stored in this class
//
//  Parameters:
//      void
//
//  Returns:
//      Constant reference to the value
//===============================================================================================//
const String& NameValue::GetValue() const
{
    return m_Value;
}

//===============================================================================================//
//  Description:
//      Reset the name/value pair
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void NameValue::Reset()
{
    m_Name.Zero();
    m_Value.Zero();
}

//===============================================================================================//
//  Description:
//      Set the name and value stored in this class
//
//  Parameters:
//      pwzName  - the name
//      pwzValue - the value
//
//  Returns:
//      void
//===============================================================================================//
void NameValue::SetNameValue( LPCWSTR pwzName, LPCWSTR pwzValue )
{
    m_Name  = pwzName;
    m_Value = pwzValue;
}

//===============================================================================================//
//  Description:
//      Set the name and value stored in this class
//
//  Parameters:
//      Name  - the name
//      Value - the value
//
//  Returns:
//      void
//===============================================================================================//
void NameValue::SetNameValue( const String& Name, const String& Value )
{
    m_Name  = Name;
    m_Value = Value;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
