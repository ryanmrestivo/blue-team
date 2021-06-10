///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Audit Record Class Implementation
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
#include "WinAudit/Header Files/AuditRecord.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/ParameterException.h"

// 5. This Project

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
AuditRecord::AuditRecord()
            :m_uCategoryID( PXS_CATEGORY_UKNOWN ),
             m_Values()
{
}

// Constructor with a category identifier
AuditRecord::AuditRecord( DWORD categoryID )
            :m_uCategoryID( categoryID ),
             m_Values()
{
}

// Copy constructor
AuditRecord::AuditRecord( const AuditRecord& oAuditRecord )
{
    AuditRecord();
    *this = oAuditRecord;
}

// Destructor
AuditRecord::~AuditRecord()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////


// Assignment operator
AuditRecord& AuditRecord::operator=( const AuditRecord& oAuditRecord )
{
    if ( this == &oAuditRecord ) return *this;

    m_uCategoryID = oAuditRecord.m_uCategoryID;
    m_Values      = oAuditRecord.m_Values;

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Add a value to the record
//
//  Parameters:
//      itemID   - the item identifier
//      pszValue - the value
//  Returns:
//      void
//===============================================================================================//
void AuditRecord::Add( DWORD itemID, LPCWSTR pszValue )
{
    String Value;

    Value = pszValue;
    Add( itemID, Value );
}

//===============================================================================================//
//  Description:
//      Add a value to the record
//
//  Parameters:
//      itemID - the item identifier
//      Value - the value
//  Returns:
//      void
//===============================================================================================//
void AuditRecord::Add( DWORD itemID, const String& Value )
{
    size_t    index = 0, newSize = 0;
    String    ErrorMessage;
    Formatter Format;

    // Must have set the category id
    if ( m_uCategoryID == PXS_CATEGORY_UKNOWN )
    {
        throw FunctionException( L"m_uCategoryID", __FUNCTION__ );
    }

    // Verify the item belongs to the category
    if ( ( itemID <=   m_uCategoryID ) ||
         ( itemID >= ( m_uCategoryID + PXS_CATEGORY_INTERVAL ) ) )
    {
        ErrorMessage = Format.StringUInt32_2( L"m_uCategoryID = %%1, itemID = %%2",
                                              m_uCategoryID, itemID );
        throw BoundsException( ErrorMessage.c_str(), __FUNCTION__ );
    }

    // Resize the string array if required
    newSize = itemID - m_uCategoryID;
    if ( m_Values.GetSize() < newSize )
    {
        m_Values.SetSize( newSize );
    }
    index = itemID - m_uCategoryID - 1;    // Zero based

    // Remove any invalid UTF-16 code points before adding to the record
    if ( Value.ReplaceInvalidUTF16( '?' ) )
    {
       PXSLogAppWarn2( L"Invalid UTF16 in '%%1' for item id = %%2.",
                       Value, Format.UInt32( itemID ) );
    }
    m_Values.Set( index, Value.c_str() );
}

//===============================================================================================//
//  Description:
//      Add a value to the record
//
//  Parameters:
//      pCategoryID - receives the category identifier
//      pValues     - receives the string values
//
//  Returns:
//      void
//===============================================================================================//
void AuditRecord::GetCategoryIdAndValues( DWORD* pCategoryID, StringArray* pValues ) const
{
    if ( ( pCategoryID == nullptr ) || ( pValues == nullptr ) )
    {
        throw ParameterException( L"pCategoryID/pValues", __FUNCTION__ );
    }
    *pCategoryID = m_uCategoryID;
    *pValues     = m_Values;
}

//===============================================================================================//
//  Description:
//      Get the category ID
//
//  Parameters:
//      none
//
//  Returns:
//      DWORD
//===============================================================================================//
DWORD AuditRecord::GetCategoryID() const
{
    return m_uCategoryID;
}

//===============================================================================================//
//  Description:
//      Get the number of values in this record
//
//  Parameters:
//      none
//
//  Returns:
//      number or values
//===============================================================================================//
size_t AuditRecord::GetNumberOfValues() const
{
    return m_Values.GetSize();
}

//===============================================================================================//
//  Description:
//      Get the value of the specified audit item
//
//  Parameters:
//      itemID - the audit item's identifier
//      pValue - receives the value
//
//  Returns:
//      void
//===============================================================================================//
void AuditRecord::GetItemValue( DWORD itemID, String* pValue ) const
{
    size_t    index = 0;
    String    ErrorMessage, RecordString;
    Formatter Format;

    // Must have set the category id
    if ( m_uCategoryID == PXS_CATEGORY_UKNOWN )
    {
        throw FunctionException( L"m_uCategoryID", __FUNCTION__ );
    }

    // Verify the item belongs to the category
    if ( ( itemID <=   m_uCategoryID ) ||
         ( itemID >= ( m_uCategoryID + PXS_CATEGORY_INTERVAL ) ) )
    {
       ErrorMessage = Format.StringUInt32_2( L"itemID=%%1, m_uCategoryID=%%2",
                                            itemID, m_uCategoryID );
       throw BoundsException( ErrorMessage.c_str(), __FUNCTION__ );
    }

    // Check bounds
    index = itemID - m_uCategoryID - 1;    // Zero based
    if ( index >= m_Values.GetSize() )
    {
        ToString( &RecordString );
        ErrorMessage  = L"itemID = ";
        ErrorMessage += Format.UInt32( itemID );
        ErrorMessage += L", index = ";
        ErrorMessage += Format.SizeT( index );
        ErrorMessage += L": record = ";
        ErrorMessage += RecordString;
        throw BoundsException( ErrorMessage.c_str(), __FUNCTION__ );
    }

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = m_Values.Get( index );
}

//===============================================================================================//
//  Description:
//      Reset this object
//
//  Parameters:
//      categoryID - the category id
//
//  Returns:
//      void
//===============================================================================================//
void AuditRecord::Reset( DWORD categoryID )
{
    m_uCategoryID = categoryID;
    m_Values.RemoveAll();
}

//===============================================================================================//
//  Description:
//      Convert this record object to a string
//
//  Parameters:
//      pRecordString - receives the formatted audit record
//
//  Remarks
//      In the data ';' are replaced with '?'
//
//  Returns:
//      void
//===============================================================================================//
void AuditRecord::ToString( String* pRecordString ) const
{
    size_t    i = 0, numValues = 0;
    DWORD     itemID = 0;
    String    Value;
    Formatter Format;

    if ( pRecordString == nullptr )
    {
        throw ParameterException( L"pRecordString", __FUNCTION__ );
    }
    pRecordString->Allocate( 1024 );

    Value.Allocate( 256 );
    *pRecordString = Format.StringUInt32( L"CategoryID=%%1;", m_uCategoryID);
    numValues = m_Values.GetSize();
    for ( i = 0; i < numValues; i++ )
    {
        itemID = PXSAddUInt32( m_uCategoryID, 1 );  // One based
        itemID = PXSAddUInt32( itemID, PXSCastSizeTToUInt32( i ) );
        Value  = m_Values.Get( i );
        Value.ReplaceChar( ';', '?' );
        *pRecordString += Format.UInt32( itemID );
        *pRecordString += L"=";
        *pRecordString += Value;
        *pRecordString += L";";
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
