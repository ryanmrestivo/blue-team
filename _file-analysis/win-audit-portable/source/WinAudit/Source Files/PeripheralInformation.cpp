///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Peripheral Information Class Implementation
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
#include "WinAudit/Header Files/PeripheralInformation.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/TArray.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/DisplayInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
PeripheralInformation::PeripheralInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
PeripheralInformation::~PeripheralInformation()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator - not allowed so no implementation

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//     Get the standard peripherals
//
//  Parameters:
//      pRecords - receives the audit records
//
//  Returns:
//      void
//===============================================================================================//
void PeripheralInformation::GetAuditRecords( TArray< AuditRecord >* pRecords )
{
    AuditRecord MouseRecord, KeyboardRecord, DisplayRecord;
    AuditRecord NetworkRecord;
    TArray< AuditRecord > PrinterRecords, MappedDriveRecords;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Mouse
    try
    {
        GetMouseRecord( &MouseRecord );
        pRecords->Add( MouseRecord );
    }
    catch ( const Exception& e )
    {
        // Log and continue
        PXSLogException( L"Error getting mouse information.", e, __FUNCTION__ );
    }

    // Keyboard
    try
    {
        GetKeyboardRecord( &KeyboardRecord );
        pRecords->Add( KeyboardRecord );
    }
    catch ( const Exception& e )
    {
        // Log and continue
        PXSLogException( L"Error getting keyboard information.", e, __FUNCTION__ );
    }

    // Display
    try
    {
        GetDisplayRecord( &DisplayRecord );
        pRecords->Add( DisplayRecord );
    }
    catch ( const Exception& e )
    {
        // Log and continue
        PXSLogException( L"Error getting display information.", e, __FUNCTION__ );
    }

    // Printers
    try
    {
        GetPrinterRecords( &PrinterRecords );
        pRecords->Append( PrinterRecords );
    }
    catch ( const Exception& e )
    {
        // Log and continue
        PXSLogException( L"Error getting printer information.", e, __FUNCTION__ );
    }

    // Mapped Drives
    try
    {
        GetMappedDriveRecords( &MappedDriveRecords );
        pRecords->Append( MappedDriveRecords );
    }
    catch ( const Exception& e )
    {
        // Log and continue
        PXSLogException( L"Error getting mapped drives information.", e, __FUNCTION__ );
    }

    // Network installed
    try
    {
        GetNetworkRecord( &NetworkRecord );
        pRecords->Add( NetworkRecord );
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error getting network information.", e, __FUNCTION__ );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//     Get information about the display
//
//  Parameters:
//      pDisplayRecord - receives the display data
//
//  Returns:
//      void
//===============================================================================================//
void PeripheralInformation::GetDisplayRecord( AuditRecord* pDisplayRecord )
{
    String  Value;
    DisplayInformation DisplayInfo;

    if ( pDisplayRecord == nullptr )
    {
        throw ParameterException( L"pDisplayRecord", __FUNCTION__ );
    }
    pDisplayRecord->Reset( PXS_CATEGORY_PERIPHERALS );
    DisplayInfo.GetDescriptionString( &Value );
    pDisplayRecord->Add( PXS_PERIPHERALS_NAME, L"Display Description" );
    pDisplayRecord->Add( PXS_PERIPHERALS_DESCRIPTION, Value );
}

//===============================================================================================//
//  Description:
//     Get information about the keyboard
//
//  Parameters:
//      pKeyboardRecord - receives the keyboard data
//
//  Returns:
//      void
//===============================================================================================//
void PeripheralInformation::GetKeyboardRecord( AuditRecord* pKeyboardRecord )
{
    int       keyboardType = 0, numFuncKeys = 0;
    size_t    index = 0;
    String    Value;
    Formatter Format;

    // Keyboard types
    LPCWSTR Types[] = { L"",                                  // 0
                        L"IBM PC/XT or compatible (83-key)",  // 1
                        L"Olivetti \"ICO\" (102-key)",        // 2
                        L"IBM PC/AT (84-key) or similar",     // 3
                        L"IBM enhanced (101- or 102-key)",    // 4
                        L"Nokia 1050 and similar",            // 5
                        L"Nokia 9140 and similar",            // 6
                        L"Japanese" };                        // 7

    // Number of function keys
    LPCWSTR Keys[] = { L"",           // 0
                       L"10",         // 1
                       L"12/18",      // 2
                       L"10",         // 3
                       L"12",         // 4
                       L"10",         // 5
                       L"24" };       // 6

    if ( pKeyboardRecord == nullptr )
    {
        throw ParameterException( L"pKeyboardRecord", __FUNCTION__ );
    }
    pKeyboardRecord->Reset( PXS_CATEGORY_PERIPHERALS );

    keyboardType = GetKeyboardType( 0 );
    if ( keyboardType == 0 )
    {
        throw SystemException( GetLastError(), L"GetKeyboardType(0)", __FUNCTION__ );
    }
    PXSLogAppInfo1( L"Result of GetKeyboardType(0), keyboard type = %%1",
                    Format.Int32( keyboardType ) );

    index = PXSCastInt32ToSizeT( keyboardType );
    if ( index < ARRAYSIZE( Types ) )
    {
        Value += Types[ index ];
    }
    else
    {
        Value = Format.Int32( keyboardType );
    }

    // Function keys
    numFuncKeys = GetKeyboardType( 2 );
    if ( numFuncKeys == 0 )
    {
        throw SystemException( GetLastError(), L"GetKeyboardType(2)", __FUNCTION__ );
    }
    PXSLogAppInfo1( L"Result of GetKeyboardType(2), number function keys = %%1",
                    Format.Int32( numFuncKeys ) );

    // 81 implies USB keyboard so ignore when out of bounds
    index = PXSCastInt32ToSizeT( numFuncKeys );
    if ( index < ARRAYSIZE( Keys ) )
    {
        Value += L", ";
        Value += Keys[ index ];
        Value += L" function keys";
    }
    pKeyboardRecord->Add( PXS_PERIPHERALS_NAME, L"Keyboard" );
    pKeyboardRecord->Add( PXS_PERIPHERALS_DESCRIPTION, Value );
}

//===============================================================================================//
//  Description:
//     Get information about the mapped drives
//
//  Parameters:
//      pMappedDriveRecords - receives the mapped drive data
//
//  Returns:
//      void
//===============================================================================================//
void PeripheralInformation::GetMappedDriveRecords( TArray< AuditRecord >* pMappedDriveRecords )
{
    DWORD  errorCode = 0, bufferSize = 0, cCount = 0, errCloseEnum = 0;
    DWORD  i = 0;
    String Value;
    HANDLE hEnum = nullptr;
    AuditRecord   Record;
    AllocateBytes AllocBytes;
    LPNETRESOURCE pBuffer = nullptr;

    if ( pMappedDriveRecords == nullptr )
    {
        throw ParameterException( L"pMappedDriveRecords", __FUNCTION__ );
    }
    pMappedDriveRecords->RemoveAll();

    // Start enumeration
    errorCode = WNetOpenEnum( RESOURCE_REMEMBERED, RESOURCETYPE_DISK, 0, nullptr, &hEnum );
    if ( errorCode != NO_ERROR )
    {
        throw SystemException( errorCode, L"WNetOpenEnum", __FUNCTION__ );
    }

    // MSDN says 16K should be enough for the maximum number of entries
    bufferSize = 16384;
    pBuffer    = (LPNETRESOURCE)AllocBytes.New( bufferSize );
    cCount     = 0xFFFFFFFF;
    errorCode  = WNetEnumResource( hEnum, &cCount, pBuffer, &bufferSize );

    // First test for no items, then for error
    if ( errorCode == ERROR_NO_MORE_ITEMS )
    {
        errCloseEnum = WNetCloseEnum( hEnum );
        PXSLogAppInfo( L"No mapped drives found." );
        return;
    }

    if ( errorCode != NO_ERROR )
    {
        errCloseEnum = WNetCloseEnum( hEnum );
        if ( errCloseEnum != NO_ERROR )
        {
            PXSLogSysError( errCloseEnum, L"WNetCloseEnum" );
        }
        throw SystemException( errCloseEnum, L"WNetEnumResource", __FUNCTION__ );
    }

    // Need to catch exceptions to close the handle
    try
    {
        for ( i = 0; i < cCount; i++ )
        {
            Value = pBuffer[ i ].lpLocalName;
            if ( pBuffer[ i ].lpRemoteName )
            {
                Value += L" [ ";
                Value += pBuffer[ i ].lpRemoteName;
                Value += L" ]";
            }
            Record.Reset( PXS_CATEGORY_PERIPHERALS );
            Record.Add( PXS_PERIPHERALS_NAME, L"Mapped Drive" );
            Record.Add( PXS_PERIPHERALS_DESCRIPTION, Value );
            pMappedDriveRecords->Add( Record );
        }
    }
    catch ( const Exception& )
    {
        // Clean up
        errCloseEnum = WNetCloseEnum( hEnum );
        if ( errCloseEnum != NO_ERROR )
        {
            PXSLogSysError( errCloseEnum, L"WNetCloseEnum" );
        }
        throw;
    }

    // Clean up
    errCloseEnum = WNetCloseEnum( hEnum );
    if ( errCloseEnum != NO_ERROR )
    {
        PXSLogSysError( errCloseEnum, L"WNetCloseEnum" );
    }
}

//===============================================================================================//
//  Description:
//     Get information about the mouse
//
//  Parameters:
//      pMouseRecord - receives the mouse data
//
//  Returns:
//      void
//===============================================================================================//
void PeripheralInformation::GetMouseRecord( AuditRecord* pMouseRecord )
{
    int       metric = 0;
    String    Value, Temp;
    Formatter Format;

    if ( pMouseRecord == nullptr )
    {
        throw ParameterException( L"pMouseRecord", __FUNCTION__ );
    }
    pMouseRecord->Reset( PXS_CATEGORY_PERIPHERALS );

    Value = PXS_STRING_EMPTY;
    if ( GetSystemMetrics( SM_MOUSEPRESENT ) )
    {
        // One user reported that the 16 buttons were identified on a
        // 5-button mouse, so limit to a reasonable value.
        metric = GetSystemMetrics( SM_CMOUSEBUTTONS );
        Temp = Format.Int32( metric );
        PXSLogAppInfo1( L"GetSystemMetrics: %%1 mouse button(s).", Temp );
        if ( metric <= 5 )
        {
            Value  = Temp;
            Value += L" buttons";
        }

        // Buttons are swapped: Win95+
        if ( GetSystemMetrics( SM_SWAPBUTTON ) )
        {
            if ( Value.GetLength() )
            {
                Value += L", ";
            }
            Value += L"buttons swapped";
        }

        // Mouse wheel available
        if ( GetSystemMetrics( SM_MOUSEWHEELPRESENT ) )
        {
            if ( Value.GetLength() )
            {
                Value += L", ";
            }
            Value += L"has wheel";
        }
    }
    pMouseRecord->Add( PXS_PERIPHERALS_NAME, L"Mouse" );
    pMouseRecord->Add( PXS_PERIPHERALS_DESCRIPTION, Value );
}

//===============================================================================================//
//  Description:
//     Determine if a network is installed
//
//  Parameters:
//      pNetworkRecord - receives the network information
//
//  Returns:
//      void
//===============================================================================================//
void PeripheralInformation::GetNetworkRecord( AuditRecord* pNetworkRecord )
{
    String    Value;
    Formatter Format;

    if ( pNetworkRecord == nullptr )
    {
        throw ParameterException( L"pNetworkRecord", __FUNCTION__ );
    }
    pNetworkRecord->Reset( PXS_CATEGORY_PERIPHERALS );

    Value = Format.Int32Bool( GetSystemMetrics( SM_NETWORK ) );
    pNetworkRecord->Add( PXS_PERIPHERALS_NAME, L"Network Installed" );
    pNetworkRecord->Add( PXS_PERIPHERALS_DESCRIPTION, Value );
}

//===============================================================================================//
//  Description:
//     Get information about the printers
//
//  Parameters:
//      pPrinterRecords - receives the printer data
//
//  Returns:
//      void
//===============================================================================================//
void PeripheralInformation::GetPrinterRecords( TArray< AuditRecord >* pPrinterRecords )
{
    DWORD   cbNeeded = 0, Returned = 0, cbBuf = 0, i = 0, lastError = 0;
    AuditRecord     Record;
    AllocateBytes   AllocBytes;
    PRINTER_INFO_4* pPrinterEnum = nullptr;

    if ( pPrinterRecords == nullptr )
    {
        throw ParameterException( L"pPrinterRecords", __FUNCTION__ );
    }
    pPrinterRecords->RemoveAll();

    // Will only query the local machine at a low information level. Do
    // the first call to get the buffer size
    if ( EnumPrinters( PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS,
                       nullptr,
                       4,            // Level 4
                       nullptr, 0, &cbNeeded, &Returned ) )
    {
        // Success without a buffer, i.e, no printers
        return;
    }

    // Expecting ERROR_INSUFFICIENT_BUFFER
    lastError = GetLastError();
    if ( lastError != ERROR_INSUFFICIENT_BUFFER )
    {
        throw SystemException( lastError, L"EnumPrinters", __FUNCTION__ );
    }

    // Second call, allocate extra in case printers were added
    cbNeeded     = PXSMultiplyUInt32( cbNeeded, 2 );
    pPrinterEnum = reinterpret_cast<PRINTER_INFO_4*>(AllocBytes.New(cbNeeded));
    cbBuf        = cbNeeded;
    cbNeeded     = 0;
    if ( EnumPrinters( PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS,
                        nullptr,
                        4,            // Level 4
                       (LPBYTE)pPrinterEnum,
                       cbBuf, &cbNeeded, &Returned ) == 0 )
    {
        throw SystemException( GetLastError(), L"EnumPrinters", __FUNCTION__ );
    }

    for ( i = 0; i < Returned; i++)
    {
        Record.Reset( PXS_CATEGORY_PERIPHERALS );
        Record.Add( PXS_PERIPHERALS_NAME, L"Printer" );
        Record.Add( PXS_PERIPHERALS_DESCRIPTION, pPrinterEnum[i].pPrinterName);
        pPrinterRecords->Add( Record );
    }
}
