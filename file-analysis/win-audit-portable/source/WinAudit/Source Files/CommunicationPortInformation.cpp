///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Communication Port Information Class Implementation
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
#include "WinAudit/Header Files/CommunicationPortInformation.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/TArray.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/WindowsInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
CommunicationPortInformation::CommunicationPortInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
CommunicationPortInformation::~CommunicationPortInformation()
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
//      Get communication ports data
//
//  Parameters:
//      pRecords - array to receive port data
//
//  Returns:
//      void
//===============================================================================================//
void CommunicationPortInformation::GetAuditRecords(  TArray< AuditRecord >* pRecords)
{
    BYTE*        pbBuffer = nullptr;
    DWORD        i = 0, lastError = 0, cbBuf = 0, cbNeeded = 0, cReturned = 0;
    String       Type;
    Formatter    Format;
    AuditRecord  Record;
    PORT_INFO_2* pPorts   = nullptr;
    AllocateBytes AllocBytes;
    WindowsInformation   WindowsInfo;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Call EnumPorts for local ports to determine the amount of memory required
    if ( EnumPorts( nullptr, 2, nullptr, 0, &cbNeeded, &cReturned ) )
    {
        // Unexpected success without a buffer
        PXSLogAppInfo( L"EnumPorts said no ports." );
        return;
    }

    // Test for expected error
    lastError = GetLastError();
    if ( lastError != ERROR_INSUFFICIENT_BUFFER )
    {
        throw SystemException( lastError, L"EnumPorts", __FUNCTION__ );
    }

    // Allocate memory and call again. Add some extra in case port(s) created
    // between calls
    cbNeeded = PXSMultiplyUInt32( cbNeeded, 2 );
    pbBuffer = AllocBytes.New( cbNeeded );
    cbBuf    = cbNeeded;
    cbNeeded = 0;
    if ( EnumPorts( nullptr, 2, pbBuffer, cbBuf, &cbNeeded, &cReturned )  == 0 )
    {
        throw SystemException( lastError, L"EnumPorts", __FUNCTION__ );
    }

    pPorts = reinterpret_cast<PORT_INFO_2*>( pbBuffer );
    for ( i = 0; i < cReturned; i++ )
    {
        // Add to output array
        Record.Reset( PXS_CATEGORY_COMMPORTS );
        Record.Add( PXS_COMMPORTS_NUMBER, Format.UInt32( i + 1 ) );
        Record.Add( PXS_COMMPORTS_NAME, pPorts[ i ].pPortName );
        Record.Add( PXS_COMMPORTS_MONITOR_NAME, pPorts[ i ].pMonitorName );
        Record.Add( PXS_COMMPORTS_DESCRIPTION, pPorts[ i ].pDescription );

        Type = PXS_STRING_EMPTY;
        TranslatePortType( pPorts[ i ].fPortType, &Type );
        Record.Add( PXS_COMMPORTS_TYPE, Type );
        pRecords->Add( Record );
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
//      Translate a port type into a string
//
//  Parameters:
//      portType - port type
//      pType    - string object to receive the type
//
//  Remarks:
//
//  Returns:
//      void
//===============================================================================================//
void CommunicationPortInformation::TranslatePortType( DWORD portType, String* pType )
{
    if ( pType == nullptr )
    {
        throw ParameterException( L"pType", __FUNCTION__ );
    }
    *pType = PXS_STRING_EMPTY;

    // PORT_TYPE_WRITE
    if ( portType & PORT_TYPE_WRITE )
    {
        *pType += L"Can write, ";
    }
    else
    {
        *pType += L"Cannot write, ";
    }

    // PORT_TYPE_READ
    if ( portType & PORT_TYPE_READ )
    {
        *pType += L"Can read, ";
    }
    else
    {
        *pType += L"Cannot read, ";
    }

    // PORT_TYPE_REDIRECTED
    if ( portType & PORT_TYPE_REDIRECTED )
    {
        *pType += L"Redirected, ";
    }

    // PORT_TYPE_NET_ATTACHED
    if ( portType & PORT_TYPE_NET_ATTACHED )
    {
        *pType += L"Network attached";
    }

    pType->Trim();
    if ( pType->EndsWithStringI( L"," ) )
    {
        pType->Truncate( pType->GetLength() - 1 );
    }
}
