///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Audit Database Class Implementation
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
#include "WinAudit/Header Files/AuditDatabase.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/NameValue.h"
#include "PxsBase/Header Files/Registry.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"

// 5. This Project
#include "WinAudit/Header Files/AuditData.h"
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/OdbcDatabase.h"
#include "WinAudit/Header Files/SmbiosInformation.h"
#include "WinAudit/Header Files/TcpIpInformation.h"
#include "WinAudit/Header Files/WindowsInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
AuditDatabase::AuditDatabase()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
AuditDatabase::~AuditDatabase()
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
//      Get a computer's identifier from the Computer_Master table using
//      various identifying data
//
//  Parameters:
//      ComputerMaster - the computer master record
//
//  Remarks:
//      20/04/2014 - will accept WINAUDIT_GUID alone a reasonable combination
//                   of two identifiers in the event identifiers are duplicated
//                   on different machines.
//
//  Returns:
//      SQLINTEGER of the computer identifier, zero if not found
//===============================================================================================//
SQLINTEGER AuditDatabase::IdentifyComputerID( const AuditRecord& ComputerMaster)
{
    String     UpperCaseKeyword, SqlQuery, WinAuditGuid, WinAuditGuidClean;
    String     MacAddress, MacAddressClean, SmbiosUuid, SmbiosUuidClean;
    String     OsProductID, OsProductIDClean, FQDN, FQDNClean, ComputerName;
    String     ComputerNameClean, SqlUpdateWinAuditGuid;
    Formatter  Format;
    SQLINTEGER computerID = 0;

    // Get if the database is case sensitive
    UpperCaseKeyword = PXS_STRING_EMPTY;
    if ( IsCaseSensitiveSort() )
    {
        GetDbKeyWord( PXS_KEYWORD_UPPER, &UpperCaseKeyword );
    }

    // Look for the WinAuditGuid
    WinAuditGuid      = PXS_STRING_EMPTY;
    WinAuditGuidClean = PXS_STRING_EMPTY;
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_WINAUDIT_GUID , &WinAuditGuid );
    if ( Format.IsValidStringGuid( WinAuditGuid ) == false )
    {
        PXSLogAppInfo1( L"Invalid WinAuditGUID '%%1'", WinAuditGuid );
        throw SystemException( ERROR_INVALID_DATA, PXS_WINAUDIT_COMPUTER_GUID, __FUNCTION__ );
    }
    FixUpStringSQL( WinAuditGuid, SQL_VARCHAR, &WinAuditGuidClean );
    SqlQuery  = L"SELECT Computer_ID FROM COMPUTER_MASTER WHERE ";
    SqlQuery += Format.String2( L"%%1(WINAUDIT_GUID) = %%1(%%2)",
                                UpperCaseKeyword, WinAuditGuidClean );
    ExecuteSelectSqlInteger( SqlQuery, &computerID );
    if ( computerID )
    {
        PXSLogAppInfo( L"Matched computer using WINAUDIT_GUID" );
        return computerID;
    }

    // Did not find the WinAuditGuid. For legacy reasons, search for a computer
    // with matching identifiers that do not have a WinAudidGuid. If a record
    // is found will update the WinAuditGuid
    SqlUpdateWinAuditGuid  = L"UPDATE Computer_Master SET WinAudit_GUID = ";
    SqlUpdateWinAuditGuid += WinAuditGuidClean;
    SqlUpdateWinAuditGuid += L" WHERE Computer_ID = ";

    MacAddress      = PXS_STRING_EMPTY;
    SmbiosUuid      = PXS_STRING_EMPTY;
    OsProductID     = PXS_STRING_EMPTY;
    FQDN            = PXS_STRING_EMPTY;
    ComputerName    = PXS_STRING_EMPTY;
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_MAC_ADDRESS   , &MacAddress   );
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_SMBIOS_UUID   , &SmbiosUuid   );
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_OS_PRODUCT_ID , &OsProductID  );
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_FQDN          , &FQDN         );
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_COMPUTER_NAME , &ComputerName );

    MacAddressClean   = PXS_STRING_EMPTY;
    SmbiosUuidClean   = PXS_STRING_EMPTY;
    OsProductIDClean  = PXS_STRING_EMPTY;
    FQDNClean         = PXS_STRING_EMPTY;
    ComputerNameClean = PXS_STRING_EMPTY;
    FixUpStringSQL( MacAddress   , SQL_VARCHAR, &MacAddressClean   );
    FixUpStringSQL( SmbiosUuid   , SQL_VARCHAR, &SmbiosUuidClean   );
    FixUpStringSQL( OsProductID  , SQL_VARCHAR, &OsProductIDClean  );
    FixUpStringSQL( FQDN         , SQL_VARCHAR, &FQDNClean         );
    FixUpStringSQL( ComputerName , SQL_VARCHAR, &ComputerNameClean );

    // MAC_ADDRESS + SMBIOS_UUID
    if ( MacAddressClean.GetLength() && SmbiosUuidClean.GetLength() )
    {
        SqlQuery  = L"SELECT Computer_ID FROM COMPUTER_MASTER WHERE ";
        SqlQuery += Format.String2( L"%%1(MAC_ADDRESS) = %%1(%%2)",
                                    UpperCaseKeyword, MacAddressClean );
        SqlQuery += Format.String2( L" AND %%1(SMBIOS_UUID) = %%1(%%2)",
                                    UpperCaseKeyword, SmbiosUuidClean );
        SqlQuery += L" AND WINAUDIT_GUID IS NULL";
        ExecuteSelectSqlInteger( SqlQuery, &computerID );
    }
    if ( computerID )
    {
        PXSLogAppInfo( L"Matched computer using MAC_ADDRESS + SMBIOS_UUID" );
        SqlUpdateWinAuditGuid += Format.Int32( computerID );
        ExecuteDirect( SqlUpdateWinAuditGuid );
        return computerID;
    }

    // MAC_ADDRESS + OS_PRODUCT_ID
    if ( MacAddressClean.GetLength() && OsProductIDClean.GetLength() )
    {
        SqlQuery  = L"SELECT Computer_ID FROM COMPUTER_MASTER WHERE ";
        SqlQuery += Format.String2( L"%%1(MAC_ADDRESS) = %%1(%%2)",
                                    UpperCaseKeyword, MacAddressClean );
        SqlQuery += Format.String2( L" AND %%1(OS_PRODUCT_ID) = %%1(%%2)",
                                    UpperCaseKeyword, OsProductIDClean );
        ExecuteSelectSqlInteger( SqlQuery, &computerID );
    }
    if ( computerID )
    {
        PXSLogAppInfo( L"Matched computer using MAC_ADDRESS + OS_PRODUCT_ID" );
        SqlUpdateWinAuditGuid += Format.Int32( computerID );
        ExecuteDirect( SqlUpdateWinAuditGuid );
        return computerID;
    }

    // SMBIOS_UUID + OS_PRODUCT_ID
    if ( SmbiosUuidClean.GetLength() && OsProductIDClean.GetLength() )
    {
        SqlQuery  = L"SELECT Computer_ID FROM COMPUTER_MASTER WHERE ";
        SqlQuery += Format.String2( L"%%1(SMBIOS_UUID) = %%1(%%2)",
                                    UpperCaseKeyword, SmbiosUuidClean );
        SqlQuery += Format.String2( L" AND %%1(OS_PRODUCT_ID) = %%1(%%2)",
                                    UpperCaseKeyword, OsProductIDClean );
        ExecuteSelectSqlInteger( SqlQuery, &computerID );
    }
    if ( computerID )
    {
        PXSLogAppInfo( L"Matched computer using SMBIOS_UUID + OS_PRODUCT_ID" );
        SqlUpdateWinAuditGuid += Format.Int32( computerID );
        ExecuteDirect( SqlUpdateWinAuditGuid );
        return computerID;
    }

    // FULLY_QUALIFIED_DOMAIN_NAME + OS_PRODUCT_ID
    if ( FQDNClean.GetLength() && OsProductIDClean.GetLength() )
    {
        SqlQuery  = L"SELECT Computer_ID FROM COMPUTER_MASTER WHERE ";
        SqlQuery += Format.String2( L"%%1(FULLY_QUALIFIED_DOMAIN_NAME) = %%1(%%2)",
                                    UpperCaseKeyword, FQDNClean );
        SqlQuery += Format.String2( L" AND %%1(OS_PRODUCT_ID) = %%1(%%2)",
                                    UpperCaseKeyword, OsProductIDClean );
        ExecuteSelectSqlInteger( SqlQuery, &computerID );
    }
    if ( computerID )
    {
        PXSLogAppInfo( L"Matched computer using FULLY_QUALIFIED_DOMAIN_NAME + OS_PRODUCT_ID" );
        SqlUpdateWinAuditGuid += Format.Int32( computerID );
        ExecuteDirect( SqlUpdateWinAuditGuid );
        return computerID;
    }

    // COMPUTER_NAME + OS_PRODUCT_ID
    if ( ComputerNameClean.GetLength() && OsProductIDClean.GetLength() )
    {
        SqlQuery  = L"SELECT Computer_ID FROM COMPUTER_MASTER WHERE ";
        SqlQuery += Format.String2( L"%%1(COMPUTER_NAME) = %%1(%%2)",
                                    UpperCaseKeyword, ComputerNameClean );
        SqlQuery += Format.String2( L" AND %%1(OS_PRODUCT_ID) = %%1(%%2)",
                                    UpperCaseKeyword, OsProductIDClean );
        ExecuteSelectSqlInteger( SqlQuery, &computerID );
    }
    if ( computerID )
    {
        PXSLogAppInfo( L"Matched computer using COMPUTER_NAME + OS_PRODUCT_ID" );
        SqlUpdateWinAuditGuid += Format.Int32( computerID );
        ExecuteDirect( SqlUpdateWinAuditGuid );
        return computerID;
    }

    // Log it
    PXSLogAppInfo( L"Did not find matching Computer_ID COMPUTER_MASTER table." );

    return 0;
}

//===============================================================================================//
//  Description:
//      Insert a record into the audit_master table
//
//  Parameters:
//      computerID - the computer identifier for this audit, same as in the
//                   computer_master table
//      AuditMaster- the audit master record
//
//  Remarks:
//      The procedure is pxs_sp_insert_audit_master
//
//      20140328:
//      Bugfix - Czech version of SQL Server 2008 requires timestamps to be
//      ISO8601 conformat (thanks Miroslav P.).
//
//  Returns:
//      SQLINTEGER of the audit identifier
//===============================================================================================//
SQLINTEGER AuditDatabase::InsertAuditMaster( SQLINTEGER computerID,
                                             const AuditRecord& AuditMaster )
{
    String     DbUserName, Value, LocalTime, ComputerUTC, CurrentTimestamp;
    String     DatabaseUTC, AuditGUID, SqlQuery, RecordString, ErrorMessage;
    String     DbmsName;
    Formatter  Format;
    SQLINTEGER auditID = 0;

    if ( computerID <= 0 )
    {
        throw ParameterException( L"computerID", __FUNCTION__ );
    }
    AuditMaster.ToString( &RecordString );

    // PXS_AUDIT_MASTER_AUDIT_GUID
    AuditMaster.GetItemValue( PXS_AUDIT_MASTER_AUDIT_GUID, &Value );
    if ( Format.IsValidStringGuid( Value ) == false )
    {
        ErrorMessage  = L"PXS_AUDIT_MASTER_AUDIT_GUID=";
        ErrorMessage += Value;
        throw SystemException( ERROR_INVALID_DATA, ErrorMessage.c_str(), __FUNCTION__ );
    }
    FixUpStringSQL( Value, SQL_VARCHAR, &AuditGUID );

    // PXS_AUDIT_MASTER_DB_USER_NAME
    Value = GetUserName();
    FixUpStringSQL( Value, SQL_VARCHAR, &DbUserName );

    // PXS_AUDIT_MASTER_DATABASE_LOCAL, if no keywords will use NULL
    GetDbKeyWord( PXS_KEYWORD_CURRENT_TIMESTAMP, &CurrentTimestamp );
    if ( CurrentTimestamp.IsEmpty() )
    {
        CurrentTimestamp = L"NULL";
        PXSLogAppWarn( L"Database keyword CURRENT_TIMESTAMP is unknown." );
    }

    // PXS_AUDIT_MASTER_DATABASE_UTC, Database UTC
    GetDbKeyWord( PXS_KEYWORD_TIMESTAMP, &DatabaseUTC );
    if ( DatabaseUTC.IsEmpty() )
    {
        DatabaseUTC = L"NULL";
        PXSLogAppWarn( L"Database keyword for UTC is unknown." );
    }

    // PXS_AUDIT_MASTER_COMPUTER_LOCAL, Computer local time
    AuditMaster.GetItemValue( PXS_AUDIT_MASTER_COMPUTER_LOCAL, &Value );
    if ( Format.IsValidIsoTimestamp( Value, nullptr ) == false )
    {
        ErrorMessage  = L"PXS_AUDIT_MASTER_COMPUTER_LOCAL=";
        ErrorMessage += Value;
        throw SystemException( ERROR_INVALID_DATA, ErrorMessage.c_str(), __FUNCTION__ );
    }
    FixUpStringSQL( Value, SQL_VARCHAR, &LocalTime );

    // PXS_AUDIT_MASTER_COMPUTER_UTC, Computer UTC
    AuditMaster.GetItemValue( PXS_AUDIT_MASTER_COMPUTER_UTC, &Value );
    if ( Format.IsValidIsoTimestamp( Value, nullptr ) == false )
    {
        ErrorMessage  = L"PXS_AUDIT_MASTER_COMPUTER_UTC=";
        ErrorMessage += Value;
        throw SystemException( ERROR_INVALID_DATA, ErrorMessage.c_str(), __FUNCTION__ );
    }
    FixUpStringSQL( Value, SQL_VARCHAR, &ComputerUTC );

    // SQL Server uses ISO8601 format timestamps. However this is not always
    // the case as it seems to depend on the locale. The format is
    // yyyy-mm-ddThh:mm:ss[.mmm]
    DbmsName = GetDbmsName();
    if ( DbmsName.CompareI( PXS_DBMS_NAME_SQL_SERVER ) == 0 )
    {
        LocalTime.ReplaceCharAt( 11, 'T' );     // Index 11 as string is quoted
        ComputerUTC.ReplaceCharAt( 11, 'T' );
    }

    // Use a stored procedure if supported
    SqlQuery.Allocate( 1024 );
    if ( SupportsStoredProcedures() )
    {
        SqlQuery  = L"{call pxs_sp_insert_audit_master( ";
        SqlQuery += Format.String1( L"%%1, ", AuditGUID );
        SqlQuery += Format.String1( L"%%1, ", DbUserName );
        SqlQuery += Format.String1( L"%%1, ", LocalTime );
        SqlQuery += Format.String1( L"%%1, ", ComputerUTC );
        SqlQuery += Format.Int32( computerID );
        SqlQuery += L" ) }";
    }
    else
    {
        // Use an SQL statement
        SqlQuery  = L"INSERT INTO Audit_Master ( ";
        SqlQuery += L"Audit_GUID, ";
        SqlQuery += L"DB_User_Name,  ";
        SqlQuery += L"Database_Local, ";
        SqlQuery += L"Database_UTC, ";
        SqlQuery += L"Computer_Local, ";
        SqlQuery += L"Computer_UTC, ";
        SqlQuery += L"Computer_ID ";
        SqlQuery += L") VALUES (\r\n";
        SqlQuery += Format.String1( L"%%1, ", AuditGUID );
        SqlQuery += Format.String1( L"%%1, ", DbUserName );
        SqlQuery += Format.String1( L"%%1, ", CurrentTimestamp );
        SqlQuery += Format.String1( L"%%1, ", DatabaseUTC );
        SqlQuery += Format.String1( L"%%1, ", LocalTime );
        SqlQuery += Format.String1( L"%%1, ", ComputerUTC );
        SqlQuery += Format.Int32( computerID );
        SqlQuery += L" )";
    }

    // Test for a transaction otherwise execute directly.
    PXSLogAppInfo1( L"SQL: %%1", SqlQuery );
    if ( StartedTrans() )
    {
        ExecuteTrans( SqlQuery );
    }
    else
    {
        ExecuteDirect( SqlQuery );
    }

    // Get the audit id.
    SqlQuery  = L"SELECT Audit_ID FROM Audit_Master WHERE Audit_GUID = ";
    SqlQuery += AuditGUID;
    ExecuteSelectSqlInteger( SqlQuery, &auditID );
    if ( auditID <= 0 )
    {
        ErrorMessage  = L"auditID <= 0\r\n\r\n";
        ErrorMessage += L"SQL:\r\n";
        ErrorMessage += SqlQuery;
        ErrorMessage += PXS_STRING_CRLF;
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), __FUNCTION__ );
    }

    return auditID;
}


//===============================================================================================//
//  Description:
//      Insert a record into the computer_master table
//
//  Parameters:
//      ComputerMaster - record for the computer_master table
//
//  Remarks:
//      The procedure name is pxs_sp_insert_computer_master
//
//  Returns:
//      SQLINTEGER of the computer identifier
//===============================================================================================//
SQLINTEGER AuditDatabase::InsertComputerMaster( const AuditRecord& ComputerMaster )
{
    String     ComputerGUID, CurrentTimestamp, DbUserName;
    String     Value, MacAddress, SmbiosUUID, AssetTag, FQDN;
    String     SiteName, DomainName, ComputerName, OsProductID;
    String     OtherIdentifier, WinAuditGUID, SqlQuery, ErrorMessage;
    Formatter  Format;
    SQLINTEGER computerID = 0;

    ComputerGUID  = L"'";
    ComputerGUID += Format.CreateGuid();
    ComputerGUID += L"'";

    // CURRENT_TIMESTAMP
    GetDbKeyWord( PXS_KEYWORD_CURRENT_TIMESTAMP, &CurrentTimestamp );
    if ( CurrentTimestamp.IsEmpty() )
    {
        CurrentTimestamp = L"NULL";
        PXSLogAppWarn( L"Database keyword CURRENT_TIMESTAMP is unknown." );
    }

    // User Name
    DbUserName = L"NULL";
    Value      = GetUserName();
    Value.Trim();
    if ( Value.GetLength() )
    {
        FixUpStringSQL( Value, SQL_VARCHAR, &DbUserName );
    }
    ComputerMasterRecordToValues(ComputerMaster,
                                 &MacAddress,
                                 &SmbiosUUID,
                                 &AssetTag,
                                 &FQDN,
                                 &SiteName,
                                 &DomainName,
                                 &ComputerName, &OsProductID, &OtherIdentifier, &WinAuditGUID);

    // Use a stored procedure if supported
    SqlQuery.Allocate( 1024 );
    if ( SupportsStoredProcedures() )
    {
        SqlQuery  = L"{call pxs_sp_insert_computer_master( ";
        SqlQuery += Format.String1( L"%%1, ", ComputerGUID );
        SqlQuery += Format.String1( L"%%1, ", DbUserName );
        SqlQuery += Format.String1( L"%%1, ", MacAddress );
        SqlQuery += Format.String1( L"%%1, ", SmbiosUUID );
        SqlQuery += Format.String1( L"%%1, ", AssetTag );
        SqlQuery += Format.String1( L"%%1, ", FQDN );
        SqlQuery += Format.String1( L"%%1, ", SiteName );
        SqlQuery += Format.String1( L"%%1, ", DomainName );
        SqlQuery += Format.String1( L"%%1, ", ComputerName );
        SqlQuery += Format.String1( L"%%1, ", OsProductID );
        SqlQuery += Format.String1( L"%%1, ", OtherIdentifier );
        SqlQuery += Format.String1( L"%%1  ", WinAuditGUID );
        SqlQuery += L" ) }";
    }
    else
    {
        SqlQuery  = L"INSERT INTO Computer_Master ( ";
        SqlQuery += L"Computer_GUID, ";
        SqlQuery += L"DB_User_Name, ";
        SqlQuery += L"Database_Local, ";
        SqlQuery += L"Last_Audit_ID, ";
        SqlQuery += L"MAC_Address, ";
        SqlQuery += L"Smbios_UUID, ";
        SqlQuery += L"Asset_Tag, ";
        SqlQuery += L"Fully_Qualified_Domain_Name, ";
        SqlQuery += L"Site_Name, ";
        SqlQuery += L"Domain_Name, ";
        SqlQuery += L"Computer_Name, ";
        SqlQuery += L"OS_Product_ID, ";
        SqlQuery += L"Other_Identifier, ";
        SqlQuery += L"WinAudit_GUID ";
        SqlQuery += L" ) VALUES ( ";
        SqlQuery += Format.String1( L"%%1, ", ComputerGUID );
        SqlQuery += Format.String1( L"%%1, ", DbUserName );
        SqlQuery += Format.String1( L"%%1, ", CurrentTimestamp );
        SqlQuery += L"0, ";        // Set the last_audit_id to zero
        SqlQuery += Format.String1( L"%%1, ", MacAddress );
        SqlQuery += Format.String1( L"%%1, ", SmbiosUUID );
        SqlQuery += Format.String1( L"%%1, ", AssetTag );
        SqlQuery += Format.String1( L"%%1, ", FQDN );
        SqlQuery += Format.String1( L"%%1, ", SiteName );
        SqlQuery += Format.String1( L"%%1, ", DomainName );
        SqlQuery += Format.String1( L"%%1, ", ComputerName );
        SqlQuery += Format.String1( L"%%1, ", OsProductID );
        SqlQuery += Format.String1( L"%%1, ", OtherIdentifier );
        SqlQuery += Format.String1( L"%%1 )", WinAuditGUID );
    }

    // Test for a transaction otherwise execute directly.
    if ( StartedTrans() )
    {
        ExecuteTrans( SqlQuery );
    }
    else
    {
        ExecuteDirect( SqlQuery );
    }

    // Get the computer id.
    SqlQuery  = L"SELECT Computer_ID FROM Computer_Master ";
    SqlQuery += L"WHERE Computer_GUID = ";
    SqlQuery += ComputerGUID;
    ExecuteSelectSqlInteger( SqlQuery, &computerID );
    if ( computerID <= 0 )
    {
        ErrorMessage  = L"computerID <= 0\r\n\r\n";
        ErrorMessage += L"SQL:\r\n";
        ErrorMessage += SqlQuery;
        ErrorMessage += PXS_STRING_CRLF;
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), __FUNCTION__ );
    }

    return computerID;
}

//===============================================================================================//
//  Description:
//      Update a record in the computer_master table
//
//  Parameters:
//      ComputerMaster- record for the computer_master table
//      computerID    - the 64-bit computer identifier
//      auditID       - the audit identifier
//
//  Returns:
//      void
//===============================================================================================//
void AuditDatabase::UpdateComputerMaster( const AuditRecord& ComputerMaster,
                                          SQLINTEGER computerID, SQLINTEGER auditID )
{
    String CurrentTimestamp, DbUserName, Value, MacAddress, SmbiosUUID;
    String AssetTag, FQDN, SiteName, DomainName, ComputerName, OsProductID;
    String OtherIdentifier, WinAuditGUID, SqlQuery, AuditIDString;
    Formatter Format;

    if ( computerID <= 0 )
    {
       throw ParameterException( L"computerID", __FUNCTION__ );
    }
    if ( auditID <= 0 )
    {
        throw ParameterException( L"auditID", __FUNCTION__ );
    }
    AuditIDString = Format.Int32( auditID );

    // CURRENT_TIMESTAMP
    GetDbKeyWord( PXS_KEYWORD_CURRENT_TIMESTAMP, &CurrentTimestamp );
    if ( CurrentTimestamp.IsEmpty() )
    {
        CurrentTimestamp = L"NULL";
        PXSLogAppWarn( L"Database keyword CURRENT_TIMESTAMP is unknown." );
    }

    // User Name
    DbUserName = L"NULL";
    Value = GetUserName();
    Value.Trim();
    if ( Value.GetLength() )
    {
        FixUpStringSQL( Value, SQL_VARCHAR, &DbUserName );
    }
    ComputerMasterRecordToValues(ComputerMaster,
                                 &MacAddress,
                                 &SmbiosUUID,
                                 &AssetTag,
                                 &FQDN,
                                 &SiteName,
                                 &DomainName,
                                 &ComputerName, &OsProductID, &OtherIdentifier, &WinAuditGUID);

    SqlQuery.Allocate( 1024 );
    SqlQuery  = L"UPDATE Computer_Master SET ";
    SqlQuery += Format.String1( L"DB_User_Name=%%1, "    , DbUserName );
    SqlQuery += Format.String1( L"Database_Local=%%1, "  , CurrentTimestamp );
    SqlQuery += Format.String1( L"Last_Audit_ID=%%1, "   , AuditIDString );
    SqlQuery += Format.String1( L"MAC_Address=%%1, "     , MacAddress );
    SqlQuery += Format.String1( L"Smbios_UUID=%%1, "     , SmbiosUUID );
    SqlQuery += Format.String1( L"Asset_Tag=%%1, "       , AssetTag );
    SqlQuery += Format.String1( L"Fully_Qualified_Domain_Name=%%1, ", FQDN );
    SqlQuery += Format.String1( L"Site_Name=%%1, "       , SiteName );
    SqlQuery += Format.String1( L"Domain_Name=%%1, "     , DomainName );
    SqlQuery += Format.String1( L"Computer_Name=%%1, "   , ComputerName );
    SqlQuery += Format.String1( L"OS_Product_ID=%%1, "   , OsProductID );
    SqlQuery += Format.String1( L"Other_Identifier=%%1, ", OtherIdentifier );
    SqlQuery += Format.String1( L"WinAudit_GUID=%%1 "    , WinAuditGUID );
    SqlQuery += L"WHERE Computer_ID = ";
    SqlQuery += Format.Int32( computerID );

    // Test for a transaction otherwise execute directly.
    if ( StartedTrans() )
    {
        ExecuteTrans( SqlQuery, 1 );    // Limit to 1 record
    }
    else
    {
        ExecuteDirect( SqlQuery );
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
//      Get the values in the computer_master record
//
//  Parameters:
//      ComputerMaster  - the computer_master record
//      pMacAddress     - receives the MAC Address
//      pSmbiosUUID     - receives the SMBIOS UUID
//      pAssetTag       - receives the Asset Tag
//      pFQDN           - receives the fully qualified domain name
//      pSiteName       - receives the site name
//      pDomainName     - receives the domain name
//      pComputerName   - receives the computer name
//      pOsProductID    - receives the OS Product ID
//      pOtherIdentifier- receives the other identifier
//      pWinAuditGUID   - receives the WinAudit GUID
//
//  Returns:
//      void
//===============================================================================================//
void AuditDatabase::ComputerMasterRecordToValues(
                                              const AuditRecord& ComputerMaster,
                                              String* pMacAddress,
                                              String* pSmbiosUUID,
                                              String* pAssetTag,
                                              String* pFQDN,
                                              String* pSiteName,
                                              String* pDomainName,
                                              String* pComputerName,
                                              String* pOsProductID,
                                              String* pOtherIdentifier,
                                              String* pWinAuditGUID )
{
    String RecordString, Value;
    Formatter Format;

    if ( ( pMacAddress      == nullptr ) ||
         ( pSmbiosUUID      == nullptr ) ||
         ( pAssetTag        == nullptr ) ||
         ( pFQDN            == nullptr ) ||
         ( pSiteName        == nullptr ) ||
         ( pDomainName      == nullptr ) ||
         ( pComputerName    == nullptr ) ||
         ( pOsProductID     == nullptr ) ||
         ( pOtherIdentifier == nullptr ) ||
         ( pWinAuditGUID    == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }

    ComputerMaster.ToString( &RecordString );
    PXSLogAppInfo1( L"Computer_Master: %%1", RecordString );

    // MAC Address
    *pMacAddress = L"NULL";
    Value = PXS_STRING_EMPTY;
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_MAC_ADDRESS, &Value );
    Value.Trim();
    if ( Format.IsValidMacAddress( Value ) )
    {
        FixUpStringSQL( Value, SQL_VARCHAR, pMacAddress );
    }
    else
    {
        PXSLogAppWarn1( L"Invalid MAC address: '%%1'.", Value );
    }

    // SMBIOS UUID
    *pSmbiosUUID = L"NULL";
    Value = PXS_STRING_EMPTY;
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_SMBIOS_UUID, &Value );
    Value.Trim();
    if ( Format.IsValidStringGuid( Value ) )
    {
        FixUpStringSQL( Value, SQL_VARCHAR, pSmbiosUUID );
    }
    else
    {
        PXSLogAppWarn1( L"Invalid SMBIOS UUID: '%%1'.", Value );
    }

    // Asset tag
    *pAssetTag = L"NULL";
    Value    = PXS_STRING_EMPTY;
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_ASSET_TAG, &Value );
    Value.Trim();
    if ( Value.GetLength() )
    {
        FixUpStringSQL( Value, SQL_VARCHAR, pAssetTag );
    }

    // FQDN
    *pFQDN  = L"NULL";
    Value = PXS_STRING_EMPTY;
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_FQDN, &Value );
    Value.Trim();
    if ( Value.GetLength() )
    {
        FixUpStringSQL( Value, SQL_VARCHAR, pFQDN );
    }

    // Site name
    *pSiteName = L"NULL";
    Value    = PXS_STRING_EMPTY;
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_SITE_NAME, &Value );
    Value.Trim();
    if ( Value.GetLength() )
    {
        FixUpStringSQL( Value, SQL_VARCHAR, pSiteName );
    }

    // Domain name
    *pDomainName = L"NULL";
    Value = PXS_STRING_EMPTY;
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_DOMAIN_NAME, &Value );
    Value.Trim();
    if ( Value.GetLength() )
    {
        FixUpStringSQL( Value, SQL_VARCHAR, pDomainName );
    }

    // Computer name
    *pComputerName = L"NULL";
    Value = PXS_STRING_EMPTY;
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_COMPUTER_NAME, &Value );
    Value.Trim();
    if ( Value.GetLength() )
    {
        FixUpStringSQL( Value, SQL_VARCHAR, pComputerName );
    }

    // OS Product ID
    *pOsProductID = L"NULL";
    Value = PXS_STRING_EMPTY;
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_OS_PRODUCT_ID, &Value );
    Value.Trim();
    if ( Value.GetLength() )
    {
        FixUpStringSQL( Value, SQL_VARCHAR, pOsProductID );
    }

    // Other Identifier
    *pOtherIdentifier = L"NULL";
    Value = PXS_STRING_EMPTY;
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_OTHER_ID, &Value );
    Value.Trim();
    if ( Value.GetLength() )
    {
        FixUpStringSQL( Value, SQL_VARCHAR, pOtherIdentifier );
    }

    // WinAudit GUID
    *pWinAuditGUID = L"NULL";
    Value        = PXS_STRING_EMPTY;
    ComputerMaster.GetItemValue( PXS_COMP_MASTER_WINAUDIT_GUID, &Value );
    Value.Trim();
    if ( Format.IsValidStringGuid( Value ) )
    {
        FixUpStringSQL( Value, SQL_VARCHAR, pWinAuditGUID );
    }
    else
    {
        PXSLogAppError1( L"Invalid WinAudit GUID: '%%1'", Value );
    }
}
