///////////////////////////////////////////////////////////////////////////////////////////////////
//
// WinAudit Global Function Definitions
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
#include "WinAudit/Header Files/WinAudit.h"

// 2. C System Files
#include <shlobj.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/AllocateChars.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/Directory.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/Registry.h"
#include "PxsBase/Header Files/StringT.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/SystemInformation.h"
#include "PxsBase/Header Files/TArray.h"

// 5. This Project
#include "WinAudit/Header Files/AuditData.h"
#include "WinAudit/Header Files/AuditDatabase.h"
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/OdbcExportDialog.h"
#include "WinAudit/Header Files/Resources.h"
#include "WinAudit/Header Files/SmbiosInformation.h"
#include "WinAudit/Header Files/TcpIpInformation.h"
#include "WinAudit/Header Files/WinauditFrame.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Global POD Variables
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Global Functions
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Convert the specified audit records to GUI content
//
//  Parameters:
//      AuditRecords   - the audit records
//      pCategoryItems - receives the contents for the category tree view
//      pTableCounter  - receives the update count of tables in the rich text
//      pRichText      - receives the rich text for the audit report
//
//  Remarks:
//      8900 twips fits WordPad A4 with 1" left and right margins
//
//  Returns:
//      void
//===============================================================================================//
void PXSAuditRecordsToContent( const TArray< AuditRecord >&  AuditRecords,
                               TArray< TreeViewItem >* pCategoryItems,
                               DWORD* pTableCounter, String* pRichText )
{
    const DWORD TABLE_WIDTH = 8900;
    bool   isNode = false, isColumnar = false, tableOpen = false;
    BYTE   indent = 0, leafIndent = 0;
    DWORD  categoryID = 0, itemID = 0, previousCategoryID = 0, captionID = 0;
    size_t i = 0, j = 0, numRecords = 0, numValues = 0, columnWidth = 0;
    String CategoryName, LeafName, ItemName, Value, TableTitle, Separator;
    String HeaderCol_1, HeaderCol_2, RowEven_1, RowEven_2, RowOdd_1, RowOdd_2;
    String CategoryStringData;
    Formatter    Format;
    StringArray  Values;
    AuditRecord  Record;
    TreeViewItem Category;

    // Markup strings
    LPCWSTR STR_SEPARATOR    = L"\\par\\trowd\\trgaph108\\trleft-108"
                               L"\\trrh-60\\clcbpat1\\cellx%%1\\pard\\intbl"
                               L"\\cell\\row\\pard\\par\r\n";
    LPCWSTR STR_HEADER_COL_1 = L"\\clcbpat1\\clbrdrl\\brdrw10\\brdrs"
                               L"\\clbrdrt\\brdrw10\\brdrs\\clbrdrr"
                               L"\\brdrw10\\brdrs\\clbrdrb\\brdrw10\\brdrs"
                               L"\\cellx%%1\r\n";
    LPCWSTR STR_HEADER_COL_2 = L"\\clcbpat1\\clbrdrl\\brdrw10\\brdrs"
                               L"\\clbrdrt\\brdrw10\\brdrs\\clbrdrr"
                               L"\\brdrw10\\brdrs\\clbrdrb\\brdrw10\\brdrs"
                               L"\\cellx%%1\r\n";
    LPCWSTR STR_ROW_EVEN_1   = L"\\clcbpat2\\clbrdrl\\brdrw10\\brdrs"
                               L"\\clbrdrt\\brdrw10\\brdrs\\clbrdrr"
                               L"\\brdrw10\\brdrs\\clbrdrb\\brdrw10\\brdrs"
                               L"\\cellx%%1\r\n";
    LPCWSTR STR_ROW_EVEN_2   = L"\\clcbpat2\\clbrdrl\\brdrw10\\brdrs"
                               L"\\clbrdrt\\brdrw10\\brdrs\\clbrdrr"
                               L"\\brdrw10\\brdrs\\clbrdrb\\brdrw10"
                               L"\\brdrs\\cellx%%1\r\n";
    LPCWSTR STR_ROW_ODD_1    = L"\\clbrdrl\\brdrw10\\brdrs\\clbrdrt"
                               L"\\brdrw10\\brdrs\\clbrdrr\\brdrw10\\brdrs"
                               L"\\clbrdrb\\brdrw10\\brdrs \\cellx%%1\r\n";
    LPCWSTR STR_ROW_ODD_2    = L"\\clbrdrl\\brdrw10\\brdrs\\clbrdrt"
                               L"\\brdrw10\\brdrs\\clbrdrr\\brdrw10"
                               L"\\brdrs\\clbrdrb\\brdrw10\\brdrs"
                               L"\\cellx%%1\r\n";


    if ( ( pCategoryItems == nullptr ) ||
         ( pTableCounter  == nullptr ) ||
         ( pRichText      == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    Separator   = Format.StringUInt32( STR_SEPARATOR   , TABLE_WIDTH );
    HeaderCol_1 = Format.StringUInt32( STR_HEADER_COL_1, TABLE_WIDTH / 3 );
    HeaderCol_2 = Format.StringUInt32( STR_HEADER_COL_2, TABLE_WIDTH );
    RowEven_1   = Format.StringUInt32( STR_ROW_EVEN_1  , TABLE_WIDTH / 3 );
    RowEven_2   = Format.StringUInt32( STR_ROW_EVEN_2  , TABLE_WIDTH );
    RowOdd_1    = Format.StringUInt32( STR_ROW_ODD_1   , TABLE_WIDTH / 3 );
    RowOdd_2    = Format.StringUInt32( STR_ROW_ODD_2   , TABLE_WIDTH );

    // Guesstimate the memory, usually need about 4096 bytes per record
    pCategoryItems->RemoveAll();
    numRecords = AuditRecords.GetSize();
    pRichText->Allocate( numRecords * 4096 );
    *pRichText = PXS_STRING_EMPTY;
    PXSGetRichTextDocumentStart( pRichText );

    // Fill the tree view and rich text report
    for ( i = 0; i < numRecords; i++ )
    {
        Record = AuditRecords.Get( i );
        Record.GetCategoryIdAndValues( &categoryID, &Values );

        // Category tree view
        if ( categoryID != previousCategoryID )
        {
            // New category, if a table is open, close it
            if ( tableOpen )
            {
                *pRichText += L"\\pard\\par\\par\r\n\r\n";
                *pRichText += Separator;
                tableOpen = false;
            }

            // New category
            PXSGetDataCategoryProperties( categoryID,
                                          &CategoryName,
                                          &captionID, &isColumnar, &isNode, &indent );
            *pTableCounter = PXSAddUInt32( *pTableCounter, 1 );
            TableTitle   = Format.UInt32( *pTableCounter );
            TableTitle  += L") ";
            TableTitle  += CategoryName;
            CategoryStringData = TableTitle;    // NB This does not need RTF escaping

            // Title - centred and bold
            TableTitle.EscapeForRichText();
            *pRichText += L"\\par\\qc\\ul\\b ";
            *pRichText += TableTitle;
            *pRichText += L" \\b0\\ul0\\par\\par\\pard\r\n";

            // Treeview
            Category.Reset();
            Category.SetNode( isNode );
            Category.SetIndent( indent );
            Category.SetLabel( CategoryName );
            Category.SetStringData( CategoryStringData );
            pCategoryItems->Add( Category );
        }

        // Message if no data, except for grouping categories
        if ( ( Values.GetSize() == 0 ) && (categoryID > PXS_CATEGORY_INTERVAL) )
        {
            *pRichText += L"\\par\\b No data available "
                          L"\\b0\\par\\par\\pard\r\n";
            *pRichText += Separator;
        }

        numValues = Values.GetSize();
        if ( numValues )
        {
            if ( isColumnar )
            {
                // Only need to a leaf title if have more than 1 record
                // otherwise would have already added one for a node category
                if ( numRecords > 1 && isNode )
                {
                    if ( captionID )
                    {
                        Record.GetItemValue( captionID, &LeafName );
                        *pTableCounter = PXSAddUInt32( *pTableCounter, 1 );
                        TableTitle   = Format.UInt32( *pTableCounter );
                        TableTitle  += L") ";
                        TableTitle  += LeafName;
                    }
                    CategoryStringData = TableTitle;    // NB This does not need RTF escaping

                    // Title - bold
                    TableTitle.EscapeForRichText();
                    *pRichText += L"\\par ";
                    *pRichText += TableTitle;
                    *pRichText += L" \\par\r\n";

                    // Leaf
                    leafIndent = PXSAddUInt8( indent, 1 );
                    Category.Reset();
                    Category.SetNode( false );
                    Category.SetIndent( leafIndent );
                    Category.SetLabel( LeafName );
                    Category.SetStringData( CategoryStringData );
                    pCategoryItems->Add( Category );
                }

                // Table header
                *pRichText += L"\\trowd\\trgaph108\\trleft-108\r\n";
                *pRichText += HeaderCol_1;
                *pRichText += HeaderCol_2;
                *pRichText += L"\\pard\\intbl\r\n";
                *pRichText += L"\\b Item\\b0\\cell\r\n";
                *pRichText += L"\\b Value\\b0\\cell\r\n";
                *pRichText += L"\\row\r\n\r\n";

                // Table rows
                for ( j = 0; j < numValues; j++ )
                {
                    Value = Values.Get( j );
                    Value.EscapeForRichText();
                    itemID = PXSCastSizeTToUInt32( categoryID + j + 1 );
                    PXSGetAuditItemDisplayName( itemID, &ItemName );
                    ItemName.EscapeForRichText();

                    *pRichText += L"\\trowd\\trgaph108\\trleft-108\r\n";
                    if ( j % 2 )
                    {
                        *pRichText += RowEven_1;
                        *pRichText += RowEven_2;
                    }
                    else
                    {
                        *pRichText += RowOdd_1;
                        *pRichText += RowOdd_2;
                    }
                    *pRichText += L"\\pard\\intbl\r\n";
                    *pRichText += ItemName;
                    *pRichText += L"\\cell\r\n";
                    *pRichText += Value;
                    *pRichText += L"\\cell\r\n";
                    *pRichText += L"\\row\r\n\r\n";
                }
                *pRichText += L"\\pard\\par\\par\r\n\r\n";     // Table end
                *pRichText += Separator;
            }
            else
            {
                // Tabular format, append this record as a row
                columnWidth = TABLE_WIDTH / numValues;
                if ( tableOpen == false )
                {
                    // Table header
                    tableOpen = true;
                    *pRichText += L"\\trowd\\trgaph108\\trleft-108\r\n";
                    for ( j = 0; j < numValues; j++ )
                    {
                        *pRichText += L"\\clcbpat1\\clbrdrl\\brdrw10\\brdrs"
                                      L"\\clbrdrt\\brdrw10\\brdrs\\clbrdrr"
                                      L"\\brdrw10\\brdrs\\clbrdrb\\brdrw10"
                                      L"\\brdrs\\cellx";
                        *pRichText += Format.SizeT( ( j + 1 ) * columnWidth );
                        *pRichText += PXS_STRING_CRLF;
                    }
                    *pRichText += L"\\pard\\intbl\r\n";
                    for ( j = 0; j < numValues; j++ )
                    {
                        itemID = PXSCastSizeTToUInt32( categoryID + j + 1 );
                        PXSGetAuditItemDisplayName( itemID, &ItemName );
                        ItemName.EscapeForRichText();
                        *pRichText += L"\\b ";
                        *pRichText += ItemName;
                        *pRichText += L"\\b0\\cell\r\n";
                    }
                    *pRichText += L"\\row\r\n\r\n";
                }

                // Row
                *pRichText += L"\\trowd\\trgaph108\\trleft-108\r\n";
                for ( j = 0; j < numValues; j++ )
                {
                    if ( i % 2 )
                    {
                        *pRichText += L"\\clcbpat2";
                    }
                    *pRichText += L"\\clbrdrl\\brdrw10\\brdrs\\clbrdrt"
                                  L"\\brdrw10\\brdrs\\clbrdrr\\brdrw10"
                                  L"\\brdrs\\clbrdrb\\brdrw10\\brdrs"
                                  L"\\cellx";
                    *pRichText += Format.SizeT( ( j + 1 ) * columnWidth );
                    *pRichText += PXS_STRING_CRLF;
                }
                *pRichText += L"\\pard\\intbl\r\n";
                for ( j = 0; j < numValues; j++ )
                {
                    Value = Values.Get( j );
                    Value.EscapeForRichText();
                    *pRichText += Value;
                    *pRichText += L" \\cell\r\n";
                }
                *pRichText += L"\\row\r\n\r\n";
            }
        }
        previousCategoryID = categoryID;    // Next pass
    }

    // If a table is open, close it
    if ( tableOpen )
    {
        *pRichText += L"\\pard\\par\\par\r\n\r\n";
        *pRichText += Separator;
    }
    *pRichText += L"}";      // Document end
}

//===============================================================================================//
//  Description:
//      Convert the specified audit records to html format
//
//  Parameters:
//      AuditRecords - the audit records
//      pHtmlText    - receives the html text
//
//  Returns:
//      void
//===============================================================================================//
void PXSAuditRecordsToHtml( const TArray< AuditRecord >&  AuditRecords,
                            String* pHtmlText )
{
    bool   isNode = false, isColumnar = false, tableOpen = false;
    BYTE   indent = 0, leafIndent = 0;
    DWORD  categoryID = 0, previousCategoryID = 0, captionID = 0, itemID = 0;
    DWORD  tableCounter = 0;
    size_t i = 0, j = 0, numRecords = 0, numValues = 0;
    String CategoryName, LeafName, ItemName, Value, TableTitle, HtmlBookmarks;
    String ResourceString, ComputerName, ApplicationName, DataString;
    Formatter    Format;
    StringArray  Values;
    AuditRecord  Record;
    SystemInformation SystemInfo;

    // Markup strings
    LPCWSTR STR_HR         = L"<p><hr/></p>\r\n";
    LPCWSTR STR_NBSP       = L"&nbsp;";
    LPCWSTR STR_TABLE      = L"<table width=\"100%\" border=\"1\" "
                             L"cellspacing=\"0\" cellpadding=\"2\">\r\n";
    LPCWSTR STR_HEADER_ROW = L"<tr bgcolor=\"#c2d4fb\"><td width=\"33%\">"
                             L"<b>Item</b></td><td width=\"67%\">"
                             L"<b>Value</b></td></tr>\r\n";
    LPCWSTR STR_BOOKMARK_INDENT = L"&nbsp;&nbsp;&nbsp;";

    if ( pHtmlText == nullptr )
    {
        throw ParameterException( L"pHtmlText", __FUNCTION__ );
    }
    SystemInfo.GetComputerNetBiosName( &ComputerName );
    ComputerName.EscapeForHtml();

    // Guesstimate the memory, usually need about 1024 bytes per record
    numRecords = AuditRecords.GetSize();
    DataString.Allocate( numRecords * 1024 );

    // HTML start
    DataString  = L"<!DOCTYPE html PUBLIC \"-//W3C//DTD "
                  L"XHTML 1.0 Strict//EN\" "
                  L"\"http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd\">";
    DataString += L"<html>\r\n<head>\r\n";
    DataString += L"<title>WinAudit Computer Audit</title>\r\n";
    DataString += L"<meta name=\"author\" content=\"Parmavex "
                  L"Services\"/>\r\n";
    DataString += L"<meta http-equiv=\"Content-Type\" "
                  L"content=\"text/html;\"/>\r\n";
    DataString += L"<style type=\"text/css\">\r\n";
    DataString += L"<!--\r\n";
    DataString += L"body\r\n";
    DataString += L"{\r\n";
    DataString += L"\tmargin-left: 185px;\r\n";
    DataString += L"}\r\n";
    DataString += L"hr\r\n";
    DataString += L"{\r\n";
    DataString += L"\tcolor: #c2d4fb;\r\n";
    DataString += L"\theight: 5px;\r\n";
    DataString += L"}\r\n";
    DataString += L"p\r\n";
    DataString += L"{\r\n";
    DataString += L"\tfont-family: verdana, sans-serif, arial, "
                  L"helvetica;\r\n";
    DataString += L"\tfont-size: 11px;\r\n";
    DataString += L"}\r\n";
    DataString += L"td\r\n";
    DataString += L"{\r\n";
    DataString += L"\tfont-family: verdana, sans-serif, arial, "
                  L"helvetica;\r\n";
    DataString += L"\tfont-size: 11px;\r\n";
    DataString += L"}\r\n";
    DataString += L"div.bookmarks\r\n";
    DataString += L"{\r\n";
    DataString += L"\tbackground-color: #f1f1f1;\r\n";
    DataString += L"\tposition: fixed;\r\n";
    DataString += L"\tleft: 5px;\r\n";
    DataString += L"\ttop: 0px;\r\n";
    DataString += L"\twidth: 175px;\r\n";
    DataString += L"\theight: 100%;\r\n";
    DataString += L"\toverflow:auto;\r\n";
    DataString += L"\twhite-space:nowrap;\r\n";
    DataString += L"}\r\n";
    DataString += L"-->\r\n";
    DataString += L"</style>\r\n";
    DataString += L"</head>\r\n";
    DataString += L"<body vlink=\"#0000ff\">\r\n";
    DataString += L"<center><b><u>Computer Audit for ";
    DataString += ComputerName;
    DataString += L"</u></b></center><p>&nbsp;</p>\r\n";
    for ( i = 0; i < numRecords; i++ )
    {
        Record = AuditRecords.Get( i );
        Record.GetCategoryIdAndValues( &categoryID, &Values );

        // Category tree view
        if ( categoryID != previousCategoryID )
        {
            // New category, if a table is open, close it
            if ( tableOpen )
            {
                DataString += L"</table>\r\n";
                DataString += STR_HR;
                tableOpen = false;
            }

            // New category
            PXSGetDataCategoryProperties( categoryID,
                                          &CategoryName,
                                          &captionID, &isColumnar, &isNode, &indent );
            tableCounter = PXSAddUInt32( tableCounter, 1 );
            TableTitle   = Format.UInt32( tableCounter );
            TableTitle  += L") ";
            TableTitle  += CategoryName;

            // Title - centred and bold
            TableTitle.EscapeForHtml();
            DataString += L"<a name=\"anchor";
            DataString += Format.UInt32( tableCounter );
            DataString += L"\"></a><center><b>";
            DataString += TableTitle;
            DataString += L"</b></center>\r\n";

            // Bookmark
            CategoryName.EscapeForHtml();
            for ( j = 0; j < indent; j++ )
            {
                HtmlBookmarks += STR_BOOKMARK_INDENT;
            }
            HtmlBookmarks += L"<a href=\"#anchor";
            HtmlBookmarks += Format.UInt32( tableCounter );
            HtmlBookmarks += L"\">";
            HtmlBookmarks += CategoryName;
            HtmlBookmarks += L"</a><br/>\r\n";
        }

        // Message if no data, except for grouping categories
        if ( ( Values.GetSize() == 0 ) && (categoryID > PXS_CATEGORY_INTERVAL) )
        {
            DataString += L"<center><b>No data available</b></center>\r\n";
            DataString += STR_HR;
        }

        numValues = Values.GetSize();
        if ( numValues )
        {
            if ( isColumnar )
            {
                // Only need to a leaf title if have more than 1 record
                // otherwise would have already added one for a node category
                if ( ( numRecords > 1 ) && isNode )
                {
                    if ( captionID )
                    {
                        Record.GetItemValue( captionID, &LeafName );
                        tableCounter = PXSAddUInt32( tableCounter, 1 );
                        TableTitle   = Format.UInt32( tableCounter );
                        TableTitle  += L") ";
                        TableTitle  += LeafName;
                    }

                    // Title - bold
                    TableTitle.EscapeForHtml();
                    DataString += L"<a name=\"anchor";
                    DataString += Format.UInt32( tableCounter );
                    DataString += L"\"></a><b>";
                    DataString += TableTitle;
                    DataString += L"</b>\r\n";

                    // Leaf bookmark
                    leafIndent = PXSAddUInt8( indent, 1 );
                    LeafName.EscapeForHtml();
                    for ( j = 0; j < leafIndent; j++ )
                    {
                        HtmlBookmarks += STR_BOOKMARK_INDENT;
                    }
                    HtmlBookmarks += L"<a href=\"#anchor";
                    HtmlBookmarks += Format.UInt32( tableCounter );
                    HtmlBookmarks += L"\">";
                    HtmlBookmarks += LeafName;
                    HtmlBookmarks += L"</a><br/>\r\n";
                }

                // Table header
                DataString += STR_TABLE;
                DataString += STR_HEADER_ROW;

                // Table rows
                for ( j = 0; j < numValues; j++ )
                {
                    Value = Values.Get( j );
                    Value.EscapeForHtml();
                    itemID = PXSCastSizeTToUInt32( categoryID + j + 1 );
                    PXSGetAuditItemDisplayName( itemID, &ItemName );
                    ItemName.EscapeForHtml();
                    if ( j % 2 )
                    {
                        DataString += L"<tr bgcolor=\"#f1f1f1\">";
                    }
                    else
                    {
                        DataString += L"<tr>";
                    }
                    DataString += L"<td>";
                    DataString += ItemName;
                    DataString += L"</td><td>";
                    if ( Value.GetLength() )
                    {
                        DataString += Value;
                    }
                    else
                    {
                        DataString += STR_NBSP;
                    }
                    DataString += L"</td></tr>\r\n";
                }
                DataString += L"</table>\r\n";       // Table end
                DataString += STR_HR;
            }
            else
            {
                // Tabular format, append this record as a row
                if ( tableOpen == false )
                {
                    // Table header
                    tableOpen = true;
                    DataString += STR_TABLE;
                    DataString += L"<tr bgcolor=\"#c2d4fb\">";
                    for ( j = 0; j < numValues; j++ )
                    {
                        itemID = PXSCastSizeTToUInt32( categoryID + j + 1 );
                        PXSGetAuditItemDisplayName( itemID, &ItemName );
                        ItemName.EscapeForHtml();
                        DataString += L"<td width=\"";
                        DataString += Format.SizeT( 100 / numValues );
                        DataString += L"%\"><b>";
                        DataString += ItemName;
                        DataString += L"</b></td>";
                    }
                    DataString += L"</tr>\r\n";
                }

                // Row
                if ( i % 2 )
                {
                    DataString += L"<tr bgcolor=\"#f1f1f1\">";
                }
                else
                {
                    DataString += L"<tr>";
                }
                for ( j = 0; j < numValues; j++ )
                {
                    DataString += L"<td>";
                    Value = Values.Get( j );
                    Value.EscapeForHtml();
                    if ( Value.GetLength() )
                    {
                        DataString += Value;
                    }
                    else
                    {
                        DataString += STR_NBSP;
                    }
                    DataString += L"</td>";
                }
                DataString += L"</tr>\r\n";
            }
        }
        previousCategoryID = categoryID;    // Next pass
    }

    // If a table is open, close it
    if ( tableOpen )
    {
        DataString += L"</table>\r\n";
        DataString += STR_HR;
    }

    // Add the bookmarks as a div
    PXSGetResourceString( PXS_IDS_1160_CATEGORIES, &ResourceString );
    ResourceString.EscapeForHtml();
    DataString += L"<div class=\"bookmarks\">\r\n";
    DataString += L"<p>\r\n";
    DataString += L"<b>";
    DataString += ResourceString;
    DataString += L"<br/></b>";
    DataString += HtmlBookmarks;
    DataString += L"</p>\r\n";
    DataString += L"</div>\r\n";

    // Footer
    PXSGetApplicationName( &ApplicationName );
    ApplicationName.EscapeForHtml();
    DataString += L"<center>";
    DataString += L"<font size=\"2\" color=\"#c0c0c0\">Generated by ";
    DataString += ApplicationName;
    DataString += L"</font></center>\r\n";

    // Document end
    DataString += L"</body>\r\n</html>";

    *pHtmlText = DataString;
}

//===============================================================================================//
//  Description:
//      Convert the specified audit records to csv format
//
//  Parameters:
//      AuditRecords  - the audit records
//      wantHeaderRow - true if want a header row
//      pCsvText      - receives the csv text
//
//  Returns:
//      void
//===============================================================================================//
void PXSAuditRecordsToCsv( const TArray< AuditRecord >&  AuditRecords,
                           bool wantHeaderRow, String* pCsvText )
{
    bool   isColumnar = false, isNode = false;
    BYTE   indent = 0;
    DWORD  categoryID = 0, previousCategoryID = 0, captionID = 0, itemID = 0;
    size_t i = 0, j = 0, numRecords = 0, numValues = 0;
    String Value, CsvLine, CategoryName, ItemName;
    Formatter    Format;
    StringArray  Values;
    AuditRecord  Record;

    if ( pCsvText == nullptr )
    {
        throw ParameterException( L"pCsvText", __FUNCTION__ );
    }
    *pCsvText = PXS_STRING_EMPTY;

    // Guesstimate the memory, about 1024 bytes per record seems reasonable
    numRecords = AuditRecords.GetSize();
    pCsvText->Allocate( numRecords * 512 );
    CsvLine.Allocate( 1024 );
    for ( i = 0; i < numRecords; i++ )
    {
        Record = AuditRecords.Get( i );
        Record.GetCategoryIdAndValues( &categoryID, &Values );

        // Make header row
        if ( categoryID != previousCategoryID )
        {
            PXSGetDataCategoryProperties( categoryID,
                                          &CategoryName,
                                          &captionID, &isColumnar, &isNode, &indent );
            // Escape any quotes then quote the strings
            CategoryName.ReplaceChar( PXS_CHAR_QUOTE, L"\"\"" );
            if ( wantHeaderRow )
            {
                CsvLine  = PXS_CHAR_QUOTE;
                CsvLine += CategoryName;
                CsvLine += PXS_CHAR_QUOTE;
                numValues = Values.GetSize();
                for ( j = 0; j < numValues; j++ )
                {
                    itemID = PXSCastSizeTToUInt32( categoryID + j + 1 );
                    PXSGetAuditItemDisplayName( itemID, &ItemName );
                    ItemName.ReplaceChar( PXS_CHAR_QUOTE, L"\"\"" );
                    CsvLine += PXS_CHAR_COMMA;
                    CsvLine += PXS_CHAR_QUOTE;
                    CsvLine += ItemName;
                    CsvLine += PXS_CHAR_QUOTE;
                }
                CsvLine += PXS_STRING_CRLF;
                *pCsvText += CsvLine;
            }
        }

        // Make the data row
        CsvLine  = PXS_CHAR_QUOTE;
        CsvLine += Format.UInt32( categoryID );
        CsvLine += PXS_CHAR_QUOTE;
        CsvLine += PXS_CHAR_COMMA;
        CsvLine += PXS_CHAR_QUOTE;
        CsvLine += CategoryName;
        CsvLine += PXS_CHAR_QUOTE;
        numValues = Values.GetSize();
        for ( j = 0; j < numValues; j++ )
        {
            // Escape any quotes then quote the value
            Value = Values.Get( j );
            Value.ReplaceChar( PXS_CHAR_TAB, PXS_CHAR_SPACE );
            Value.ReplaceChar( PXS_CHAR_QUOTE, L"\"\"" );
            CsvLine += PXS_CHAR_COMMA;
            CsvLine += PXS_CHAR_QUOTE;
            CsvLine += Value;
            CsvLine += PXS_CHAR_QUOTE;
        }
        CsvLine += PXS_STRING_CRLF;
        *pCsvText += CsvLine;

        previousCategoryID = categoryID;    // Next pass
    }
}

//===============================================================================================//
//  Description:
//      Convert the specified audit records to csv columnar format
//
//  Parameters:
//      AuditRecords  - the audit records
//      wantHeaderRow - true if want a header row
//      pCsvText      - receives the csv text
//
//  Returns:
//      void
//===============================================================================================//
void PXSAuditRecordsToCsv2( const TArray< AuditRecord >&  AuditRecords,
                            bool wantHeaderRow, String* pCsvText )
{
    bool   isColumnar = false, isNode = false;
    BYTE   indent = 0;
    DWORD  categoryID = 0, previousCategoryID = 0, captionID = 0, itemID = 0, itemOrder = 0;
    size_t i = 0, j = 0, numRecords = 0, numValues = 0;
    String Value, CsvLine, CategoryName, ItemName;
    Formatter    Format;
    StringArray  Values;
    AuditRecord  Record;

    if ( pCsvText == nullptr )
    {
        throw ParameterException( L"pCsvText", __FUNCTION__ );
    }
    *pCsvText = PXS_STRING_EMPTY;

    // Make header row
    if ( wantHeaderRow )
    {
        *pCsvText = L"ItemOrder,RecordNumber,CategoryID,CategoryName,ItemID,ItemName,ItemValue\r\n";
    }

    // Guesstimate the memory, about 1024 bytes per record seems reasonable
    numRecords = AuditRecords.GetSize();
    pCsvText->Allocate( numRecords * 512 );
    CsvLine.Allocate( 1024 );
    for ( i = 0; i < numRecords; i++ )
    {
        Record = AuditRecords.Get( i );
        Record.GetCategoryIdAndValues( &categoryID, &Values );

        if ( categoryID != previousCategoryID )
        {
            PXSGetDataCategoryProperties( categoryID,
                                          &CategoryName,
                                          &captionID, &isColumnar, &isNode, &indent );
            // Escape any quotes then quote the strings
            CategoryName.ReplaceChar( PXS_CHAR_QUOTE, L"\"\"" );
        }

        // Make the data rows
        numValues = Values.GetSize();
        for ( j = 0; j < numValues; j++ )
        {
            // ItemOrder
            itemOrder = PXSAddUInt32( itemOrder, 1 );
            CsvLine   = Format.UInt32( itemOrder );
            CsvLine  += PXS_CHAR_COMMA;

            // RecordNumber
            CsvLine  += Format.SizeT( i + 1 );
            CsvLine  += PXS_CHAR_COMMA;

            // CategoryID
            CsvLine += Format.UInt32( categoryID );
            CsvLine += PXS_CHAR_COMMA;

            // CategoryName
            CsvLine += PXS_CHAR_QUOTE;
            CsvLine += CategoryName;
            CsvLine += PXS_CHAR_QUOTE;
            CsvLine += PXS_CHAR_COMMA;

            // ItemID
            itemID   = PXSCastSizeTToUInt32( categoryID + j + 1 );
            CsvLine += Format.UInt32( itemID );
            CsvLine += PXS_CHAR_COMMA;

            // ItemName
            ItemName = PXS_STRING_EMPTY;
            PXSGetAuditItemDisplayName( itemID, &ItemName );
            ItemName.ReplaceChar( PXS_CHAR_QUOTE, L"\"\"" );
            CsvLine += PXS_CHAR_QUOTE;
            CsvLine += ItemName;
            CsvLine += PXS_CHAR_QUOTE;
            CsvLine += PXS_CHAR_COMMA;

            // ItemValue
            Value = Values.Get( j );
            Value.ReplaceChar( PXS_CHAR_TAB, PXS_CHAR_SPACE );
            Value.ReplaceChar( PXS_CHAR_QUOTE, L"\"\"" );
            CsvLine += PXS_CHAR_QUOTE;
            CsvLine += Value;
            CsvLine += PXS_CHAR_QUOTE;

            CsvLine += PXS_STRING_CRLF;
            *pCsvText += CsvLine;
        }

        previousCategoryID = categoryID;    // Next pass
    }
}


//===============================================================================================//
//  Description:
//     bsearch comparison function for array of type PXS_TYPE_AUDIT_ITEM
//
//  Parameters:
//      pItem1 - pointer to the first item
//      pItem2 - pointer to the second item
//
//  Returns:
//      < 0 pItem1 less than pItem2
//      = 0 pItem1 equal to pItem2
//      > 0 pItem1 greater than pItem2
//===============================================================================================//
int PXSBSearchCompareAuditItems( const void* pItem1, const void* pItem2 )
{
    int nReturn = 0;
    const PXS_TYPE_AUDIT_ITEM* p1 = nullptr;
    const PXS_TYPE_AUDIT_ITEM* p2 = nullptr;

    if ( pItem1 && pItem2 )
    {
        p1 = reinterpret_cast<const PXS_TYPE_AUDIT_ITEM*>( pItem1 );
        p2 = reinterpret_cast<const PXS_TYPE_AUDIT_ITEM*>( pItem2 );
        if ( p1->itemID > p2->itemID )
        {
            nReturn = 1;
        }
        else if ( p1->itemID < p2->itemID )
        {
            nReturn = -1;
        }
    }
    else if ( pItem1 && ( pItem2 == nullptr ) )
    {
        nReturn = 1;
    }
    else if ( ( pItem1 == nullptr ) && pItem2 )
    {
        nReturn = -1;
    }

    return nReturn;
}

//===============================================================================================//
//  Description:
//     bsearch comparison function for array of type PXS_TYPE_PORT_SERVICE
//
//  Parameters:
//      pPort1 - pointer to the first port service
//      pPort2 - pointer to the second port service
//
//  Returns:
//      < 0 pPort1 less than pPort2
//      = 0 pPort1 equal to pPort2
//      > 0 pPort1 greater than pPort2
//===============================================================================================//
int PXSBSearchComparePortServices( const void* pPort1, const void* pPort2 )
{
    int nReturn = 0;
    const PXS_TYPE_PORT_SERVICE* p1 = nullptr;
    const PXS_TYPE_PORT_SERVICE* p2 = nullptr;

    if ( pPort1 && pPort2 )
    {
        p1 = reinterpret_cast<const PXS_TYPE_PORT_SERVICE*>( pPort1 );
        p2 = reinterpret_cast<const PXS_TYPE_PORT_SERVICE*>( pPort2 );
        if ( p1->portID > p2->portID )
        {
            nReturn = 1;
        }
        else if ( p1->portID < p2->portID )
        {
             nReturn = -1;
        }
    }
    else if ( pPort1 && ( pPort2 == nullptr ) )
    {
        nReturn = 1;
    }
    else if ( ( pPort1 == nullptr ) && pPort2 )
    {
        nReturn = -1;
    }

    return nReturn;
}

//===============================================================================================//
//  Description:
//      Get the name of an audit item for display purposes
//
//  Parameters:
//      itemID       - defined constant identifying the item
//      pDisplayName - string to receive the name
//
//  Returns:
//      void
//===============================================================================================//
void PXSGetAuditItemDisplayName( DWORD itemID, String* pDisplayName )
{
    void* pResult = nullptr;
    String    ErrorMessage;
    Formatter Format;
    SmbiosInformation   SmbiosInfo;
    PXS_TYPE_AUDIT_ITEM* pAuditItem = nullptr;

    if ( pDisplayName == nullptr )
    {
        throw ParameterException( L"pDisplayName", __FUNCTION__ );
    }
    *pDisplayName = PXS_STRING_EMPTY;

    // First look in PXS_AUDIT_ITEMS
    pResult = bsearch( &itemID,
                        PXS_AUDIT_ITEMS,
                        ARRAYSIZE( PXS_AUDIT_ITEMS ),
                        sizeof ( PXS_AUDIT_ITEMS[ 0 ] ),
                        PXSBSearchCompareAuditItems );
    pAuditItem = reinterpret_cast<PXS_TYPE_AUDIT_ITEM*>( pResult );
    if ( pAuditItem )
    {
        *pDisplayName = pAuditItem->pszName;
    }
    else
    {
        // Then look for SMBIOS specified items
        SmbiosInfo.GetItemName( itemID, pDisplayName );
    }

    if ( pDisplayName->IsEmpty() )
    {
        ErrorMessage = Format.StringUInt32( L"itemID = %%1.", itemID );
        throw ParameterException( ErrorMessage.c_str(), __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Get the values associated with the command line switches
//
//  Parameters:
//      Switches              - the switches
//      pReportSwitchValue    - receives the /r= value
//      pFileSwitchValue      - receives the /f= value
//      pLogSwitchValue       - receives the /l= value
//      pTimestampSwitchValue - receives the /T= value
//      pLanguageSwitchValue  - receives the /L= value
//
//  Returns:
//      void
//===============================================================================================//
void PXSGetCommandLineSwitchValues( const StringArray& Switches,
                                    String* pReportSwitchValue,
                                    String* pFileSwitchValue,
                                    String* pLogSwitchValue,
                                    String* pTimestampSwitchValue, String* pLanguageSwitchValue )
{
    size_t i = 0;
    size_t numSwitches = Switches.GetSize();
    String Switch;

    if ( ( pReportSwitchValue    == nullptr ) ||
         ( pFileSwitchValue      == nullptr ) ||
         ( pLogSwitchValue       == nullptr ) ||
         ( pTimestampSwitchValue == nullptr ) ||
         ( pLanguageSwitchValue  == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    *pReportSwitchValue    = PXS_STRING_EMPTY;
    *pFileSwitchValue      = PXS_STRING_EMPTY;
    *pLogSwitchValue       = PXS_STRING_EMPTY;
    *pTimestampSwitchValue = PXS_STRING_EMPTY;
    *pLanguageSwitchValue  = PXS_STRING_EMPTY;

    // Identify the switches
    for ( i = 0; i < numSwitches; i++ )
    {
        Switch = Switches.Get( i );
        Switch.Trim();
        if ( Switch.StartsWith( L"r=", true ) )
        {
            Switch.SubString( 2, PXS_MINUS_ONE, pReportSwitchValue );
            pReportSwitchValue->Trim();
        }
        else if ( Switch.StartsWith( L"f=", true ) )
        {
            Switch.SubString( 2, PXS_MINUS_ONE, pFileSwitchValue );
            PXSUnQuoteString( pFileSwitchValue );
            pFileSwitchValue->Trim();
        }
        else if ( Switch.StartsWith( L"l=", true ) )
        {
            Switch.SubString( 2, PXS_MINUS_ONE, pLogSwitchValue );
            PXSUnQuoteString( pLogSwitchValue );
            pLogSwitchValue->Trim();
        }
        else if ( Switch.StartsWith( L"T=", true ) )
        {
            Switch.SubString( 2, PXS_MINUS_ONE, pTimestampSwitchValue );
            pTimestampSwitchValue->Trim();
        }
        else if ( Switch.StartsWith( L"L=", true ) )
        {
            Switch.SubString( 2, PXS_MINUS_ONE, pLanguageSwitchValue );
            pLanguageSwitchValue->Trim();
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the category properties for the specified category
//
//  Parameters:
//      categoryID    - the category ID
//      pCategoryName - receives the category name
//      pCaptionID    - receives the item id used as the caption for the record
//      pIsColumnar   - receives if the data is columnar
//
//  Returns:
//      void
//===============================================================================================//
void PXSGetDataCategoryProperties( DWORD categoryID,
                                   String* pCategoryName,
                                   DWORD* pCaptionID,
                                   bool* pIsColumnar, bool* pIsNode, BYTE* pIndent )
{
    bool      found = false;
    size_t    i = 0;
    String    ErrorMessage;
    Formatter Format;

    if ( ( pCategoryName == nullptr ) ||
         ( pCaptionID    == nullptr ) ||
         ( pIsColumnar   == nullptr ) ||
         ( pIsNode       == nullptr ) ||
         ( pIndent       == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    *pCategoryName = PXS_STRING_EMPTY;
    *pCaptionID  = 0;
    *pIsColumnar = false;
    *pIsNode     = false;
    *pIndent     = 0;
    for( i = 0; i < ARRAYSIZE( PXS_DATA_CATEGORY_PROPERTIES ); i++ )
    {
        if ( categoryID == PXS_DATA_CATEGORY_PROPERTIES[ i ].categoryID )
        {
            found = true;
            PXSGetResourceString( PXS_DATA_CATEGORY_PROPERTIES[ i ].nameStringID, pCategoryName );
            *pCaptionID  = PXS_DATA_CATEGORY_PROPERTIES[ i ].captionID;
            if ( PXS_DATA_CATEGORY_PROPERTIES[ i ].isColumnar )
            {
                *pIsColumnar = true;
            }
            if ( PXS_DATA_CATEGORY_PROPERTIES[ i ].isNode )
            {
                *pIsNode = true;
            }
            *pIndent = PXS_DATA_CATEGORY_PROPERTIES[ i ].indent;
            break;
        }
    }

    if ( found == false )
    {
       ErrorMessage = Format.StringUInt32( L"categoryID = %%1", categoryID);
       throw SystemException( ERROR_NOT_FOUND, ErrorMessage.c_str(), __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//     Set the captions
//
//  Parameters:
//      pRichText - receives the RTF document header
//
//  Remarks:
//      Includes RTF header as well as left margin of 500 twips and default
//      font of 10 points
//
//  Returns:
//      void
//===============================================================================================//
void PXSGetRichTextDocumentStart( String* pRichText )
{
    String ApplicationName;

    if ( pRichText == nullptr )
    {
        throw ParameterException( L"pRichText", __FUNCTION__ );
    }
    PXSGetApplicationName( &ApplicationName );

    *pRichText  = L"{\\rtf1\\ansi\\ansicpg1252\\deff0\\deflang2057\r\n"
                  L"{\\fonttbl{\\f0\\fswiss\\fprq2\\fcharset0 "
                  L"Verdana;}}\r\n"
                  L"{\\colortbl;\\red194\\green212\\blue251;\\red241"
                  L"\\green241\\blue241;\\red192\\green192\\blue192;"
                  L"\\red255\\green255\\blue255;}\r\n"
                  L"{\\*\\generator ";
    *pRichText += ApplicationName;
    *pRichText += L";}\r\n\r\n"
                  L"\\fs18\\margl500\r\n";
}

//===============================================================================================//
//  Description:
//      Get the file path for the WinAudit database GUID file
//
//  Parameters:
//      pFilePath - receives the file path
//
//  Remarks
//      Does not check for the existence of the file
//
//  Returns:
//      void
//===============================================================================================//
void PXSGetWinAuditGuidFilePath( String* pFilePath )
{
    Directory DirectoryObj;

    if ( pFilePath == nullptr )
    {
        throw ParameterException( L"pFilePath", __FUNCTION__ );
    }
    *pFilePath = PXS_STRING_EMPTY;

    // Make the file path
    DirectoryObj.GetSpecialDirectory( CSIDL_COMMON_APPDATA, pFilePath );
    if ( pFilePath->EndsWithCharacter( PXS_PATH_SEPARATOR ) == false )
    {
        *pFilePath += PXS_PATH_SEPARATOR;
    }
    *pFilePath += L"Parmavex\\WinAudit\\";
    *pFilePath += PXS_WINAUDIT_GUID_TXT;
}

//===============================================================================================//
//  Description:
//      Make the output path when in command line mode
//
//  Parameters:
//      FileName             - /f= switch value
//      TimestampSwitchValue - /T= switch value
//      LocalTime            - in ISO format
//      OutputPath           - receives the output path or connection string
//
//  Returns:
//      void
//===============================================================================================//
void PXSMakeCommandLineOutputPath( const String& FileSwitchValue,
                                   const String& TimestampSwitchValue,
                                   const String& LocalTimeIso, String* pOutputPath )
{
    File      FileObject;
    String    FileName, MacAddress, FilenameTimeStamp;
    String    Drive, Dir, Fname, Ext;
    Directory DirObject;
    TcpIpInformation   TcpIpInfo;
    SystemInformation SystemInfo;

    if ( pOutputPath == nullptr )
    {
        throw ParameterException( L"pOutputPath", __FUNCTION__ );
    }
    *pOutputPath = PXS_STRING_EMPTY;

    // Tidy up
    FileName   = FileSwitchValue;
    FileName.ReplaceI( L"\"", PXS_STRING_EMPTY );
    FileName.Trim();

    // For connection strings use the input as is
    if ( FileName.StartsWith( L"DBQ="   , false ) ||
         FileName.StartsWith( L"DRIVER=", false )  )
    {
        *pOutputPath = FileName;
        return;
    }

    // Use the computername.html if nothing was specified, assuming its a valid
    // file name
    if ( FileName.IsEmpty() )
    {
        SystemInfo.GetComputerNetBiosName( &FileName );
        if ( FileObject.IsValidFileName( FileName ) == false )
        {
            PXSLogAppWarn1( L"Computer name '%%1' is not a valid file name.", FileName );
            FileName = L"Audit";      // Fall back
        }
        FileName += L".html";
    }

    // Replace MAC Address keyword
    TcpIpInfo.GetMacAddress( &MacAddress );
    MacAddress.ReplaceChar( ':', PXS_STRING_EMPTY );
    FileName.ReplaceI( PXS_KEYWORD_MAC_ADDRESS, MacAddress.c_str() );

    // Insert the timestamp into the file name
    if ( TimestampSwitchValue.CompareI( L"date" ) == 0 )
    {
        LocalTimeIso.SubString( 0, 10, &FilenameTimeStamp );  // YYYY-MM-DD
        FilenameTimeStamp.ReplaceI( L"-", PXS_STRING_EMPTY );
    }
    else if ( TimestampSwitchValue.CompareI( L"time" ) == 0 )
    {
        LocalTimeIso.SubString( 11, 8, &FilenameTimeStamp );  // hh:mm:ss
        FilenameTimeStamp.ReplaceI( L":", PXS_STRING_EMPTY );
    }
    else if ( TimestampSwitchValue.CompareI( L"datetime" ) == 0 )
    {
        FilenameTimeStamp = LocalTimeIso;
        FilenameTimeStamp.ReplaceI( L"-", PXS_STRING_EMPTY );
        FilenameTimeStamp.ReplaceI( L":", PXS_STRING_EMPTY );
        FilenameTimeStamp.ReplaceI( L" ", L"-" );
    }

    if ( FilenameTimeStamp.GetLength() )
    {
        DirObject.SplitPath( FileName, &Drive, &Dir, &Fname, &Ext );
        FileName  = Dir;
        FileName += Fname;
        FileName += L"-";
        FileName += FilenameTimeStamp;
        FileName += Ext;
    }

    // Want a full path
    if ( FileName.IndexOf( PXS_PATH_SEPARATOR, 0 ) == PXS_MINUS_ONE )
    {
        PXSGetExeDirectory( pOutputPath );
    }
    *pOutputPath += FileName;
}

//===============================================================================================//
//  Description:
//      Read the GUID in the file at CSIDL_COMMON_APPDATA\Parmavex\WinAudit\
//
//  Parameters:
//      pWinAuditGuid - receives the GUID in string form
//
//  Returns:
//      void
//===============================================================================================//
void PXSReadWinAuditGuidFile( String* pWinAuditGuid )
{
    File        FileObject;
    size_t      i = 0, numLines;
    String      FilePath, Line, Name, Value;
    Formatter   Format;
    StringArray Lines, Tokens;

    if ( pWinAuditGuid == nullptr )
    {
        throw ParameterException( L"pWinAuditGuid", __FUNCTION__ );
    }
    *pWinAuditGuid = PXS_STRING_EMPTY;

    // Read the GUID
    PXSGetWinAuditGuidFilePath( &FilePath );
    FileObject.OpenText( FilePath );
    FileObject.ReadLineArray( FilePath, 1000, &Lines );
    FileObject.Close();
    numLines = Lines.GetSize();
    for ( i = 0; i < numLines; i++ )
    {
        Line = Lines.Get( i );
        Line.Trim();
        Line.ToArray( '=', &Tokens );
        if ( ( Line.StartsWith( L"!", false ) == false ) &&
             ( Tokens.GetSize() == 2 ) )
        {
            Name  = Tokens.Get( 0 );
            Value = Tokens.Get( 1 );
            Name.Trim();
            Value.Trim();
            if ( Name.CompareI( PXS_WINAUDIT_COMPUTER_GUID ) == 0 )
            {
                *pWinAuditGuid = Value;
            }
        }
    }

    if ( Format.IsValidStringGuid( *pWinAuditGuid ) == false )
    {
        throw SystemException( ERROR_INVALID_DATA, PXS_WINAUDIT_COMPUTER_GUID, __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Callback for qsort for a case insensitive ascending sort of an
//      PXS_TYPE_NUMBER_STRING array
//
//  Parameters:
//      pArg1 - pointer to string 1
//      pArg2 - pointer to string 2
//
//  Returns:
//      > 0  if arg1 greater than arg2
//      = 0  if arg1 equivalent to arg2
//      < 0  if arg1 less than arg2
//===============================================================================================//
int PXSQSortNumberStringAscending( const void* pArg1, const void* pArg2 )
{
    LPCWSTR pszOne  = nullptr;
    LPCWSTR pszTwo  = nullptr;

    if ( pArg1 )
    {
        pszOne = ( (const PXS_TYPE_NUMBER_STRING*)pArg1 )->szValue;
    }

    if ( pArg2 )
    {
        pszTwo = ( (const PXS_TYPE_NUMBER_STRING*)pArg2 )->szValue;
    }

    return PXSCompareString( pszOne, pszTwo, false );
}

//===============================================================================================//
//  Description:
//      Save the audit to a database or a file when in command line mode
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void PXSSaveAuditCommandLine( const String& OutputPath, TArray< AuditRecord >& AuditRecords )
{
    const  DWORD  MAX_DATABASE_TRIES = 5;
    bool   success  = false;
    char*  pszAnsi  = nullptr;
    DWORD  tries    = 0, random = 0, tableCounter = 0;
    size_t numChars = 0;
    File   FileObject;
    String ResultMessage, DataString;
    Formatter     Format;
    AuditData     Auditor;
    AuditRecord   AuditMasterRecord, ComputerMasterRecord;
    AllocateChars AnsiChars;
    AuditDatabase Database;
    OdbcExportDialog OdbcExport;
    TArray< TreeViewItem > CategoryItems;

    if ( OutputPath.StartsWith( L"DBQ="   , false ) ||
         OutputPath.StartsWith( L"DRIVER=", false )  )
    {
        // Send to database, will connect with default timeouts
        Auditor.MakeAuditMasterRecord( &AuditMasterRecord );
        Auditor.MakeComputerMasterRecord( &ComputerMasterRecord );
        OdbcExport.SetAuditRecords( AuditMasterRecord,
                                    ComputerMasterRecord, AuditRecords );
        Database.Connect( OutputPath,
                          PXS_DB_CONNECT_TIMEOUT_SECS_DEF,
                          PXS_DB_QUERY_TIMEOUT_SECS_DEF, nullptr );
        // Will try to send the data a few times in case of heavy
        // database load
        srand( GetTickCount() );
        while ( ( success == false ) && ( tries < MAX_DATABASE_TRIES ) )
        {
            try
            {
                // Wait for a 5s
                if ( tries )
                {
                    PXSLogAppInfo( L"Waiting for another try." );
                    random = PXSCastInt32ToUInt32( ( rand() % 1000 ) );
                    Sleep( 4500 + random );
                }
                OdbcExport.ExportRecordsToDatabase( &Database, &ResultMessage );
                success = true;
            }
            catch ( const Exception& eDB )
            {
                PXSLogException( eDB, __FUNCTION__ );
            }
            tries++;
        }
    }
    else
    {
        PXSLogAppInfo1( L"Output file path: '%%1'", OutputPath );
        if ( OutputPath.EndsWithStringI( L".csv" ) )
        {
            PXSAuditRecordsToCsv( AuditRecords, false, &DataString );
        }
        else if ( OutputPath.EndsWithStringI( L".csv2" ) )
        {
            PXSAuditRecordsToCsv2( AuditRecords, true, &DataString );
        }
        else if ( OutputPath.EndsWithStringI( L".rtf" ) )
        {
            PXSAuditRecordsToContent( AuditRecords, &CategoryItems, &tableCounter, &DataString );
        }
        else
        {
            // HTML is the default
            PXSAuditRecordsToHtml( AuditRecords, &DataString );
        }

        // Write out
        numChars = DataString.GetAnsiMultiByteLength();
        pszAnsi  = AnsiChars.New( numChars );
        Format.StringToAnsi( DataString, pszAnsi, numChars );
        FileObject.CreateNew( OutputPath, 0, false );
        FileObject.Write( pszAnsi, numChars );
    }
}

//===============================================================================================//
//  Description:
//     Sort audit records in ascending on the specified item
//
//  Parameters:
//      Records - the records to sort
//      itemID  - the item ID
//
//  Remarks:
//      Will sort on the first 32 characters of the item's string value as
//      strings are usually different before getting to this length
//
//  Returns:
//      void
//===============================================================================================//
void PXSSortAuditRecords( TArray< AuditRecord >* pRecords, DWORD itemID )
{
    size_t i = 0, numRecords = 0, numBytes = 0, lenChars = 0;
    String Value;
    AuditRecord   Record;
    AllocateBytes AllocBytes;
    TArray< AuditRecord > SortedArray;
    PXS_TYPE_NUMBER_STRING* pArray = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }

    numRecords = pRecords->GetSize();
    if ( numRecords < 1 )
    {
        return;     // Nothing to do
    }

    // Fill array
    numBytes = PXSMultiplySizeT( sizeof ( PXS_TYPE_NUMBER_STRING ), numRecords);
    pArray = reinterpret_cast<PXS_TYPE_NUMBER_STRING* >( AllocBytes.New( numBytes ) );
    for ( i = 0; i < numRecords; i++ )
    {
        // Get the first field of the record
        Value  = PXS_STRING_EMPTY;
        Record = pRecords->Get( i );
        Record.GetItemValue( itemID, &Value );

        pArray[ i ].number = i;
        lenChars = Value.GetLength();
        if ( lenChars )
        {
            // Copy as much as can fit
            StringCchCopyN( pArray[ i ].szValue,
                            ARRAYSIZE( pArray[ i ].szValue ), Value.c_str(), lenChars );
        }
    }

    // Sort
    qsort( pArray,
           numRecords,
           sizeof ( PXS_TYPE_NUMBER_STRING ), PXSQSortNumberStringAscending );

    // Replace
    SortedArray.SetSize( numRecords );
    for ( i = 0; i < numRecords; i++ )
    {
        Record = pRecords->Get( pArray[ i ].number );
        SortedArray.Set( i, Record );
    }
    *pRecords = SortedArray;
}

//===============================================================================================//
//  Description:
//     Create the write a GUID to CSIDL_COMMON_APPDATA\Parmavex\WinAudit\
//
//  Parameters:
//      none
//
//  Remarks:
//      Used in the Computer_Master.WinAudit_GUID table.
//      Does not overwrite any existing file
//
//  Returns:
//      void
//===============================================================================================//
void PXSWriteWinAuditGuidFile()
{
    File      FileObject;
    String    FilePath, FileContents, WinAuditGuid;
    Registry  RegObject;
    Formatter Format;
    Directory DirectoryObj;

    // Make the directory path and create if it does not exist.
    DirectoryObj.GetSpecialDirectory( CSIDL_COMMON_APPDATA, &FilePath );
    if ( FilePath.EndsWithCharacter( PXS_PATH_SEPARATOR ) == false )
    {
        FilePath += PXS_PATH_SEPARATOR;
    }

    FilePath += L"Parmavex";
    if ( DirectoryObj.Exists( FilePath ) == false )
    {
        DirectoryObj.CreateNew( FilePath );
    }

    FilePath += L"\\WinAudit";
    if ( DirectoryObj.Exists( FilePath ) == false )
    {
        DirectoryObj.CreateNew( FilePath );
    }

    FilePath += PXS_PATH_SEPARATOR;
    FilePath += PXS_WINAUDIT_GUID_TXT;
    if ( FileObject.Exists( FilePath ) )
    {
        throw SystemException( ERROR_INVALID_FUNCTION, L"PXS_WINAUDIT_GUID_TXT", __FUNCTION__ );
    }

    // For legacy reasons look in the registry for a GUID, if not present
    // will create one
    RegObject.Connect( HKEY_LOCAL_MACHINE );
    RegObject.GetStringValue( PXS_REG_PATH_COMPUTER_GUID,
                              PXS_WINAUDIT_COMPUTER_GUID, &WinAuditGuid );
    WinAuditGuid.Trim();
    if ( WinAuditGuid.GetLength() )
    {
        if ( Format.IsValidStringGuid( WinAuditGuid ) == false )
        {
            throw SystemException( ERROR_INVALID_DATA, PXS_WINAUDIT_COMPUTER_GUID, __FUNCTION__ );
        }
    }
    else
    {
        WinAuditGuid = Format.CreateGuid();
    }
    FileContents  = L"!WINAUDIT GUID FILE - DO NOT DELETE OR EDIT\r\n";
    FileContents += L"!This file is used to send data to a database.\r\n";
    FileContents += L"!The GUID below should match the value in\r\n";
    FileContents += L"!Computer_Master.WinAudit_GUID for this computer.\r\n";
    FileContents += PXS_WINAUDIT_COMPUTER_GUID;
    FileContents += L"=";
    FileContents += WinAuditGuid;
    FileContents += L"\r\n";

    FileObject.CreateNew( FilePath, 0, false );
    FileObject.WriteChars( FileContents );
    FileObject.Close();
    PXSLogAppInfo2( L"Wrote WinAuditGUID='%%1' in %%2.", WinAuditGuid, FilePath );
}
