///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Printers Information Class Implementation
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
#include "WinAudit/Header Files/PrinterInformation.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/AllocateWChars.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/FileVersion.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/TArray.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
PrinterInformation::PrinterInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
PrinterInformation::~PrinterInformation()
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
//      Get printer data as an array of audit records
//
//  Parameters:
//      pRecords - array to receive the printer data
//
//  Remarks:
//      Query local printers only, with flag PRINTER_ENUM_LOCAL at level 2.
//
//  Returns:
//      void
//===============================================================================================//
void PrinterInformation::GetAuditRecords( TArray< AuditRecord >* pRecords )
{
    DWORD    i = 0, cbNeeded = 0, Returned = 0, lastError = 0, cbBuf = 0;
    String   Value, DriverFile;
    String   DriverManufacturer, DriverVersion, DriverDescription;
    String   DriverDataFile, DriverConfigFile, Insert1;
    Formatter       Format;
    AuditRecord     Record;
    AllocateBytes   AllocBytes;
    PRINTER_INFO_2* pPrinterEnum = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    ////////////////////////////////////////////////////////////////////////////
    // Local printers

    // First call to determine the amount of memory required
    if ( EnumPrinters( PRINTER_ENUM_LOCAL,
                       nullptr,
                       2,
                       nullptr,       // NULL buffer
                       0, &cbNeeded, &Returned ) )
    {
        // Success with a NULL buffer, log and return
        Insert1.SetAnsi( __FUNCTION__ );
        PXSLogAppInfo1(L"EnumPrinters said no data in '%%1'", Insert1);
        return;
    }

    // Test for error other than insufficient buffer
    lastError = GetLastError();
    if ( lastError != ERROR_INSUFFICIENT_BUFFER )
    {
        throw SystemException( lastError, L"EnumPrinters failed on first call.", __FUNCTION__ );
    }


    // Second call, allocate extra bytes
    cbNeeded     = PXSMultiplyUInt32( cbNeeded, 2 );
    pPrinterEnum = reinterpret_cast<PRINTER_INFO_2*>( AllocBytes.New(cbNeeded));
    cbBuf        = cbNeeded;
    cbNeeded     = 0;
    Returned     = 0;
    if ( EnumPrinters( PRINTER_ENUM_LOCAL,
                       nullptr,
                       2, (LPBYTE)pPrinterEnum, cbBuf, &cbNeeded, &Returned ) == 0 )
    {
        throw SystemException( GetLastError(),
                               L"EnumPrinters failed on second call.", __FUNCTION__ );
    }

    for ( i = 0; i < Returned; i++ )
    {
        // Catch and proceed to the next printer
        try
        {
            // Make the record
            Record.Reset( PXS_CATEGORY_PRINTERS );

            // Printer counting starts at 1
            Record.Add( PXS_PRINTERS_ITEM_NUMBER, Format.UInt32( i + 1 ) );
            Record.Add( PXS_PRINTERS_NAME       , pPrinterEnum[i].pPrinterName);
            Record.Add( PXS_PRINTERS_SHARE_NAME , pPrinterEnum[i].pShareName  );
            Record.Add( PXS_PRINTERS_PORT_NAME  , pPrinterEnum[i].pPortName   );
            Record.Add( PXS_PRINTERS_LOCATION   , pPrinterEnum[i].pLocation   );

            Value = PXS_STRING_EMPTY;
            if ( pPrinterEnum[ i ].AveragePPM > 0 )
            {
                Value = Format.UInt32( pPrinterEnum[ i ].AveragePPM );
            }
            Record.Add( PXS_PRINTERS_AVERAGE_PPM, Value );

            Value = PXS_STRING_EMPTY;
            TranslateAttributes( pPrinterEnum[ i ].Attributes, &Value );
            Record.Add( PXS_PRINTERS_ATTRIBUTES, Value );

            Value = PXS_STRING_EMPTY;
            TranslateStatus( pPrinterEnum[ i ].Status, &Value );
            Record.Add( PXS_PRINTERS_PRINTER_STATUS, Value );

            if ( pPrinterEnum[ i ].pDevMode )
            {
                Value = PXS_STRING_EMPTY;
                TranslatePaperSize( pPrinterEnum[ i ].pDevMode->dmPaperSize,
                                    &Value );
                Record.Add( PXS_PRINTERS_PAPER_SIZE, Value );

                Value = PXS_STRING_EMPTY;
                TranslateOrientation( pPrinterEnum[ i ].pDevMode->dmOrientation,
                                      &Value);
                Record.Add( PXS_PRINTERS_ORIENTATION, Value );

                Value = PXS_STRING_EMPTY;
                TranslatePrintQuality( pPrinterEnum[i].pDevMode->dmPrintQuality,
                                       &Value);
                Record.Add( PXS_PRINTERS_PRINT_QUALITY, Value );
            }

            ////////////////////////////////////////////////////////////////////
            // Add the printer driver

            DriverFile         = PXS_STRING_EMPTY;
            DriverManufacturer = PXS_STRING_EMPTY;
            DriverVersion      = PXS_STRING_EMPTY;
            DriverDescription  = PXS_STRING_EMPTY;
            DriverDataFile     = PXS_STRING_EMPTY;
            DriverConfigFile   = PXS_STRING_EMPTY;
            GetPrinterDriverDetails( pPrinterEnum[ i ].pPrinterName,
                                     &DriverFile,
                                     &DriverManufacturer,
                                     &DriverVersion,
                                     &DriverDescription,
                                     &DriverDataFile, &DriverConfigFile );

            Record.Add( PXS_PRINTERS_DRIVER_NAME, pPrinterEnum[i].pDriverName );
            Record.Add( PXS_PRINTERS_DRIVER_FILE        , DriverFile );
            Record.Add( PXS_PRINTERS_DRIVER_VERSION     , DriverVersion );
            Record.Add( PXS_PRINTERS_DRIVER_MANUFACTURER, DriverManufacturer);
            Record.Add( PXS_PRINTERS_DRIVER_DESCRIPTION , DriverDescription);
            Record.Add( PXS_PRINTERS_DRIVER_DATA_FILE   , DriverDataFile);
            Record.Add( PXS_PRINTERS_DRIVER_CONFIG_FILE , DriverConfigFile);

            pRecords->Add( Record );
        }
        catch ( const Exception& ePrinter )
        {
            // Log and continue
            PXSLogException( ePrinter, __FUNCTION__ );
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the names of the installed printers
//
//  Parameters:
//      pPrinterNames - string array to receive the names
//
//  Remarks:
//      Query local printers only, with flag PRINTER_ENUM_LOCAL at  level 1.
//
//  Returns:
//      void
//===============================================================================================//
void PrinterInformation::GetPrinterNames( StringArray* pPrinterNames )
{
    size_t  i = 0, numRecords = 0;
    String  Name;
    AuditRecord Record;
    TArray< AuditRecord > Records;

    if ( pPrinterNames == nullptr )
    {
        throw ParameterException( L"pPrinterNames", __FUNCTION__ );
    }
    pPrinterNames->RemoveAll();

    GetAuditRecords( &Records );
    numRecords = Records.GetSize();
    for ( i = 0; i < numRecords; i++ )
    {
        Record = Records.Get( i );
        Name   = PXS_STRING_EMPTY;
        Record.GetItemValue( PXS_PRINTERS_NAME, &Name );
        Name.Trim();
        if ( Name.GetLength() )
        {
            pPrinterNames->Add( Name );
        }
    }
    pPrinterNames->Sort( true );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the details about a printer's driver
//
//  Parameters:
//      pszPrinterName     - the printer's name
//      pDriverPath        - receives the driver file path
//      pDriverManufacturer- receives the driver's manufacturer name
//      pDriverVersion     - receives the driver's version
//      pDriverDescription - receives the driver's description
//      pDriverDataFile    - receives the driver's data file path
//      pDriverConfigFile  - receives the driver's configuration path
//
//  Remarks:
//      Will log errors rather than throw as this may be optional data
//
//  Returns:
//      void
//===============================================================================================//
void PrinterInformation::GetPrinterDriverDetails( LPCWSTR pszPrinterName,
                                                  String* pDriverPath,
                                                  String* pDriverManufacturer,
                                                  String* pDriverVersion,
                                                  String* pDriverDescription,
                                                  String* pDriverDataFile,
                                                  String* pDriverConfigFile )
{
    File     FileObject;
    size_t   numChars = 0;
    DWORD    cbNeeded = 0, lastError = 0, cbBuf = 0;
    HANDLE   hPrinter = nullptr;
    wchar_t* pszTemp  = nullptr;
    String   Insert2, PrinterName;
    FileVersion    FileVer;
    AllocateBytes  AllocBytes;
    AllocateWChars AllocWChars;
    DRIVER_INFO_2* pDriverInfo = nullptr;

    if ( ( pDriverPath         == nullptr ) ||
         ( pDriverManufacturer == nullptr ) ||
         ( pDriverVersion      == nullptr ) ||
         ( pDriverDescription  == nullptr ) ||
         ( pDriverDataFile     == nullptr ) ||
         ( pDriverConfigFile   == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    *pDriverPath         = PXS_STRING_EMPTY;
    *pDriverManufacturer = PXS_STRING_EMPTY;
    *pDriverVersion      = PXS_STRING_EMPTY;
    *pDriverDescription  = PXS_STRING_EMPTY;
    *pDriverDataFile     = PXS_STRING_EMPTY;
    *pDriverConfigFile   = PXS_STRING_EMPTY;

    PrinterName = pszPrinterName;
    PrinterName.Trim();
    if ( PrinterName.IsEmpty() )
    {
        return;     // Nothing to do
    }
    numChars = PrinterName.GetLength();
    numChars = PXSAddSizeT( numChars, 1 );          // NULL terminator
    pszTemp  = AllocWChars.New( numChars );
    StringCchCopy( pszTemp, numChars, PrinterName.c_str() );
    if ( OpenPrinter( pszTemp, &hPrinter, nullptr ) == 0 )
    {
        throw SystemException( GetLastError(), PrinterName.c_str(), "OpenPrinter" );
    }

    // First call at level 2 to get the required buffer size
    if ( GetPrinterDriver( hPrinter,
                           nullptr,
                           2,
                           nullptr,     // i.e. NULL buffer
                           0,           // i.e. zero size buffer
                           &cbNeeded ) == 0 )
    {
        // Call succeeded despite a NULL buffer, log and return
        ClosePrinter( hPrinter );
        Insert2.SetAnsi( __FUNCTION__ );
        PXSLogAppInfo2( L"GetPrinterDriver said no data for '%%1' in '%%2'.",
                        PrinterName, Insert2 );
        return;
    }

    // Test for error, other than no buffer
    lastError = GetLastError();
    if ( lastError != ERROR_INSUFFICIENT_BUFFER )
    {
        ClosePrinter( hPrinter );
        throw SystemException( lastError, PrinterName.c_str(), "GetPrinterDriver" );
    }

    // Second call, allocate extra bytes
    cbNeeded    = PXSMultiplyUInt32( cbNeeded, 2 );
    cbBuf       = cbNeeded;
    pDriverInfo = reinterpret_cast<DRIVER_INFO_2*>(AllocBytes.New(cbNeeded));
    cbNeeded    = 0;
    if ( GetPrinterDriver( hPrinter, nullptr, 2, (LPBYTE)&pDriverInfo, cbBuf, &cbNeeded ) == 0 )
    {
        ClosePrinter( hPrinter );
        throw SystemException( GetLastError(),
                               PrinterName.c_str(), "GetPrinterDriver" );
    }

    // Must catch exceptions to clean up
    try
    {
        *pDriverPath       = pDriverInfo->pDriverPath;
        *pDriverDataFile   = pDriverInfo->pDataFile;
        *pDriverConfigFile = pDriverInfo->pConfigFile;

        // Get version info, should be a dll
        try
        {
            if ( FileObject.Exists( *pDriverPath ) &&
                 pDriverPath->EndsWithStringI( L".dll" ) )
            {
                FileVer.GetVersion( *pDriverPath,
                                    pDriverManufacturer,
                                    pDriverVersion, pDriverDescription );
            }
        }
        catch ( const Exception& eVersion )
        {
            PXSLogException( L"Error getting file version.", eVersion, __FUNCTION__ );
        }
    }
    catch ( const Exception& )
    {
        ClosePrinter( hPrinter );
        throw;
    }
    ClosePrinter( hPrinter );
}

//===============================================================================================//
//  Description:
//      Translate a printer's attributes into a string
//
//  Parameters:
//      attributes - defined constant of orientation
//      pValue     - receives the attributes
//
//  Remarks:
//      See PRINTER_INFO_2 structure or similar
//      PRINTER_ATTRIBUTE_FAX is in winspool.h 0x00004000
//
//  Returns:
//      void
//===============================================================================================//
void PrinterInformation::TranslateAttributes( DWORD attributes, String* pValue )
{
    size_t    i = 0;
    Formatter Format;

    struct _ATTRIBUTE
    {
        DWORD   attribute;
        LPCWSTR pszAttribute;
    } Attributes[] =
        { { PRINTER_ATTRIBUTE_SHARED   , L"Shared"                     },
          { PRINTER_ATTRIBUTE_QUEUED   , L"Queued"                     },
          { PRINTER_ATTRIBUTE_NETWORK  , L"Network Printer Connection" },
          { PRINTER_ATTRIBUTE_LOCAL    , L"Local Printer"              },
          { PRINTER_ATTRIBUTE_RAW_ONLY , L"Raw Only"                   },
          { PRINTER_ATTRIBUTE_DIRECT   , L"Direct Printing"            },
          { PRINTER_ATTRIBUTE_DIRECT   , L"Spooled Printing"           },
          { PRINTER_ATTRIBUTE_PUBLISHED, L"Directory Service Published" },
          { 0x00004000                 , L"Fax Printer"                }
        };

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Attributes ); i++ )
    {
        if ( attributes & Attributes[ i ].attribute )
        {
            if ( pValue->GetLength() )
            {
                *pValue += L", ";
            }
            *pValue += Attributes[ i ].pszAttribute;
        }
    }

    // If none, use the numerical value
    if ( pValue->IsEmpty() )
    {
        *pValue = Format.StringUInt32( L"Code %%1", attributes );
        PXSLogAppWarn1( L"Unrecognised printer attributes, %%1.", *pValue );
    }
}

//===============================================================================================//
//  Description:
//      Translate a paper orientation constant into a string
//
//  Parameters:
//      orientation - defined constant of orientation
//      pValue      - string object to receive type
//
//  Remarks:
//
//  Returns:
//      void
//===============================================================================================//
void PrinterInformation::TranslateOrientation(SHORT orientation, String* pValue)
{
    if ( orientation == DMORIENT_PORTRAIT )
    {
        *pValue = L"Portrait";
    }
    else if ( orientation == DMORIENT_LANDSCAPE )
    {
        *pValue = L"Landscape";
    }
    else
    {
        *pValue = PXS_STRING_EMPTY;  // Error or unknown
    }
}

//===============================================================================================//
//  Description:
//      Translate a paper size constant into a string
//
//  Parameters:
//      paperSize - defined constant of paper size
//      pValue    - receive the paper size
//
//  Remarks:
//      See "Paper Sizes" of International Features
//  Returns:
//      void
//===============================================================================================//
void PrinterInformation::TranslatePaperSize( SHORT paperSize, String* pValue )
{
    size_t    i = 0;
    Formatter Format;

    struct _SIZE
    {
        SHORT   paperSize;
        LPCWSTR pszPaperSize;
    } Sizes[] = {
    {DMPAPER_LETTER            , L"Letter, 8 1/2- by 11-inches"            },
    {DMPAPER_LEGAL             , L"Legal, 8 1/2- by 14-inches"             },
    {DMPAPER_10X14             , L"10- by 14-inch sheet"                   },
    {DMPAPER_11X17             , L"11- by 17-inch sheet"                   },
    {DMPAPER_A3                , L"A3 sheet, 297- by 420-millimeters"      },
    {DMPAPER_A4                , L"A4 sheet, 210- by 297-millimeters"      },
    {DMPAPER_A4SMALL           , L"A4 small sheet, 210- by 297-millimeters"},
    {DMPAPER_A5                , L"A5 sheet, 148- by 210-millimeters"      },
    {DMPAPER_CSHEET            , L"C Sheet, 17- by 22-inches"              },
    {DMPAPER_DSHEET            , L"D Sheet, 22- by 34-inches"              },
    {DMPAPER_ENV_9             , L"#9 Envelope, 3 7/8- by 8 7/8-inches"    },
    {DMPAPER_ENV_10            , L"#10 Envelope, 4 1/8- by 9 1/2-inches"   },
    {DMPAPER_ENV_11            , L"#11 Envelope, 4 1/2- by 10 3/8-inches"  },
    {DMPAPER_ENV_12            , L"#12 Envelope, 4 3/4- by 11-inches"      },
    {DMPAPER_ENV_14            , L"#14 Envelope, 5- by 11 1/2-inches"      },
    {DMPAPER_ENV_C5            , L"C5 Envelope, 162- by 229-millimeters"   },
    {DMPAPER_ENV_C3            , L"C3 Envelope, 324- by 458-millimeters"   },
    {DMPAPER_ENV_C4            , L"C4 Envelope, 229- by 324-millimeters"   },
    {DMPAPER_ENV_C6            , L"C6 Envelope, 114- by 162-millimeters"   },
    {DMPAPER_ENV_C65           , L"C65 Envelope, 114- by 229-millimeters"  },
    {DMPAPER_ENV_B4            , L"B4 Envelope, 250- by 353-millimeters"   },
    {DMPAPER_ENV_B5            , L"B5 Envelope, 176- by 250-millimeters"   },
    {DMPAPER_ENV_B6            , L"B6 Envelope, 176- by 125-millimeters"   },
    {DMPAPER_ENV_DL            , L"DL Envelope, 110- by 220-millimeters"   },
    {DMPAPER_ENV_ITALY         , L"Italy Envelope, 110- by 230-millimeters"},
    {DMPAPER_ENV_MONARCH       , L"Monarch Envelope 3 7/8- by 7 1/2-inches"},
    {DMPAPER_ENV_PERSONAL      , L"6 3/4 Envelope, 3 5/8- by 6 1/2-inches" },
    {DMPAPER_ESHEET            , L"E Sheet, 34- by 44-inches"              },
    {DMPAPER_EXECUTIVE         , L"Executive, 7 1/4- by 10 1/2-inches"     },
    {DMPAPER_FANFOLD_US        , L"US Std Fanfold, 14 7/8- by 11-inches"   },
    {DMPAPER_FOLIO             , L"Folio, 8 1/2- by 13-inch paper"         },
    {DMPAPER_LEDGER            , L"Ledger, 17- by 11-inches"               },
    {DMPAPER_LETTERSMALL       , L"Letter Small, 8 1/2- by 11-inches"      },
    {DMPAPER_NOTE              , L"Note, 8 1/2- by 11-inches"              },
    {DMPAPER_QUARTO            , L"Quarto, 215- by 275-millimeter paper"   },
    {DMPAPER_STATEMENT         , L"Statement, 5 1/2- by 8 1/2-inches"      },
    {DMPAPER_TABLOID           , L"Tabloid, 11- by 17-inches"              },
    {DMPAPER_FANFOLD_STD_GERMAN, L"German Std Fanfold, 8 1/2- by 12-inches"},
    {DMPAPER_FANFOLD_LGL_GERMAN, L"German Legal Fanfold, 8 - by 13-inches" }
    };

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( Sizes ); i++ )
    {
        if ( paperSize == Sizes[ i ].paperSize )
        {
            *pValue = Sizes[ i ].pszPaperSize;
            break;
        }
    }

    if ( pValue->IsEmpty() )
    {
        *pValue = Format.StringInt16( L"Code %%1", paperSize );
        PXSLogAppWarn1( L"Unrecognised paper size, %%1.", *pValue );
    }
}

//===============================================================================================//
//  Description:
//      Translate a paper quality constant into a string
//
//  Parameters:
//      printQuality - defined constant of print size
//      pValue       - string object to receive type
//
//  Remarks:
//      If a positive value is specified, it specifies the number of dots
//      per inch (DPI)
//
//  Returns:
//      void
//===============================================================================================//
void PrinterInformation::TranslatePrintQuality( SHORT printQuality,
                                                String* pValue )
{
    Formatter Format;

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = PXS_STRING_EMPTY;

    switch( printQuality )
    {
        default:
            if ( printQuality > 0 )
            {
                *pValue = Format.StringInt16( L"%%1 Dots Per Inch", printQuality );
            }
            else
            {
                *pValue = Format.StringInt16( L"Code %%1", printQuality );
            }
            break;

    case DMRES_HIGH:
            *pValue = L"High";
            break;

        case DMRES_MEDIUM:
            *pValue = L"Medium";
            break;

        case DMRES_LOW:
            *pValue = L"Low";
            break;

        case DMRES_DRAFT:
            *pValue = L"Draft";
            break;
    }
}

//===============================================================================================//
//  Description:
//      Translate a printer's status into a string
//
//  Parameters:
//        status - defined constant of orientation
//        pValue - string object to receive type
//
//  Remarks:
//      See PRINTER_INFO_2 or similar
//
//  Returns:
//      void
//===============================================================================================//
void PrinterInformation::TranslateStatus( DWORD status, String* pValue )
{
    size_t    i = 0;
    Formatter Format;

    struct _STATUS
    {
      DWORD   status;
      LPCWSTR pszStatus;
    } PrinterStatus[] =
        { { PRINTER_STATUS_BUSY             , L"Busy"                      },
          { PRINTER_STATUS_DOOR_OPEN        , L"Door is open"              },
          { PRINTER_STATUS_ERROR            , L"Error state"               },
          { PRINTER_STATUS_INITIALIZING     , L"Initializing"              },
          { PRINTER_STATUS_IO_ACTIVE        , L"Active input/output state" },
          { PRINTER_STATUS_MANUAL_FEED      , L"Manual feed state"         },
          { PRINTER_STATUS_NO_TONER         , L"Out of toner"              },
          { PRINTER_STATUS_NOT_AVAILABLE    , L"Not available"             },
          { PRINTER_STATUS_OFFLINE          , L"Offline"                   },
          { PRINTER_STATUS_OUT_OF_MEMORY    , L"Out of memory"             },
          { PRINTER_STATUS_OUTPUT_BIN_FULL  , L"Output bin is full"        },
          { PRINTER_STATUS_PAGE_PUNT        , L"Cannot print current page" },
          { PRINTER_STATUS_PAPER_JAM        , L"Paper jam"                 },
          { PRINTER_STATUS_PAPER_OUT        , L"Out of paper"              },
          { PRINTER_STATUS_PAPER_PROBLEM    , L"Paper problem"             },
          { PRINTER_STATUS_PAUSED           , L"Paused"                    },
          { PRINTER_STATUS_PENDING_DELETION , L"Being deleted"             },
          { PRINTER_STATUS_POWER_SAVE       , L"In power save mode"        },
          { PRINTER_STATUS_PRINTING         , L"Printing"                  },
          { PRINTER_STATUS_PROCESSING       , L"Processing a print job"    },
          { PRINTER_STATUS_SERVER_UNKNOWN   , L"Unknown"                   },
          { PRINTER_STATUS_TONER_LOW        , L"Low on toner"              },
          { PRINTER_STATUS_USER_INTERVENTION, L"Error requiring user "
                                              L"intervention"              },
          { PRINTER_STATUS_WAITING          , L"Waiting,"                  },
          { PRINTER_STATUS_WARMING_UP       , L"Warming up"                }
        };

    if ( pValue == nullptr )
    {
        throw ParameterException( L"pValue", __FUNCTION__ );
    }
    *pValue = PXS_STRING_EMPTY;

    for ( i = 0; i < ARRAYSIZE( PrinterStatus ); i++ )
    {
        if ( status & PrinterStatus[ i ].status )
        {
            if ( pValue->GetLength() )
            {
                *pValue += L", ";
            }
            *pValue += PrinterStatus[ i ].pszStatus;
        }
    }

    // If none, use the numerical value
    if ( status && pValue->GetLength() )
    {
        *pValue = Format.StringUInt32( L"Code %%1", status );
        PXSLogAppWarn1( L"Unrecognised printer status, %%1.", *pValue );
    }
}
