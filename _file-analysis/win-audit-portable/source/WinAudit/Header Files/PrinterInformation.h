///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Printers Information Class Header
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

#ifndef WINAUDIT_PRINTER_INFORMATION_H_
#define WINAUDIT_PRINTER_INFORMATION_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/WinAudit.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project

// 6. Forwards
class AuditRecord;
template< class T > class TArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class PrinterInformation
{
    public:
        // Default constructor
        PrinterInformation();

        // Destructor
        ~PrinterInformation();

        // Methods
        void    GetAuditRecords( TArray< AuditRecord >* pRecords );
        void    GetPrinterNames( StringArray* pNames );

    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        PrinterInformation( const PrinterInformation& oPrinterInformation );

        // Assignment operator - not allowed
        PrinterInformation& operator= ( const PrinterInformation& oPrinterInformation );

        // Methods
        void    GetPrinterDriverDetails( LPCWSTR pszPrinterName,
                                         String* pDriverPath,
                                         String* pDriverManufacturer,
                                         String* pDriverVersion,
                                         String* pDriverDescription,
                                         String* pDriverDataFile,
                                         String* pDriverConfigFile );
        void    TranslateAttributes( DWORD attributes, String* pValue );
        void    TranslateOrientation( SHORT orientation, String* pValue );
        void    TranslatePaperSize( SHORT paperSize, String* pValue );
        void    TranslatePrintQuality( SHORT printQuality, String* pValue );
        void    TranslateStatus( DWORD status, String* pValue );
};

#endif  // WINAUDIT_PRINTER_INFORMATION_H_
