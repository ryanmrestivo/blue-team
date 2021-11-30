///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Application Class Header
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

#ifndef PXSBASE_APPLICATION_H_
#define PXSBASE_APPLICATION_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/PxsBase.h"

// 2. C System Files
#include <signal.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Logger.h"
#include "PxsBase/Header Files/Mutex.h"
#include "PxsBase/Header Files/StringT.h"
#include "PxsBase/Header Files/TArray.h"
#include "PxsBase/Header Files/TList.h"

// 6. Forwards
class ByteArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class Application
{
    public:
        // Default constructor
        Application();

        // Destructor
        ~Application();

        // Methods
        void      AddAppResourceLanguageID( WORD resourceID,
                                            WORD primaryLanguage, WORD subLanguage );
        void      GetApplicationName( String* pApplicationName );
        void      GetApplicationTempPath( bool create, String* pApplicationTempPath );
        void      GetCopyrightNotice( String* pCopyrightNotice );
        void      GetDiagnosticData( bool extraInfo, StringArray* pDiagnosticData );
        void      GetDiagnostics( String* pDiagnostics );
        HINSTANCE GetExeInstanceHandle();
        void      GetGradientColours( COLORREF* pGradientOne, COLORREF* pGradientTwo );
        HWND      GetHwndMainFrame() const;
        void      GetLastAccessedDirectory( String* pLastAccessedDirectory );
        DWORD     GetLogLevel() const;
        void      GetLogMessages( bool purge, StringArray* pLogMessages );
        void      GetManufacturerAddress( String* pManufacturerAddress );
        void      GetManufacturerInfo( String* pCopyrightNotice,
                                       String* pManufacturerAddress,
                                       String* pManufacturerName,
                                       String* pRegisteredOrgName,
                                       String* pRegisteredOwnerName,
                                       String* pProductID,
                                       String* pSupportEmail, String* pWebSiteURL );
        void      GetManufacturerName( String* pManufacturerName );
        void      GetProductID( String* pProductID );
        void      GetRegisteredOrgName( String* pRegisteredOrgName );
        void      GetRegisteredOwnerName( String* pRegisteredOwnerName );
        void      GetResourceString( DWORD resouceID, String* pResourceString);
        void      GetResourceString1( DWORD resouceID,
                                      const String& Insert1,
                                      String* pResourceString );
        void      GetResourceString2( DWORD resouceID,
                                      const String& Insert1,
                                      const String& Insert2,
                                      String* pResourceString );
        void      GetResourceString3( DWORD resouceID,
                                      const String& Insert1,
                                      const String& Insert2,
                                      const String& Insert3, String* pResourceString );
        bool      GetStopBackgroundTasks() const;
        void      GetSupportEmail( String* pSupportEmail );
        void      GetWebSiteURL( String* pWebSiteURL );
        bool      IsOnRemoveableDrive();
        bool      IsLogging();
        void      LoadByteDataResource( DWORD resourceID, ByteArray* pBytes ) const;
        void      LoadStringFile( WORD primaryLanguage, WORD subLanguage );
        void      LoadTextDataResource( DWORD resourceID, String* pStringData ) const;
        void      LogFlush();
        void      RegisterNewHandler();
        void      SetApplicationName( LPCWSTR pszApplicationName );
        void      SetCopyrightNotice( LPCWSTR pszCopyrightNotice );
        void      SetGradientColours( COLORREF gradientOne, COLORREF gradientTwo );
        void      SetHwndMainFrame( HWND hWndMainFrame );
        void      SetLastAccessedDirectory( const String& LastAccessedDirectory );
        void      SetLogLevel( DWORD logLevel );
        void      SetManufacturerAddress( LPCWSTR pszManufacturerAddress );
        void      SetManufacturerName( LPCWSTR pszManufacturerName  );
        void      SetProductID( LPCWSTR pszProductID  );
        void      SetRegisteredOrgName( LPCWSTR pszRegisteredOrgName );
        void      SetRegisteredOwnerName( LPCWSTR pszRegisteredOwnerName  );
        void      SetStopBackgroundTasks( bool stopBackgroundTasks );
        void      SetSupportEmail( LPCWSTR pszSupportEmail  );
        void      SetWebSiteURL( LPCWSTR pszWebSiteURL  );
        void      StartLogger( const String& LogFilePath,
                               bool createNew, bool wantComputerName );
        void      StartLogger( bool wantComputerName, bool wantExtraInfo );
        void      StopLogger();
        void      WriteToAppLog( DWORD severity,
                                 DWORD errorType,
                                 DWORD errorCode,
                                 bool translateError,
                                 LPCWSTR pszMessage,
                                 const String& Insert1, const String& Insert2 );
    protected:
        // Methods

        // Data members

    private:
        struct ID_TEXT
        {
            DWORD   id;
            String  Text;
            ID_TEXT():id( 0 ),
                      Text()
            {
            };
        };

        struct RESOURCE_LANGIDS
        {
            WORD resourceID;
            WORD primaryLanguage;
            WORD subLanguage;
            RESOURCE_LANGIDS(): resourceID( 0 ),
                                primaryLanguage( 0 ),
                                subLanguage( 0 )
            {
            };
        };

        // Copy constructor - not allowed
        Application( const Application& oApplication );

        // Assignment operator - not allowed
        Application& operator= ( const Application& oApplication );

        // Methods
        void GetLanguageResourceIDs( WORD  primaryLanguage,
                                     WORD  subLanguage,
                                     WORD* pPxsBaseLocaleID,
                                     WORD* pAppEnglishID, WORD* pAppLocaleID );
        void GetStringTable( WORD resourceID, TList< ID_TEXT >* pStringTable );
        void MergeStringTables( TList< ID_TEXT >* pTable1, TList< ID_TEXT >* pTable2 );
        bool StringTableLineToIdString( const String& Line, ID_TEXT* pIdText );

        // Data members
        bool              m_bStopBackgroundTasks;
        sig_atomic_t      m_isAppLogging;
        HWND              m_hWndMainFrame;
        DWORD             m_uLogLevel;
        COLORREF          m_crGradientOne;
        COLORREF          m_crGradientTwo;
        Mutex             m_Mutex;
        Logger            m_Logger;
        String            m_ApplicationName;
        String            m_AppTempFolderName;
        String            m_CopyrightNotice;
        String            m_LastAccessedDirectory;
        String            m_LogDirectoryPath;
        String            m_ManufacturerAddress;
        String            m_ManufacturerName;
        String            m_ProductID;
        String            m_RegisteredOrgName;
        String            m_RegisteredOwnerName;
        String            m_SupportEmail;
        String            m_WebSiteURL;
        TList< ID_TEXT >  m_StringTable;
        TArray< RESOURCE_LANGIDS > m_AppResourceLanguageIDs;
};

#endif  // PXSBASE_APPLICATION_H_
