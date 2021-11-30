///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Exception Class Header
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

#ifndef PXSBASE_EXCEPTION_H_
#define PXSBASE_EXCEPTION_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/PxsBase.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/StringT.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class Exception
{
    public:
        // Default constructor
        Exception();

        // Constructor of an exception object with error information
        Exception( DWORD errorType,
                   DWORD errorCode,
                   LPCWSTR pszDetails, const char* pszFunction );

        // Copy constructor
        Exception( const Exception& oException );

        // Destructor
        ~Exception();

        // Assignment operator
        Exception& operator= ( const Exception& oException );

        // Methods
const String&  GetCrashString() const;
        DWORD  GetErrorCode() const;
        DWORD  GetErrorType() const;
const String&  GetMessage() const;
        void   Reset();
        bool   Translate( DWORD errorType, DWORD errorCode, String* pTranslation );
        void   TranslateType( DWORD errorType, String* pErrorName );

    protected:
        // Methods
        void   FillCrashString( const EXCEPTION_POINTERS* pException );
        void   FillInDescription();
        void   FillInDetails( LPCWSTR pszDetails, const char* pszFunction, DWORD maxDepth );
        void   FillInMessage();

        // Data members
        DWORD  m_uErrorCode;
        DWORD  m_uErrorType;
        String m_CrashString;
        String m_Description;
        String m_Details;
        String m_Message;

    private:
        // Methods
        bool  GetFunctionNameFromAddress( const void* pAddress,
                                          String* pFunctionName );
        void  MakeStackTrace( const CONTEXT* pContextRecord,
                              bool detailed, DWORD maxDepth, String* pStackTrace );
        void  TranslateAddressToModuleInformation( const void* pAddress,
                                                   bool detailed, String* pTranslation );
        bool  TranslateAppError( DWORD errorCode, String* pTranslation );
        bool  TranslateComError( HRESULT hResult, String* pTranslation );
        bool  TranslateCommonDlgError( DWORD errorCode, String* pTranslation );
        bool  TranslateCfgMgrError( CONFIGRET errorCode, String* pTranslation );
        bool  TranslateDnsError( DNS_STATUS status, String* pTranslation );
        bool  TranslateMAPIError( ULONG errorCode, String* pTranslation );
        bool  TranslateModuleError( DWORD  errorCode,
                                    LPCWSTR pszModuleRelPath, String* pTranslation );
        bool  TranslateSystemError( DWORD errorCode, String* pTranslation );
        bool  TranslateStructuredException( DWORD exceptionCode, String* pTranslation );
        bool  TranslateWBEMStatusError( HRESULT hResult, String* pTranslation );
        bool  TranslateZLibError( int errorCode, String*pTranslation );


        // Data members
        static const DWORD MAX_STACK_DEPTH = 63;  // See CaptureStackBackTrace
};

#endif  // PXSBASE_EXCEPTION_H_
