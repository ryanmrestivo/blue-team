///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Find Files Class Header
//
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Copyright 1987-2016 PARMAVEX SERVICES
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

#ifndef WINAUDIT_FIND_FILES_H_
#define WINAUDIT_FIND_FILES_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/FindFiles.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Object.h"
#include "PxsBase/Header Files/TList.h"

// 5. This Project

// 6. Forwards
class AuditRecord;
template< class T > class TArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class FindFiles : public Object
{
    public:
        // Default constructor
        FindFiles();

        // Destructor
        ~FindFiles();

        // Methods
        void    FindInDirectory( const String& Directory,
                                 const StringArray& Extensions,
                                 const StringArray& DirectoryStopList );
        void    FindOnLocalHardDrives( const StringArray& Extensions,
                                       const StringArray& DirectoryStopList );
        void    GetFileRecords( const String& Extension, TArray< AuditRecord >* pRecords );
        void    GetNonWindowsExeRecords( TArray< AuditRecord >* pRecords );


    protected:
        // Methods

        // Data members

    private:
        typedef struct _TYPE_SIZE_DATE_PATH
        {
            UINT64   fileSize;                    // File size
            FILETIME lastWrite;                   // last write time to file
            wchar_t  szFilePath[ MAX_PATH + 1 ];  // Full path to file
        } TYPE_SIZE_DATE_PATH;

        // Copy constructor - not allowed
        FindFiles( const FindFiles& oFindFiles );

        // Assignment operator - not allowed
        FindFiles& operator= ( const FindFiles& oFindFiles );

        // Methods
        UINT64 ComputeFileSize( const WIN32_FIND_DATA* pFindFileData );
        bool   FileNameHasExtension( const String& FileName,
                                     const StringArray& Extensions );
        void   Find( const String& DirectoryPath,
                     DWORD fileAttributes,
                     DWORD recursionDepth,
                     const StringArray& Extensions, const StringArray& DirectoryStopList );
        bool   FollowDirectory( const String& DirectoryPath,
                                DWORD fileAttributes,
                                DWORD recursionDepth, const StringArray& DirectoryStopList );

        // Data members
        DWORD   MAX_RECURSION_DEPTH;    // Maximum depth for a file scan
        TList< TYPE_SIZE_DATE_PATH > m_Files;
};

#endif  // WINAUDIT_FIND_FILES_H_
