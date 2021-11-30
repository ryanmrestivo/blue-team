///////////////////////////////////////////////////////////////////////////////////////////////////
//
// File Class Header
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

#ifndef PXSBASE_FILE_H_
#define PXSBASE_FILE_H_

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

// 6. Forwards
class ByteArray;
class String;
class StringArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class File
{
    public:
        // Default constructor
        File();

        // Destructor
        ~File();

        // Methods
        void    Close();
        void    CreateNew( const String& FileName, DWORD shareMode, bool unicode );
        void    Delete( const String& FileName );
        bool    Exists( const String& FileName ) const;
        void    Flush() const;
        UINT64  GetFilePointer() const;
        UINT64  GetSize() const;
        void    GetTimes( const String& FileName,
                          FILETIME* pCreation, FILETIME* pLastAccess, FILETIME* pLastWrite ) const;
        void    GetTimes( const String& FileName,
                          bool wantTime,
                          String* pCreation, String* pLastAccess, String* pLastWrite ) const;
        bool    IsOpen() const;
        bool    IsUnicode() const;
        bool    IsValidFileName( const String& FileName ) const;
        void    Open( const String& FileName,
                      DWORD desiredAccess, DWORD shareMode, DWORD numTries, bool isTextFile );
        void    OpenBinaryExclusive( const String& FileName );
        void    OpenExclusive( const String& FileName, DWORD numTries );
        void    OpenText( const String& FileName );
        size_t  Read( BYTE* pBuffer, size_t bufferBytes ) const;
        void    ReadAll( ByteArray* pBytes ) const;
        void    ReadAll( String* pContents ) const;
        size_t  ReadLine( String* pLine ) const;
        void    ReadLineArray( const String& FileName, DWORD maxLines, StringArray* pLines );
        void    Seek( UINT64 position ) const;
        void    SeekToEnd() const;
        void    SetSize( UINT64 size );
        void    Write( const void* pBuffer, size_t bufferBytes ) const;
        void    WriteChars( const String& Text ) const;

    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        File( const File& oFile );

        // Assignment operator - not allowed
        File& operator= ( const File& oFile );

        // Methods
        void    OnOpenDetectUtf16Bom();

        // Data members
        bool    m_bUnicodeLE;  // Little endian Unicode (BOM: 0xFF 0xFE )
        HANDLE  m_hFile;
};

#endif  // PXSBASE_FILE_H_
