///////////////////////////////////////////////////////////////////////////////////////////////////
//
// MAPI Mail Class Implementation
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
#include "PxsBase/Header Files/Mail.h"

// 2. C System Files
#include <MAPI.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateChars.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/Library.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
Mail::Mail()
{
}

//
// Copy constructor - not allowed so no implementation
//

// Destructor
Mail::~Mail()
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
//      Determine if the system has the Messaging Application Programming
//      Interface
//
//  Parameters:
//      None
//
//  Remarks:
//      Does not throw, on error return false, i.e. say no MAPI
//
//  Returns:
//      true if can load MAPI32.dll, else false.
//===============================================================================================//
bool Mail::HasMapi32()
{
    bool    mapi32 = false;
    Library LibMapi32;

    try
    {
        LibMapi32.LoadSystemLibrary( L"MAPI32.DLL" );
        mapi32 = true;
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }

    return mapi32;
}

//===============================================================================================//
//  Description:
//      Send a simple email message, i.e. one recipient, no attachments, text
//      format. Message is sent to the in-box of the default mail client
//      programme which is then actually responsible for sending.
//
//  Parameters:
//      Address  - the recipient's email address
//      Subject  - the subject of the e-mail
//      NoteText - the message body of the e-mail
//      AttachFilePath - path to an attached file, use "" if no attachment
//
//  Remarks:
//      Hmmm... MAPISendMail returns a different codes for different
//      versions of outlook express when the user presses the close button of
//      the compose message window. E.g. V5/NT4+SP6 returns success, whereas
//      vXX/XP+SP2 returns MAPI_E_USER_ABORT.
//
//      Hmmm... When there is no note text and a html attachment is specified
//      Outlook Express uses the html as the note text.
//
//  Returns:
//      void, else false.
//===============================================================================================//
void Mail::SendSimpleMail( const String& Address,
                           const String& Subject,
                           const String& NoteText, const String& AttachFilePath)
{
    File    FileObject;
    char    szPathName[ MAX_PATH + 1 ] = { 0 };  // Attachment path must be ansi
    ULONG   result = 0;
    char*   pszAnsiAddress  = nullptr;
    char*   pszAnsiSubject  = nullptr;
    char*   pszAnsiNoteText = nullptr;
    wchar_t szCurrentDirectory[ MAX_PATH + 1 ] = { 0 };
    size_t  numBytes = 0;
    Library LibMapi32;
    Formatter      Format;
    MapiMessage    Message;
    MapiFileDesc   FileDesc;
    MapiRecipDesc  RecipDesc;
    AllocateChars  AllocAddress, AllocSubject, AllocNoteText;
    LPMAPISENDMAIL pfnMapiSendMail = nullptr;

    // Get the MAPISendMail function
    LibMapi32.LoadSystemLibrary( L"MAPI32.DLL" );

    // Disable C4191 - unsafe conversion from 'type of expression'
    // to 'type required'
    #ifdef _MSC_VER
        #pragma warning( push )
        #pragma warning ( disable : 4191 )
    #endif

    pfnMapiSendMail = (LPMAPISENDMAIL)LibMapi32.ProcAddress( "MAPISendMail" );

    #ifdef _MSC_VER
        #pragma warning( pop )
    #endif

    if ( pfnMapiSendMail == nullptr )
    {
        throw SystemException( ERROR_PROC_NOT_FOUND,
                               L"MAPISendMail", __FUNCTION__ );
    }
    memset( &Message, 0, sizeof ( Message ) );

    // Recipient address is optional
    if ( Address.GetLength() )
    {
        numBytes = Address.GetAnsiMultiByteLength();
        pszAnsiAddress = AllocAddress.New( numBytes );
        Format.StringToAnsi( Address, pszAnsiAddress, numBytes );
        pszAnsiAddress[ numBytes - 1 ] = 0x00;

        memset( &RecipDesc, 0, sizeof ( RecipDesc ) );
        RecipDesc.ulRecipClass = MAPI_TO;
        RecipDesc.lpszName     = pszAnsiAddress;
        RecipDesc.lpszAddress  = pszAnsiAddress;
        Message.nRecipCount    = 1;
        Message.lpRecips       = &RecipDesc;
    }

    // Subject is optional
    if ( Subject.GetLength() )
    {
        numBytes = Subject.GetAnsiMultiByteLength();
        pszAnsiSubject = AllocSubject.New( numBytes );
        Format.StringToAnsi( Subject, pszAnsiSubject, numBytes );
        pszAnsiSubject[ numBytes - 1 ] = 0x00;
        Message.lpszSubject  = pszAnsiSubject;
    }

    // Note text
    if ( NoteText.GetLength() )
    {
        numBytes = NoteText.GetAnsiMultiByteLength();
        pszAnsiNoteText = AllocNoteText.New( numBytes );
        Format.StringToAnsi( NoteText, pszAnsiNoteText, numBytes );
        pszAnsiNoteText[ numBytes - 1 ] = 0x00;
        Message.lpszNoteText = pszAnsiNoteText;
    }

    // Optional attachment, needs to be an ANSI file path
    memset( &FileDesc, 0, sizeof ( FileDesc ) );
    if ( AttachFilePath.GetLength() )
    {
        if ( FileObject.Exists( AttachFilePath ) == false )
        {
            throw SystemException( ERROR_FILE_NOT_FOUND, AttachFilePath.c_str(), __FUNCTION__ );
        }
        Format.StringToAnsi( AttachFilePath, szPathName, ARRAYSIZE( szPathName ) );
        FileDesc.nPosition    = 0xFFFFFFFF;     // -1;
        FileDesc.lpszPathName = szPathName;

        Message.nFileCount = 1;
        Message.lpFiles    = &FileDesc;
    }

    // MapiSendMail seems to change the working directory, so will restore it
    // after mail has been sent
    GetCurrentDirectory( ARRAYSIZE( szCurrentDirectory ), szCurrentDirectory );
    szCurrentDirectory[ ARRAYSIZE( szCurrentDirectory ) - 1 ] = PXS_CHAR_NULL;

    result = pfnMapiSendMail( 0, 0, &Message, MAPI_LOGON_UI | MAPI_DIALOG, 0 );

    // Reset
    if ( szCurrentDirectory[ 0 ] != PXS_CHAR_NULL )
    {
        SetCurrentDirectory( szCurrentDirectory );
    }

    // Ignore if the user aborts
    if ( ( result != SUCCESS_SUCCESS ) && ( result != MAPI_E_USER_ABORT ) )
    {
        throw Exception( PXS_ERROR_TYPE_MAPI, result, L"MAPISendMail", __FUNCTION__ );
    }

    // No clean up required for MAPISendMail
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
