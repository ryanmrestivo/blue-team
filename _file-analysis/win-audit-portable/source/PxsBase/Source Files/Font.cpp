///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Font Class Implementation
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

// The font is initialised to the stock font DEFAULT_GUI_FONT. After setting
// any properties, the caller should call Create to recreate the font with its
// new properties.

///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/Font.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default Constructor
Font::Font()
     :m_bIsStockFont( false ),
      m_hFont(),
      m_LogFont()       // Non-op
{
    // Init to the stock DEFAULT_GUI_FONT
    try
    {
        SetStockFont( DEFAULT_GUI_FONT );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Copy constructor
Font::Font( const Font& oFont )
     :m_bIsStockFont( false ),
      m_hFont(),
      m_LogFont()       // Non-op
{
    *this = oFont;
}

// Destructor
Font::~Font()
{
    // Avoid deleting a stock font
    if ( m_hFont && ( m_bIsStockFont == false  ) )
    {
        DeleteObject( m_hFont );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
Font& Font::operator= ( const Font& oFont )
{
    HFONT hFont = nullptr;

    if ( this == &oFont ) return *this;

    // If the input is a stock font then just need the handle otherwise create
    // a new font.
    if ( oFont.m_bIsStockFont )
    {
        hFont = oFont.m_hFont;
    }
    else
    {
        hFont = CreateFontIndirect( &oFont.m_LogFont );
        if ( hFont == nullptr )
        {
            throw SystemException( GetLastError(), L"CreateFontIndirect", __FUNCTION__ );
        }
    }

    // Replace, avoid deleting a stock font
    if ( m_hFont && ( m_bIsStockFont == false  ) )
    {
        DeleteObject( m_hFont );
    }
    m_hFont        = hFont;
    m_bIsStockFont = oFont.m_bIsStockFont;
    memcpy( &m_LogFont, &oFont.m_LogFont, sizeof ( m_LogFont ) );

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Create the font
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Font::Create()
{
    HFONT   hFont = nullptr;
    LOGFONT logFont;

    memcpy( &logFont, &m_LogFont, sizeof ( logFont ) );
    hFont = CreateFontIndirect( &logFont );
    if ( hFont == nullptr )
    {
        throw SystemException( GetLastError(), L"CreateFontIndirect", __FUNCTION__ );
    }

    // Replace, avoid deleting a stock font
    if ( m_hFont && ( m_bIsStockFont == false  ) )
    {
        DeleteObject( m_hFont );
    }
    m_hFont        = hFont;
    m_bIsStockFont = false;
    memcpy( &m_LogFont, &logFont, sizeof ( logFont ) );
}

//===============================================================================================//
//  Description:
//      Decrement the font size
//
//  Parameters:
//      void
//
//  Returns:
//      void
//===============================================================================================//
void Font::DecrementLogicalHeight()
{
    // Different meaning for lfHeight can be <0, 0 or >0
    if ( m_LogFont.lfHeight >= 0 )
    {
        m_LogFont.lfHeight--;
    }
    else
    {
        m_LogFont.lfHeight++;
    }
}

//===============================================================================================//
//  Description:
//      Get the character set
//
//  Parameters:
//      none
//
//  Returns:
//      defined constant, see LOGFONT
//===============================================================================================//
BYTE Font::GetCharacterSet()
{
    return m_LogFont.lfCharSet;
}

//===============================================================================================//
//  Description:
//      Get the HFONT
//
//  Parameters:
//      none
//
//  Returns:
//      HFONT
//===============================================================================================//
HFONT Font::GetHandle() const
{
    return m_hFont;
}

//===============================================================================================//
//  Description:
//      Get the logical font structure
//
//  Parameters:
//      logFont - the input log font
//
//  Returns:
//      void
//===============================================================================================//
void Font::GetLogFont( LOGFONT* pLogFont ) const
{
    if ( pLogFont == nullptr )
    {
        throw ParameterException( L"pLogFont", __FUNCTION__ );
    }
    memcpy( pLogFont, &m_LogFont, sizeof ( m_LogFont ) );
}

//===============================================================================================//
//  Description:
//      Get the point size of the font in the specified device context
//
//  Parameters:
//      hdc - the device context, NULL implies the desktop
//
//  Returns:
//      the point size
//===============================================================================================//
int Font::GetSize( HDC hdc )
{
    int logPixelsY = 0, pointSize = 0;
    HDC hdcDesktop = nullptr;

    if ( hdc )
    {
        logPixelsY = GetDeviceCaps( hdc, LOGPIXELSY );
    }
    else
    {
        // Use the desktop
        hdcDesktop = GetDC( nullptr );
        if ( hdcDesktop )
        {
            logPixelsY = GetDeviceCaps( hdcDesktop, LOGPIXELSY );
            ReleaseDC( nullptr, hdcDesktop );
        }
    }

    if ( logPixelsY <= 0 )
    {
        throw SystemException( ERROR_INVALID_DATA, L"logPixelsY", __FUNCTION__);
    }
    pointSize = ( 72 * m_LogFont.lfHeight ) / logPixelsY;
    if ( pointSize < 0 )
    {
        pointSize = -pointSize;
    }

    return pointSize;
}

//===============================================================================================//
//  Description:
//      Decrement the font size
//
//  Parameters:
//      void
//
//  Returns:
//      void
//===============================================================================================//
void Font::IncrementLogicalHeight()
{
    // Different meaning for lfHeight can be <0, 0 or >0
    if ( m_LogFont.lfHeight >= 0 )
    {
        m_LogFont.lfHeight++;
    }
    else
    {
        m_LogFont.lfHeight--;
    }
}

//===============================================================================================//
//  Description:
//      Get if this font is bold
//
//  Parameters:
//      None
//
//  Returns:
//      true if bold else false
//===============================================================================================//
bool Font::IsBold() const
{
    if ( m_LogFont.lfWeight == FW_BOLD )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Get if this font is in italics
//
//  Parameters:
//      None
//
//  Returns:
//      true if font is italicised, else false
//===============================================================================================//
bool Font::IsItalic() const
{
    if ( m_LogFont.lfItalic )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Get if this font is underlined
//
//  Parameters:
//      None
//
//  Returns:
//      true if font is underlined, else false
//===============================================================================================//
bool Font::IsUnderlined() const
{
    if ( m_LogFont.lfUnderline )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Set if this font is bold or plain
//
//  Parameters:
//      bold - flag to set font as bold
//
//  Returns:
//      void
//===============================================================================================//
void Font::SetBold( bool bold )
{
    if ( bold )
    {
        m_LogFont.lfWeight = FW_BOLD;
    }
    else
    {
        m_LogFont.lfWeight = FW_NORMAL;
    }
}

//===============================================================================================//
//  Description:
//      Set the character set
//
//  Parameters:
//      characterSet - defined constant, see LOGFONT
//
//  Returns:
//      void
//===============================================================================================//
void Font::SetCharacterSet( BYTE characterSet )
{
    m_LogFont.lfCharSet = characterSet;
}

//===============================================================================================//
//  Description:
//      Set this font to italic or plain
//
//  Parameters:
//      italic - set to true to set font to italics
//
//  Returns:
//      void
//===============================================================================================//
void Font::SetItalic( bool italic )
{
    if ( italic )
    {
        m_LogFont.lfItalic = TRUE;
    }
    else
    {
        m_LogFont.lfItalic = FALSE;
    }
}

//===============================================================================================//
//  Description:
//      Set a log font structure
//
//  Parameters:
//      pLogfont - pointer to a LOGFONT structure
//
//  Returns:
//      void
//===============================================================================================//
void Font::SetLogFont( const LOGFONT* pLogfont )
{
    if ( pLogfont == nullptr )
    {
        throw ParameterException( L"pLogfont", __FUNCTION__ );
    }
    memcpy( &m_LogFont, pLogfont, sizeof ( m_LogFont ) );
}

//===============================================================================================//
//  Description:
//      Set the face name of the font
//
//  Parameters:
//      pszFaceName - pointer to the font's face name
//
//  Remarks:
//      Maximum length is 32 characters including the null terminator
//
//  Returns:
//      Reference to this object
//===============================================================================================//
void Font::SetFaceName( LPCWSTR pszFaceName )
{
    String FaceName;

    // Null is allowed, this enables the system to select the first
    // font that matches the LOGFONT characteristics.
    if ( pszFaceName == nullptr )
    {
        memset( m_LogFont.lfFaceName, 0, sizeof ( m_LogFont.lfFaceName ) );
        return;
    }

    // Copy, ensure it fits
    FaceName = pszFaceName;
    FaceName.Trim();
    if ( FaceName.GetLength() >= ARRAYSIZE( m_LogFont.lfFaceName ) )
    {
        throw ParameterException( L"pszFaceNameName", __FUNCTION__ );
    }
    StringCchCopy( m_LogFont.lfFaceName, ARRAYSIZE( m_LogFont.lfFaceName ), FaceName.c_str() );
}

//===============================================================================================//
//  Description:
//      Set the point size of the font
//
//  Parameters:
//      pointSize - point size of the font
//      hdc       - handle to the device context, NULL implies desktop
//
//  Returns:
//      void
//===============================================================================================//
void Font::SetPointSize( int pointSize, HDC hdc )
{
    int logPixelsY = 0;
    HDC hdcDesktop = nullptr;

    if ( pointSize <= 0 )
    {
        throw ParameterException( L"pointSize", __FUNCTION__ );
    }

    if ( hdc )
    {
        logPixelsY = GetDeviceCaps( hdc, LOGPIXELSY );
    }
    else
    {
        // Use the desktop
        hdcDesktop = GetDC( nullptr );
        if ( hdcDesktop )
        {
            logPixelsY = GetDeviceCaps( hdcDesktop, LOGPIXELSY );
            ReleaseDC( nullptr, hdcDesktop );
        }
    }

    if ( logPixelsY <= 0 )
    {
        throw SystemException( ERROR_INVALID_DATA, L"logPixelsY", __FUNCTION__ );
    }
    m_LogFont.lfHeight = -MulDiv( pointSize, logPixelsY, 72 );
}

//===============================================================================================//
//  Description:
//      Set the font to a system stock font
//
//  Parameters:
//      stockFont - stock font type, e.g. ANSI_FIXED_FONT
//
//  Returns:
//      void
//===============================================================================================//
void Font::SetStockFont( int stockFont )
{
    HFONT   hFont;
    LOGFONT logFont;

    hFont = PXSGetStockFont( stockFont );
    if ( hFont == nullptr )
    {
        throw SystemException( GetLastError(), L"PXSGetStockObject", __FUNCTION__ );
    }

    memset( &logFont, 0, sizeof ( logFont ) );
    if ( GetObject( hFont, sizeof ( logFont ), &logFont ) == 0 )
    {
        throw SystemException( GetLastError(), L"GetObject", __FUNCTION__ );
    }

    // Replace, avoid deleting a stock font
    if ( m_hFont && ( m_bIsStockFont == false ) )
    {
        DeleteObject( m_hFont );
    }
    m_hFont        = hFont;
    m_bIsStockFont = true;
    memcpy( &m_LogFont, &logFont, sizeof ( logFont ) );
}

//===============================================================================================//
//  Description:
//      Set if this font is underlined
//
//  Parameters:
//      underlined - flag to set the font underlined.
//
//  Returns:
//      void
//===============================================================================================//
void Font::SetUnderlined( bool underlined )
{
    if ( underlined )
    {
        m_LogFont.lfUnderline = TRUE;
    }
    else
    {
        m_LogFont.lfUnderline = FALSE;
    }
}


///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
