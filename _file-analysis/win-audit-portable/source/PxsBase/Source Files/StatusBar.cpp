///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Status Bar Class Implementation
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
#include "PxsBase/Header Files/StatusBar.h"

// 2. C System Files
#include <stdint.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StaticControl.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
StatusBar::StatusBar()
{
    // Creation parameters
    m_CreateStruct.cx  = 100;
    m_CreateStruct.cy  = 20;

    // Properties
    try
    {
        SetBackground( GetSysColor( COLOR_BTNFACE ) );
        SetLayoutProperties( PXS_LAYOUT_STYLE_NONE, 2, 0 );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Copy constructor - not allowed so no implementation

// Destructor - do not throw any exceptions
StatusBar::~StatusBar()
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
//      Add a panel to the status bar        .
//
//  Parameters:
//      width - width of panel in pixels
//
//  Returns:
//      BYTE which is the number of the panel added.
//      Counting starts from zero
//===============================================================================================//
BYTE StatusBar::AddPanel( WORD width )
{
    RECT   panelBounds= { 0, 0, 0, 0 };
    SIZE   clientSize = { 0, 0 };
    size_t numPanels  = m_Statics.GetSize();
    StaticControl Panel;

    // Don't add more than range of BYTE
    if ( numPanels >= INT8_MAX )
    {
        throw BoundsException( L"numPanels >= INT8_MAX", __FUNCTION__ );
    }

    panelBounds.right  = panelBounds.left + width;
    panelBounds.top    = 0;
    panelBounds.bottom = clientSize.cy;
    Panel.SetBounds( panelBounds );
    Panel.SetShape( PXS_SHAPE_RAISED );
    Panel.SetShapeColour( m_crForeground );
    Panel.SetAlignmentY( PXS_CENTER_ALIGNMENT );
    Panel.SetPadding( 3 );
    m_Statics.Add( Panel );
    DoLayout();
    Repaint();

    return static_cast<BYTE>( 0xff & numPanels );
}

//===============================================================================================//
//  Description:
//      Layout the panels on this status bar
//
//  Parameters:
//      none
//
//  Remarks:
//      Row layout, left/right aligned for LTR/RTL reading respectibly;
//  Returns:
//      void
//===============================================================================================//
void StatusBar::DoLayout()
{
    int    xPos = 0, horizontalGap = 0, verticalGap = 0, width = 0;
    bool   rtlReading;
    RECT   bounds      = { 0, 0, 0, 0 };
    SIZE   clientSize  = { 0, 0 };
    DWORD  layoutStyle = 0;
    size_t i = 0, numPanels = m_Statics.GetSize();
    StaticControl Panel;

    if ( m_hWindow == nullptr )
    {
        return;     // Nothing to layout
    }
    rtlReading = IsRightToLeftReading();
    GetLayoutProperties( &layoutStyle, &horizontalGap, &verticalGap );

    GetClientSize( &clientSize );
    if ( rtlReading )
    {
        xPos = clientSize.cx - horizontalGap;
    }

    for ( i = 0; i < numPanels; i++ )
    {
        Panel = m_Statics.Get( i );
        Panel.GetBounds( &bounds );
        width = bounds.right - bounds.left;
        if ( rtlReading )
        {
            xPos -= ( width + horizontalGap );
            Panel.SetAlignmentX( PXS_RIGHT_ALIGNMENT );
        }
        else
        {
            Panel.SetAlignmentX( PXS_LEFT_ALIGNMENT );
        }
        bounds.left   = xPos;
        bounds.right  = bounds.left + width;
        bounds.top    = verticalGap;
        bounds.bottom = clientSize.cy - verticalGap;
        Panel.SetBounds( bounds );
        m_Statics.Set( i, Panel );

        // Set position for next LTR panel
        if ( rtlReading == false )
        {
            xPos += ( width + horizontalGap );
        }
    }

    PositionHideBitmap();   // Do last so its on top
}

//===============================================================================================//
//  Description:
//      Get the preferred height of the status bar
//
//  Parameters:
//      None
//
//  Returns:
//      Preferred height of the status bar
//===============================================================================================//
int StatusBar::GetPreferredHeight()
{
    int     height   = 20;      // Default value
    HDC     hdc      = nullptr;
    HFONT   hfont    = nullptr;
    HGDIOBJ oldFont  = nullptr;
    TEXTMETRIC tm;
    NONCLIENTMETRICS ncm;

    // Checks, return a nominal value
    if ( m_hWindow == nullptr )
    {
        return PXS_DEFAULT_SCREEN_LINE_HEIGHT;
    }

    hdc = GetDC( m_hWindow );
    if ( hdc == nullptr )
    {
        return PXS_DEFAULT_SCREEN_LINE_HEIGHT;
    }

    // Get the status bar font
    memset( &ncm, 0, sizeof ( ncm ) );
    ncm.cbSize = sizeof ( ncm );
    if ( SystemParametersInfo( SPI_GETNONCLIENTMETRICS, 0, &ncm, 0 ) )
    {
        hfont = CreateFontIndirect( &ncm.lfStatusFont );
        if ( hfont )
        {
            oldFont = SelectObject( hdc, hfont );
            memset( &tm, 0, sizeof ( tm ) );
            if ( GetTextMetrics( hdc, &tm ) )
            {
                height = tm.tmHeight + tm.tmExternalLeading;
            }

            // Reset
            if ( oldFont )
            {
                SelectObject( hdc, oldFont );
            }
            DeleteObject( hfont );
        }
    }
    ReleaseDC( m_hWindow, hdc );

    // Set a minimum of 16 and add padding above and below any content
    height  = PXSMaxInt( height, 16 );
    height += 2;
    height += 2;

    return height;
}

//===============================================================================================//
//  Description:
//      Remove all the panels from this status bar
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void StatusBar::RemoveAllPanels()
{
    m_Statics.RemoveAll();
}

//===============================================================================================//
//  Description:
//      Set the image for this panel
//
//  Parameters:
//      panelNumber - the number of the panel, counting starts at zero
//      resourceID  - the resource ID of the bitmap
//
//  Returns:
//      void
//===============================================================================================//
void StatusBar::SetBitmap( BYTE panelIndex, WORD resourceID )
{
    HBITMAP   hBitmap    = nullptr;
    String    ErrorDetails;
    Formatter Format;
    StaticControl Static;

    if ( panelIndex >= m_Statics.GetSize() )
    {
        throw BoundsException( L"panelIndex", __FUNCTION__ );
    }
    Static = m_Statics.Get( panelIndex );

    // Test if do not want a bitmap
    if ( resourceID == 0 )
    {
        Static.SetBitmap( nullptr );
        m_Statics.Set( panelIndex, Static );
        Repaint();
        return;
    }

    // Load the bitmap
    hBitmap = LoadBitmap( GetModuleHandle( nullptr ), MAKEINTRESOURCE( resourceID ) );
    if ( hBitmap == nullptr )
    {
        ErrorDetails = Format.StringUInt32( L" LoadBitmap, resourceID=%%1", resourceID );
        throw SystemException( GetLastError(), ErrorDetails.c_str(), __FUNCTION__);
    }

    // Need to clean up
    try
    {
        Static.SetBitmap( hBitmap );
        m_Statics.Set( panelIndex, Static );
        Repaint();
    }
    catch ( const Exception& )
    {
        DeleteObject( hBitmap );
        throw;
    }
    DeleteObject( hBitmap );
}

//===============================================================================================//
//  Description:
//      Set the horizontal alignment for a given status bar panel.
//
//  Parameters:
//      panelIndex - the number of the panel, counting starts at zero
//      alignmentX - named constant specifying horizontal alignment
//
//  Returns:
//      void
//===============================================================================================//
void StatusBar::SetAlignmentX( BYTE panelIndex, DWORD alignmentX )
{
    StaticControl Static;

    if ( panelIndex >= m_Statics.GetSize() )
    {
        throw BoundsException( L"panelIndex", __FUNCTION__ );
    }

    if ( ( alignmentX != PXS_LEFT_ALIGNMENT   ) &&
         ( alignmentX != PXS_CENTER_ALIGNMENT ) &&
         ( alignmentX != PXS_RIGHT_ALIGNMENT  )  )
    {
        throw ParameterException( L"alignmentX", __FUNCTION__ );
    }

    Static = m_Statics.Get( panelIndex );
    if ( alignmentX != Static.GetAlignmentX() )
    {
        Static.SetAlignmentX( alignmentX );
        m_Statics.Set( panelIndex, Static );
    }
    Repaint();
}

//===============================================================================================//
//  Description:
//      Set if this panel is a hyper-link
//
//  Parameters:
//      panelIndex - the number of the panel, counting starts at zero
//      hyperlink  - flag specifying if panel is a hyper-link
//
//  Returns:
//      void
//===============================================================================================//
void StatusBar::SetHyperlink( BYTE panelIndex, bool hyperlink )
{
    StaticControl Static;

    if ( panelIndex >= m_Statics.GetSize() )
    {
        throw BoundsException( L"panelIndex", __FUNCTION__ );
    }

    Static = m_Statics.Get( panelIndex );
    if ( hyperlink != Static.IsHyperlink() )
    {
        Static.SetHyperlink( hyperlink );
        m_Statics.Set( panelIndex, Static );
    }
    Repaint();
}

//===============================================================================================//
//  Description:
//      Set the text for the specified panel
//
//  Parameters:
//      panelIndex - the number of the panel, counting starts at zero
//      Text       - the text to display
//
//  Returns:
//      void
//===============================================================================================//
void StatusBar::SetText( BYTE panelIndex, const String& Text )
{
    Font  FontObject;
    StaticControl    Static;
    NONCLIENTMETRICS ncm;

    if ( panelIndex >= m_Statics.GetSize() )
    {
        throw BoundsException( L"panelIndex", __FUNCTION__ );
    }
    Static = m_Statics.Get( panelIndex );
    Static.SetSingleLineText( Text );

    // Use status bar font
    memset( &ncm, 0, sizeof ( ncm ) );
    ncm.cbSize = sizeof ( ncm );
    if ( SystemParametersInfo( SPI_GETNONCLIENTMETRICS, 0, &ncm, 0 ) )
    {
        FontObject.SetLogFont( &ncm.lfStatusFont );
        FontObject.Create();
        Static.SetFont( FontObject );
    }
    m_Statics.Set( panelIndex, Static );
    Repaint();
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Paint event handler
//
//  Parameters:
//      hdc - Handle to device context
//
//  Returns:
//      void
//===============================================================================================//
void StatusBar::PaintEvent( HDC hdc )
{
    RECT   clientRect = { 0, 0, 0, 0 };
    RECT   panelBounds= { 0, 0, 0, 0 };
    size_t i = 0, numStatics = 0;
    StaticControl StaticPanel, SeparatorPanel;

    if ( ( m_hWindow == nullptr ) || ( hdc == nullptr ) )
    {
        return;
    }
    DrawBackground( hdc );

    // Draw each static
    numStatics = m_Statics.GetSize();
    for ( i = 0; i < numStatics; i++)
    {
        // Set the font
        StaticPanel = m_Statics.Get( i );
        if ( StaticPanel.IsHyperlink() == false )
        {
            StaticPanel.SetForeground( m_crForeground );
        }
        StaticPanel.Draw( hdc );

        // Draw the separator
        StaticPanel.GetBounds( &panelBounds );
        panelBounds.left  = panelBounds.right - 2;
        panelBounds.right = panelBounds.left + 3;
        SeparatorPanel.SetBounds( panelBounds );
        SeparatorPanel.SetShape( PXS_SHAPE_RAISED );
        SeparatorPanel.Draw( hdc );
    }
    DrawHideBitmap( hdc );

    // Draw border
    if ( m_uBorderStyle != PXS_SHAPE_NONE )
    {
        GetClientRect( m_hWindow, &clientRect );
        StaticPanel.Reset();
        StaticPanel.SetShape( m_uBorderStyle );
        StaticPanel.SetBounds( clientRect );
        StaticPanel.SetBackground( CLR_INVALID );
        StaticPanel.Draw( hdc );
    }
}

//===============================================================================================//
//  Description:
//      Handle WM_SIZE event.
//
//  Parameters:
//      none
//
//  Returns:
//      void.
//===============================================================================================//
void StatusBar::SizeEvent()
{
    RECT   bounds     = { 0, 0, 0, 0 };
    RECT   clientRect = { 0, 0, 0, 0 };
    size_t i = 0, numStatics = m_Statics.GetSize();
    StaticControl Static;

    if ( m_hWindow == nullptr )
    {
        return;
    }

    // Refresh the heights of the statics
    GetClientRect( m_hWindow, &clientRect );
    for ( i = 0; i < numStatics; i++ )
    {
        Static = m_Statics.Get( i );
        Static.GetBounds( &bounds );
        if ( ( bounds.top    != clientRect.top )   ||
             ( bounds.bottom != clientRect.bottom ) )
        {
            // Set to the client's height, keeping the width constant
            bounds.top    = clientRect.top;
            bounds.bottom = clientRect.bottom;
            Static.SetBounds( bounds );
            m_Statics.Set( i, Static );
        }
    }
    Repaint();
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
