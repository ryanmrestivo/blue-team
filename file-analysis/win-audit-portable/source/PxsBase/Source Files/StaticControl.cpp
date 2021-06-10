///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Static Drawable Control Class Implementation
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
#include "PxsBase/Header Files/StaticControl.h"

// 2. C System Files
#include <math.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor. Keep values in sync with the Reset method
StaticControl::StaticControl()
              :m_bEnabled( true ),
               m_bVerticalGradient( false ),
               m_bHyperlink( false ),
               m_bVisible( false ),
               m_bSingleLine( false ),
               m_nPadding( 2 ),
               m_uShape( PXS_SHAPE_NONE ),
               m_uAlignmentX( PXS_LEFT_ALIGNMENT ),
               m_uAlignmentY( PXS_TOP_ALIGNMENT ),
               m_rBounds(),
               m_crBackground( CLR_INVALID ),
               m_crForeground( CLR_INVALID ),
               m_crGradient1( CLR_INVALID ),
               m_crGradient2( CLR_INVALID ),
               m_crShapeColour( CLR_INVALID ),
               m_Font(),
               m_Text(),
               m_hIcon( nullptr ),
               m_hBitmap( nullptr )
{
    m_rBounds.left   = 0;
    m_rBounds.top    = 0;
    m_rBounds.right  = 100;
    m_rBounds.bottom = 100;
}

// Copy constructor
StaticControl::StaticControl( const StaticControl& oStaticControl )
              :m_bEnabled( true ),
               m_bVerticalGradient( false ),
               m_bHyperlink( false ),
               m_bVisible( false ),
               m_bSingleLine( false ),
               m_nPadding( 2 ),
               m_uShape( PXS_SHAPE_NONE ),
               m_uAlignmentX( PXS_LEFT_ALIGNMENT ),
               m_uAlignmentY( PXS_TOP_ALIGNMENT ),
               m_rBounds(),
               m_crBackground( CLR_INVALID ),
               m_crForeground( CLR_INVALID ),
               m_crGradient1( CLR_INVALID ),
               m_crGradient2( CLR_INVALID ),
               m_crShapeColour( CLR_INVALID ),
               m_Font(),
               m_Text(),
               m_hIcon( nullptr ),
               m_hBitmap( nullptr )
{
    m_rBounds.left   = 0;
    m_rBounds.top    = 0;
    m_rBounds.right  = 100;
    m_rBounds.bottom = 100;

    *this = oStaticControl;
}

// Destructor
StaticControl::~StaticControl()
{
    try
    {
        Reset();
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
StaticControl& StaticControl::operator= ( const StaticControl& oStaticControl )
{
    if ( this == &oStaticControl ) return *this;

    m_bEnabled          = oStaticControl.m_bEnabled;
    m_bVerticalGradient = oStaticControl.m_bVerticalGradient;
    m_bHyperlink        = oStaticControl.m_bHyperlink;
    m_bVisible          = oStaticControl.m_bVisible;
    m_bSingleLine       = oStaticControl.m_bSingleLine;
    m_nPadding          = oStaticControl.m_nPadding;
    m_uShape            = oStaticControl.m_uShape;
    m_uAlignmentX       = oStaticControl.m_uAlignmentX;
    m_uAlignmentY       = oStaticControl.m_uAlignmentY;
    m_rBounds           = oStaticControl.m_rBounds;
    m_crGradient1       = oStaticControl.m_crGradient1;
    m_crGradient2       = oStaticControl.m_crGradient2;
    m_crForeground      = oStaticControl.m_crForeground;
    m_crBackground      = oStaticControl.m_crBackground;
    m_crShapeColour     = oStaticControl.m_crShapeColour;
    m_Text              = oStaticControl.m_Text;
    m_Font              = oStaticControl.m_Font;

    // Copying these requires resource allocation
    SetIcon( oStaticControl.m_hIcon );
    SetBitmap( oStaticControl.m_hBitmap );

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Draw this static object
//
//  Parameters:
//      hdc - handle to device context to draw on
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::Draw( HDC hdc )
{
    HRGN clipRegion;

    DrawStaticBackground( hdc );
    DrawStaticShape( hdc );

    // Set a clipping region
    clipRegion = CreateRectRgn( m_rBounds.left   + m_nPadding,
                                m_rBounds.top    + m_nPadding,
                                m_rBounds.right  - m_nPadding,
                                m_rBounds.bottom - m_nPadding );
    SelectClipRgn ( hdc, clipRegion );

    DrawStaticText( hdc );
    DrawStaticBitmap( hdc );
    DrawStaticIcon( hdc );

    // Reset
    SelectClipRgn ( hdc, nullptr );
    if ( clipRegion )
    {
        DeleteObject( clipRegion );
    }
}

//===============================================================================================//
//  Description:
//      Get the horizontal alignment property
//
//  Parameters:
//      None
//
//  Returns:
//      Defined horizontal alignment constant
//===============================================================================================//
DWORD StaticControl::GetAlignmentX() const
{
    return m_uAlignmentX;
}

//===============================================================================================//
//  Description:
//      Get the vertical alignment
//
//  Parameters:
//      None
//
//  Returns:
//      A defined vertical alignment constant
//===============================================================================================//
DWORD StaticControl::GetAlignmentY() const
{
    return m_uAlignmentY;
}

//===============================================================================================//
//  Description:
//      Get the background colour
//
//  Parameters:
//      None
//
//  Returns:
//      COLORREF of background colour.
//===============================================================================================//
COLORREF StaticControl::GetBackground() const
{
    return m_crBackground;
}

//===============================================================================================//
//  Description:
//      Get this static's bounding rectangle
//
//  Parameters:
//      pBounds - receives the bounds
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::GetBounds( RECT* pBounds ) const
{
    if ( pBounds == nullptr )
    {
        throw ParameterException( L"pBounds", __FUNCTION__ );
    }
    memcpy( pBounds, &m_rBounds, sizeof ( RECT ) );
}

//===============================================================================================//
//  Description:
//      Get the font object for associated with this control
//
//  Parameters:
//      None
//
//  Returns:
//      Reference to the font object
//===============================================================================================//
const Font& StaticControl::GetFont() const
{
    return m_Font;
}

//===============================================================================================//
//  Description:
//      Get the foreground colour for this control
//
//  Parameters:
//      None
//
//  Returns:
//      COLORREF of foreground colour.
//===============================================================================================//
COLORREF StaticControl::GetForeground() const
{
    return m_crForeground;
}

//===============================================================================================//
//  Description:
//      Get the padding for this control
//
//  Parameters:
//      None
//
//  Returns:
//      Signed integer of padding
//===============================================================================================//
int StaticControl::GetPadding() const
{
    return m_nPadding;
}

//===============================================================================================//
//  Description:
//      Get type of shape to draw on window
//
//  Parameters:
//      None
//
//  Returns:
//      A defined shape constant
//===============================================================================================//
DWORD StaticControl::GetShape() const
{
    return m_uShape;
}

//===============================================================================================//
//  Description:
//      Get the static's text
//
//  Parameters:
//      None
//
//  Returns:
//     Constant reference to a string object
//===============================================================================================//
const String& StaticControl::GetText() const
{
    return m_Text;
}

//===============================================================================================//
//  Description:
//      Get if this static is visible
//
//  Parameters:
//      None
//
//  Returns:
//      true if visible, else false
//===============================================================================================//
bool StaticControl::GetVisible() const
{
    return m_bVisible;
}

//===============================================================================================//
//  Description:
//      Get if this static is enabled
//
//  Parameters:
//      None
//
//  Returns:
//      true if enabled, else false
//===============================================================================================//
bool StaticControl::IsEnabled() const
{
    return m_bEnabled;
}

//===============================================================================================//
//  Description:
//      Get if this static is a hyper-link
//
//  Parameters:
//      None
//
//  Returns:
//      true if hyper-link, else false
//===============================================================================================//
bool StaticControl::IsHyperlink() const
{
    return m_bHyperlink;
}

//===============================================================================================//
//  Description:
//      Reset the static control to default values.
//      Frees any resources used by the control.
//
//  Parameters:
//      None
//
//  Remarks:
//      Frees any resources used by the control.
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::Reset()
{
    // Keep values in sync with the constructor
    m_bEnabled          = true;
    m_bVerticalGradient = false;
    m_bHyperlink        = false;
    m_bVisible          = false;
    m_bSingleLine       = false;
    m_nPadding          = 2;
    m_uShape            = PXS_SHAPE_NONE;
    m_uAlignmentX       = PXS_LEFT_ALIGNMENT;
    m_uAlignmentY       = PXS_TOP_ALIGNMENT;
    m_rBounds.left      = 0;
    m_rBounds.top       = 0;
    m_rBounds.right     = 100;
    m_rBounds.bottom    = 100;
    m_crBackground      = CLR_INVALID;
    m_crForeground      = CLR_INVALID;
    m_crGradient1       = CLR_INVALID;
    m_crGradient2       = CLR_INVALID;
    m_crShapeColour     = CLR_INVALID;
    m_Text              = PXS_STRING_EMPTY;
    m_Font.SetStockFont( DEFAULT_GUI_FONT );

    // Free resources
    if ( m_hIcon )
    {
        DestroyIcon( m_hIcon );
        m_hIcon = nullptr;
    }

    if ( m_hBitmap )
    {
        DeleteObject( m_hBitmap );
        m_hBitmap = nullptr;
    }
}

//===============================================================================================//
//  Description:
//      Set the horizontal alignment property
//
//  Parameters:
//      alignmentX - defined horizontal alignment constant
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetAlignmentX( DWORD alignmentX )
{
    // Ignore errors
    if ( ( alignmentX == PXS_LEFT_ALIGNMENT   ) ||
         ( alignmentX == PXS_CENTER_ALIGNMENT ) ||
         ( alignmentX == PXS_RIGHT_ALIGNMENT  )  )
    {
        m_uAlignmentX = alignmentX;
    }
}

//===============================================================================================//
//  Description:
//      Set vertical alignment for the drawing of this static control
//
//  Parameters:
//      alignmentY - defined vertical alignment constant
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetAlignmentY( DWORD alignmentY )
{
    // Ignore parameter errors
    if ( alignmentY == PXS_TOP_ALIGNMENT    ||
         alignmentY == PXS_CENTER_ALIGNMENT ||
         alignmentY == PXS_BOTTOM_ALIGNMENT  )
    {
        m_uAlignmentY = alignmentY;
    }
}

//===============================================================================================//
//  Description:
//      Set the background colour
//
//  Parameters:
//      background - colour reference of the background colour
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetBackground( COLORREF background )
{
    m_crBackground = background;
}

//===============================================================================================//
//  Description:
//      Set the background gradient properties
//
//  Parameters:
//      colour1 - the first colour
//      colour2 - the second colour
//      verticalGradient - if true the gradient is vertical otherwise
//                          horizontal
//  Remarks:
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetBackgroundGradient( COLORREF colour1,
                                           COLORREF colour2, bool verticalGradient )
{
    m_crGradient1     = colour1;
    m_crGradient2     = colour2;
    m_bVerticalGradient = verticalGradient;
}

//===============================================================================================//
//  Description:
//      Set the bitmap to draw for this static
//
//  Parameters:
//      hBitmap - handle to bitmap
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetBitmap( HBITMAP hBitmap )
{
    // Delete any existing bitmap
    if ( m_hBitmap )
    {
        DeleteObject( m_hBitmap );
        m_hBitmap = nullptr;
    }

    if ( hBitmap )
    {
        m_hBitmap = (HBITMAP)CopyImage( hBitmap, IMAGE_BITMAP, 0, 0, 0 );
        if ( m_hBitmap == nullptr )
        {
            throw SystemException( GetLastError(), L"CopyImage", __FUNCTION__ );
        }
    }
}

//===============================================================================================//
//  Description:
//      Set the bounds of this static
//
//  Parameters:
//      bounds - the bounding rectangle
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetBounds( const RECT& bounds )
{
    memcpy( &m_rBounds, &bounds, sizeof ( m_rBounds ) );
}

//===============================================================================================//
//  Description:
//      Set if this static is enabled
//
//  Parameters:
//      enabled - if true this static is enabled
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetEnabled( bool enabled )
{
    m_bEnabled = enabled;
}

//===============================================================================================//
//  Description:
//      Set the font for any the static's text
//
//  Parameters:
//      Font - the font object
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetFont( const Font& Font )
{
    m_Font = Font;
}

//===============================================================================================//
//  Description:
//      Set the foreground colour for this control
//
//  Parameters:
//      foreground - the foreground's colour reference
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetForeground( COLORREF foreground )
{
    m_crForeground = foreground;
}

//===============================================================================================//
//  Description:
//      Set if this static represents a hyper-link
//
//  Parameters:
//      hyper-link - if true this static is a hyper-link
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetHyperlink( bool hyperlink )
{
    // For hyper-links, want underline
    if ( hyperlink )
    {
        m_Font.SetUnderlined( true );
    }
    else
    {
        m_Font.SetUnderlined( false );
    }
    m_Font.Create();
    m_bHyperlink = hyperlink;
}

//===============================================================================================//
//  Description:
//      Set icon for this static control
//
//  Parameters:
//      hIcon - handle to icon
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetIcon( HICON icon )
{
    if ( m_hIcon )
    {
        DestroyIcon( m_hIcon );
        m_hIcon = nullptr;
    }

    if ( icon )
    {
        m_hIcon = CopyIcon( icon );
        if ( m_hIcon == nullptr )
        {
            throw SystemException( GetLastError(), L"CopyIcon", __FUNCTION__ );
        }
    }
}

//===============================================================================================//
//  Description:
//      Set position of the static control on the window
//
//  Parameters:
//      x - x position
//      y - y position
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetLocation( long x, long y )
{
    long width  = m_rBounds.right - m_rBounds.left;
    long height = m_rBounds.top   - m_rBounds.bottom;

    m_rBounds.left   = x;
    m_rBounds.top    = y;
    m_rBounds.right  = m_rBounds.left + width;
    m_rBounds.bottom = m_rBounds.top  + height;
}

//===============================================================================================//
//  Description:
//      Set the padding for this control
//
//  Parameters:
//      padding - the padding in pixels
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetPadding( int padding )
{
    m_nPadding = padding;
}

//===============================================================================================//
//  Description:
//      Set type of shape to draw on window
//
//  Parameters:
//      shape  - a defined constant representing a shape
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetShape( DWORD shape )
{
    m_uShape = shape;
    if ( m_crShapeColour == CLR_INVALID )
    {
        m_crShapeColour = PXS_COLOUR_BLACK;
    }
}

//===============================================================================================//
//  Description:
//      Set the colour of the shape
//
//  Parameters:
//      shapeColour - the shape's colour
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetShapeColour( COLORREF shapeColour )
{
    m_crShapeColour = shapeColour;
}

//===============================================================================================//
//  Description:
//      Set a single line of text to draw on this static.
//
//  Parameters:
//      SingleLineText - the text for this static
//
//  Remarks:
//      Single lines of text can be aligned using the DT_x constants
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetSingleLineText( const String& SingleLineText )
{
    m_Text        = SingleLineText;
    m_bSingleLine = true;
}

//===============================================================================================//
//  Description:
//      Set text to draw on this static.
//
//  Parameters:
//      Text - the text for this static
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetText( const String& Text )
{
    m_Text        = Text;
    m_bSingleLine = false;
}

//===============================================================================================//
//  Description:
//      Set if this static is visible
//
//  Parameters:
//      visible - visibility flag
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::SetVisible( bool visisble )
{
    m_bVisible = visisble;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Draw the shape
//
//  Parameters:
//      hdc - the device context
//
//  Returns:
//      void
//===============================================================================================//
void StaticControl::DrawShape( HDC hdc )
{
    int x1 = m_rBounds.left;
    int y1 = m_rBounds.top;
    int x2 = m_rBounds.right;
    int y2 = m_rBounds.bottom;
    int cx = 0, cy = 0, thick = 0;
    SIZE    textSize = { 0, 0 };
    LPCWSTR pszText = nullptr;

    if ( hdc == nullptr )
    {
        return;     // Nothing to do
    }

    switch (  m_uShape )
    {
        default:
        case PXS_SHAPE_NONE:
            break;

        case PXS_SHAPE_LINE:

            MoveToEx( hdc, x1, y1, nullptr );
            LineTo( hdc, x2, y2 );
            break;

        case PXS_SHAPE_TRIANGLE_RIGHT:

            MoveToEx( hdc, x1    , y1, nullptr );
            LineTo(   hdc, x2 - 1, ( y1 + y2 ) / 2 );
            LineTo(   hdc, x1    , y2 - 1 );
            LineTo(   hdc, x1    , y1 );
            break;

        case PXS_SHAPE_TRIANGLE_LEFT:

            MoveToEx( hdc, x2 - 1, y1, nullptr );
            LineTo(   hdc, x2 - 1, y2 - 1 );
            LineTo(   hdc, x1    , ( y1 + y2 ) / 2 );
            LineTo(   hdc, x2 - 1, y1 );
            break;

        case PXS_SHAPE_RECTANGLE:
        case PXS_SHAPE_RECTANGLE_DOTTED:

            MoveToEx( hdc, x1    , y1, nullptr );
            LineTo(   hdc, x1    , y2 - 1 );
            LineTo(   hdc, x2 - 1, y2 - 1 );
            LineTo(   hdc, x2 - 1, y1 );
            LineTo(   hdc, x1, y1 );
            break;

        case PXS_SHAPE_RECTANGLE_HOLLOW:

            thick = (x2 - x1)/8;
            MoveToEx( hdc, thick     , y1, nullptr );
            LineTo(   hdc, x2-thick-1, y1 );
            LineTo(   hdc, x2-thick-1, y1+thick );
            LineTo(   hdc, thick     , y1+thick );
            LineTo(   hdc, thick     , y1 );

            MoveToEx( hdc, 0    , thick, nullptr );
            LineTo(   hdc, 0    , y2-thick-1 );
            LineTo(   hdc, thick, y2-thick-1 );
            LineTo(   hdc, thick, y1+thick );
            LineTo(   hdc, 0    , y1+thick );

            MoveToEx( hdc, thick     , y2-thick-1, nullptr );
            LineTo(   hdc, thick     , y2-1 );
            LineTo(   hdc, x2-thick-1, y2-1 );
            LineTo(   hdc, x2-thick-1, y2-thick-1 );
            LineTo(   hdc, thick     , y2-thick-1 );

            MoveToEx( hdc, x2-thick-1, y2-thick-1, nullptr );
            LineTo(   hdc, x2-1      , y2-thick-1 );
            LineTo(   hdc, x2-1      , y1+thick );
            LineTo(   hdc, x2-thick-1, y1+thick );
            LineTo(   hdc, x2-thick-1, y2-thick-1 );
            break;

        case PXS_SHAPE_DIAMOND:

            cx = (x2 - x1)/2;
            cy = (y2 - y1)/2;
            MoveToEx( hdc, cx  , y1, nullptr );
            LineTo(   hdc, x2-1, cy-1 );
            LineTo(   hdc, cx  , y2-2 );
            LineTo(   hdc, x1+1, cy-1 );
            LineTo(   hdc, cx  , y1   );
            break;

        case PXS_SHAPE_CROSS:

            cx    = (x2 - x1)/2;
            cy    = (y2 - y1)/2;
            thick = (x2 - x1)/8;
            MoveToEx( hdc, cx-thick, y1, nullptr );
            LineTo(   hdc, cx-thick, cy-thick );
            LineTo(   hdc, x1      , cy-thick );
            LineTo(   hdc, x1      , cy+thick );
            LineTo(   hdc, cx-thick, cy+thick );
            LineTo(   hdc, cx-thick, y2-1 );
            LineTo(   hdc, cx+thick, y2-1 );
            LineTo(   hdc, cx+thick, cy+thick );
            LineTo(   hdc, x2-1    , cy+thick );
            LineTo(   hdc, x2-1    , cy-thick );
            LineTo(   hdc, cx+thick, cy-thick );
            LineTo(   hdc, cx+thick, y1 );
            LineTo(   hdc, cx-thick, y1 );
           break;

        case PXS_SHAPE_FIVE_SIDES:

            cx = (x2 - x1)/2;
            MoveToEx( hdc, cx  , y1, nullptr );
            LineTo(   hdc, x2-1, y1 + ( x2-1-cx ) );    // 45 degrees
            LineTo(   hdc, x2-1, y2-1 );
            LineTo(   hdc, x1  , y2-1 );
            LineTo(   hdc, x1  , y1 + ( x2-cx ) );      // 45 degrees
            LineTo(   hdc, cx  , y1 );
            break;

        case PXS_SHAPE_CIRCLE:

            // Will draw as two semi-circles even though when start point
            // an end point are the same get a full circle
            Arc( hdc, x1, y1, x2, y2,  0, (y2 - y1)/2, x2, (y2 - y1)/2 );
            Arc( hdc, x1, y1, x2, y2, x2, (y2 - y1)/2,  0, (y2 - y1)/2 );
            break;

        case PXS_SHAPE_ARROW_DOWN:

            cx = (x2 - x1)/2;
            MoveToEx( hdc, cx-2  , y1, nullptr );
            LineTo(   hdc, cx + 2, y1 );
            LineTo(   hdc, cx + 2, y2 - ( x2-cx ) );
            LineTo(   hdc, x2 - 1, y2 - ( x2-cx ) );
            LineTo(   hdc, cx    , y2 - 1 );              // 45 degrees
            LineTo(   hdc, 1     , y2 - ( x2-cx ) );      // 45 degrees
            LineTo(   hdc, cx - 2, y2 - ( x2-cx ) );
            LineTo(   hdc, cx - 2, y1 );
            break;

        case PXS_SHAPE_ARROW_UP:

            cx = (x2 - x1)/2;
            MoveToEx( hdc, cx    , y1, nullptr );
            LineTo(   hdc, x2 - 1, y1 + ( x2-1-cx ) );      // 45 degrees
            LineTo(   hdc, cx + 2, y1 + ( x2-1-cx ) );
            LineTo(   hdc, cx + 2, y2 - 1 );
            LineTo(   hdc, cx - 2, y2 - 1 );
            LineTo(   hdc, cx - 2, y1 + ( x2-1-cx ) );      // 45 degrees
            LineTo(   hdc, x1 + 1, y1 + ( x2-1-cx ) );
            LineTo(   hdc, cx    , y1 );
            break;

        case PXS_SHAPE_RAISED:

            SetDCPenColor( hdc, PXS_COLOUR_WHITE );
            MoveToEx( hdc, x2, y1, nullptr );
            LineTo( hdc, x1, y1 );
            LineTo( hdc, x1, y2 );

            SetDCPenColor( hdc, PXS_COLOUR_GREY );
            MoveToEx( hdc, x2 - 1, y1, nullptr );
            LineTo(   hdc, x2 - 1, y2 - 1 );
            LineTo(   hdc, x1    , y2 - 1 );
            break;

        case PXS_SHAPE_SUNK:

            SetDCPenColor( hdc, PXS_COLOUR_GREY );
            MoveToEx( hdc, x2, y1, nullptr );
            LineTo(   hdc, x1, y1 );
            LineTo(   hdc, x1, y2 );

            SetDCPenColor( hdc, PXS_COLOUR_WHITE );
            MoveToEx( hdc, x2 - 1, y1, nullptr );
            LineTo(   hdc, x2 - 1, y2 );
            LineTo(   hdc, x1    , y2 - 1 );
            break;

        case PXS_SHAPE_FRAME:

            // Get the length of the text, will assume that hdc has
            // the font used for the text.
            pszText = m_Text.c_str();
            if ( pszText )
            {
                GetTextExtentPoint32( hdc, pszText, lstrlen( pszText ), &textSize );
            }

            y1 += 6;
            SetDCPenColor( hdc, PXS_COLOUR_WHITE );
            MoveToEx( hdc, x1 + 1, y1, nullptr );
            LineTo(   hdc, x1 + 1, y2 );
            LineTo(   hdc, x2    , y2 );
            LineTo(   hdc, x2    , y1 + 1 );
            LineTo(   hdc, x1 + textSize.cx, y1 + 1 );
            MoveToEx( hdc, x1 + 5, y1 + 1, nullptr );
            LineTo(   hdc, x1 + 1, y1 + 1 );

            SetDCPenColor( hdc, m_crShapeColour );
            MoveToEx( hdc, x1    , y1, nullptr );
            LineTo(   hdc, x1    , y2 - 1 );
            LineTo(   hdc, x2 - 1, y2 - 1 );
            LineTo(   hdc, x2 - 1, y1 );
            LineTo(   hdc, x1 + textSize.cx, y1 );
            MoveToEx( hdc, x1 + 5, y1, nullptr );
            LineTo(   hdc, x1    , y1 );
            break;

        case PXS_SHAPE_FRAME_RTL:

            // Get the length of the text, will assume that hdc has
            // the font used for the text.
            pszText = m_Text.c_str();
            if ( pszText )
            {
                GetTextExtentPoint32( hdc, pszText, lstrlen( pszText ), &textSize );
            }

            y1 += 6;
            SetDCPenColor( hdc, PXS_COLOUR_WHITE );
            MoveToEx( hdc, x1 + 1, y1, nullptr );
            LineTo(   hdc, x1 + 1, y2 );
            LineTo(   hdc, x2    , y2 );
            LineTo(   hdc, x2    , y1 + 1 );
            LineTo(   hdc, x2 - 5, y1 + 1 );
            MoveToEx( hdc, x2 - textSize.cx, y1 + 1, nullptr );
            LineTo(   hdc, x1 + 1, y1 + 1 );

            SetDCPenColor( hdc, m_crShapeColour );
            MoveToEx( hdc, x1    , y1, nullptr );
            LineTo(   hdc, x1    , y2 - 1 );
            LineTo(   hdc, x2 - 1, y2 - 1 );
            LineTo(   hdc, x2 - 1, y1 );
            LineTo(   hdc, x2 - 5, y1 );
            MoveToEx( hdc, x2 - textSize.cx, y1, nullptr );
            LineTo(   hdc, x1    , y1 );
            break;

        case PXS_SHAPE_3D_SUNK:

            SetDCPenColor( hdc, PXS_COLOUR_BLACK );
            MoveToEx( hdc, x2, y1, nullptr );
            LineTo(   hdc, x1, y1 );
            LineTo(   hdc, x1, y2 );

            SetDCPenColor( hdc, PXS_COLOUR_GREY );
            MoveToEx( hdc, x2 - 1, y1 + 1, nullptr );
            LineTo(   hdc, x1 + 1, y1 + 1 );
            LineTo(   hdc, x1 + 1, y2 - 1 );

            SetDCPenColor( hdc, PXS_COLOUR_WHITE );
            MoveToEx( hdc, x1    , y2 - 1, nullptr );
            LineTo(   hdc, x2    , y2 - 1 );
            LineTo(   hdc, x2 - 1, y1 );
            break;

        case PXS_SHAPE_3D_RAISED:

            SetDCPenColor( hdc, PXS_COLOUR_WHITE );
            MoveToEx( hdc, x1    , y2, nullptr );
            LineTo(   hdc, x1    , y1 );
            LineTo(   hdc, x2 - 2, y1 );

            SetDCPenColor( hdc, PXS_COLOUR_GREY );
            MoveToEx( hdc, x1 + 1, y2 - 2, nullptr );
            LineTo(   hdc, x2 - 2, y2 - 2 );
            LineTo(   hdc, x2 - 2, y1 + 1 );

            SetDCPenColor( hdc, PXS_COLOUR_BLACK );
            MoveToEx( hdc, x1    , y2 - 1, nullptr );
            LineTo(   hdc, x2 - 1, y2 - 1 );
            LineTo(   hdc, x2 - 1, y1 );
            break;

        case PXS_SHAPE_SHADOW:

            MoveToEx( hdc, x1    , y1, nullptr );
            LineTo(   hdc, x1    , y2 - 3 );
            LineTo(   hdc, x2 - 3, y2 - 3 );
            LineTo(   hdc, x2 - 3, y1 );
            LineTo(   hdc, x1    , y1 );

            SetDCPenColor( hdc, PXS_COLOUR_GREY );
            MoveToEx( hdc, x1 + 1, y2 - 2, nullptr );
            LineTo(   hdc, x2 - 2, y2 - 2 );
            LineTo(   hdc, x2 - 2, y1 + 1 );

            SetDCPenColor( hdc, PXS_COLOUR_LITEGREY );
            MoveToEx( hdc, x1 + 2, y2 - 1, nullptr );
            LineTo(   hdc, x2 - 1, y2 - 1 );
            LineTo(   hdc, x2 - 1, y1 + 2 );
            break;

        // Fall through
        case PXS_SHAPE_TAB:
        case PXS_SHAPE_TAB_DOTTED:
        case PXS_SHAPE_TAB_RTL:
        case PXS_SHAPE_TAB_DOTTED_RTL:

            MoveToEx( hdc, x1    , y2 - 1, nullptr );
            LineTo(   hdc, x1    , y1 + 4 );
            LineTo(   hdc, x1 + 4, y1     );        // Bevel
            LineTo(   hdc, x2 - 5, y1     );
            LineTo(   hdc, x2 - 1, y1 + 4 );        // Bevel
            LineTo(   hdc, x2 - 1, y2 - 1 );
            LineTo(   hdc, x1    , y2 - 1 );
            break;

        case PXS_SHAPE_MENU_BARS:

            y1 = ( y2 - y1 - 14 ) / 2;  // Three horizontal bars
            Rectangle( hdc, x1 + 1, y1     , x2 - 1, y1 +  2 );
            Rectangle( hdc, x1 + 1, y1 +  6, x2 - 1, y1 +  8 );
            Rectangle( hdc, x1 + 1, y1 + 12, x2 - 1, y1 + 14 );
            break;

        case PXS_SHAPE_HOME:

            cx = (x2 - x1)/2;
            MoveToEx( hdc, cx    , y1, nullptr );
            LineTo(   hdc, x2 - 1, y1 + ( x2-1-cx ) );      // 45 degrees
            LineTo(   hdc, x2 - 4, y1 + ( x2-1-cx ) );
            LineTo(   hdc, x2 - 4, y2 - 1 );
            LineTo(   hdc, cx + 3, y2 - 1 );
            LineTo(   hdc, cx + 3, y1 + ( x2-1-cx ) );
            LineTo(   hdc, cx - 3, y1 + ( x2-1-cx ) );
            LineTo(   hdc, cx - 3, y2 - 1 );
            LineTo(   hdc, x1 + 4, y2 - 1 );
            LineTo(   hdc, x1 + 4, y1 + ( x2-1-cx ) );
            LineTo(   hdc, x1 + 1, y1 + ( x2-1-cx ) );
            LineTo(   hdc, cx    , y1 );

            break;

        case PXS_SHAPE_REFRESH:

            cx = (x2 - x1)/2;
            cy = (y2 - y1)/2;
            MoveToEx( hdc, 0     , y2 - 1, nullptr );   // 1
            LineTo(   hdc, 0     , y1 );                // 2
            LineTo(   hdc, x2 - 1, y1     );            // 3
            LineTo(   hdc, x2 - 1, y2 - 4 );            // 4
            LineTo(   hdc, cx - 1, y2 - 4 );            // 5
            LineTo(   hdc, cx - 1, cy - 1 );            // 6
            LineTo(   hdc, cx - 4, cy - 1 );            // 7
            LineTo(   hdc, cx    , y1 + 5 );            // 8
            LineTo(   hdc, cx + 4, cy - 1 );            // 9
            LineTo(   hdc, cx + 1, cy - 1 );            // 10
            LineTo(   hdc, cx + 1, cy + 4 );            // 11
            LineTo(   hdc, x2 - 3, cy + 4 );            // 12
            LineTo(   hdc, x2 - 3, y1 + 2 );            // 13
            LineTo(   hdc, x1 + 2, y1 + 2 );            // 14
            LineTo(   hdc, x1 + 2, y2 - 1 );            // 15
            LineTo(   hdc, 0     , y2 - 1 );            // 16
            break;
    }
}

//===============================================================================================//
//  Description:
//      Draw the background
//
//  Parameters:
//      hdc - the device context
//
//  Returns:
//    void
//===============================================================================================//
void StaticControl:: DrawStaticBackground( HDC hdc )
{
    DWORD     mode    = GRADIENT_FILL_RECT_H;
    HRGN      hRegion = nullptr;
    HBRUSH    backgroundBrush  = nullptr;
    TRIVERTEX vertex[ 3 ];
    GRADIENT_RECT mesh[ 2 ];

    if ( hdc == nullptr ) return;

    // For closed paths, set the clip
    if ( IsClosedShape() )
    {
        BeginPath( hdc );
        DrawShape( hdc );
        EndPath( hdc );
        hRegion = PathToRegion( hdc );
        if ( hRegion )
        {
            SelectClipRgn( hdc, hRegion );
        }
        else
        {
            // Unexpected as have tested for a closed path
            PXSLogSysError( GetLastError(), L"PathToRegion failed." );
        }
    }

    if ( ( m_crGradient1 != CLR_INVALID ) && ( m_crGradient2 != CLR_INVALID ) )
    {
       // In vertical mode, this will render the gradient's end colour in the
       // middle of the window. The colour at the bottom will be the same as
       // at the top. In horizontal mode, will render a single gradient from
       // left to right.
       if ( m_bVerticalGradient )
       {
           mode = GRADIENT_FILL_RECT_V;
       }

       RECT     rect = m_rBounds;
       COLORREF cr1  = m_crGradient1, cr2 = m_crGradient2;

       vertex[ 0 ].x     = rect.left;
       vertex[ 0 ].y     = rect.top;
       vertex[ 0 ].Red   = static_cast<COLOR16>(0XFFFF & (GetRValue(cr1) << 8));
       vertex[ 0 ].Green = static_cast<COLOR16>(0XFFFF & (GetGValue(cr1) << 8));
       vertex[ 0 ].Blue  = static_cast<COLOR16>(0XFFFF & (GetBValue(cr1) << 8));
       vertex[ 0 ].Alpha = 0;

       vertex[ 1 ].x     = rect.right;
       vertex[ 1 ].y     = rect.top + ( ( rect.bottom - rect.top ) / 2 );
       vertex[ 1 ].Red   = static_cast<COLOR16>(0XFFFF & (GetRValue(cr2) << 8));
       vertex[ 1 ].Green = static_cast<COLOR16>(0XFFFF & (GetGValue(cr2) << 8));
       vertex[ 1 ].Blue  = static_cast<COLOR16>(0XFFFF & (GetBValue(cr2) << 8));
       vertex[ 1 ].Alpha = 0;

       vertex[ 2 ].x     = rect.left;
       vertex[ 2 ].y     = rect.bottom;
       vertex[ 2 ].Red   = static_cast<COLOR16>(0XFFFF & (GetRValue(cr1) << 8));
       vertex[ 2 ].Green = static_cast<COLOR16>(0XFFFF & (GetGValue(cr1) << 8));
       vertex[ 2 ].Blue  = static_cast<COLOR16>(0XFFFF & (GetBValue(cr1) << 8));
       vertex[ 2 ].Alpha = 0;

       mesh[ 0 ].UpperLeft  = 0;
       mesh[ 0 ].LowerRight = 1;
       mesh[ 1 ].UpperLeft  = 1;
       mesh[ 1 ].LowerRight = 2;
       GradientFill( hdc, vertex, 3, &mesh, 2, mode );
    }
    else if ( m_crBackground == PXS_COLOUR_WHITE )
    {
        // White is common and can use a stock object
        FillRect( hdc, &m_rBounds, (HBRUSH)GetStockObject( WHITE_BRUSH ) );
    }
    else if ( m_crBackground != CLR_INVALID )
    {
        backgroundBrush = CreateSolidBrush( m_crBackground );
        if ( backgroundBrush )
        {
            FillRect( hdc, &m_rBounds, backgroundBrush );
            DeleteObject( backgroundBrush );
        }
    }
    SelectClipRgn( hdc, nullptr );
    if ( hRegion ) DeleteObject( hRegion );
}

//===============================================================================================//
//  Description:
//      Draw the static's bitmap onto the specified device context
//
//  Parameters:
//      hdc - the device context
//
//  Returns:
//    void
//===============================================================================================//
void StaticControl::DrawStaticBitmap( HDC hdc )
{
    int x1 = m_rBounds.left;
    int y1 = m_rBounds.top;
    int x2 = m_rBounds.right;
    int y2 = m_rBounds.bottom;
    int xPos = x1 + m_nPadding;
    int yPos = y1 + m_nPadding;
    BITMAP bitmap;

    if ( ( hdc == nullptr ) || ( m_hBitmap == nullptr ) )
    {
        return;     // Nothing to do
    }

    memset( &bitmap, 0, sizeof ( bitmap ) );
    if ( GetObject( m_hBitmap, sizeof ( BITMAP ), &bitmap ) == 0 )
    {
        return;     // Perhaps an error, anyway cannot proceed
    }

    // x Position
    if ( m_uAlignmentX == PXS_CENTER_ALIGNMENT )
    {
        xPos = x1 + ( x2 - x1 - bitmap.bmWidth ) / 2;
    }
    else if ( m_uAlignmentX == PXS_RIGHT_ALIGNMENT )
    {
        xPos = x2 - bitmap.bmWidth - m_nPadding;
    }

    // y Position
    if ( m_uAlignmentY == PXS_CENTER_ALIGNMENT )
    {
        yPos = y1 + ( y2 - y1 - bitmap.bmHeight ) / 2;
    }
    else if ( m_uAlignmentY == PXS_BOTTOM_ALIGNMENT )
    {
        yPos = y2 - bitmap.bmHeight - m_nPadding;
    }
    PXSDrawTransparentBitmap( hdc, m_hBitmap, xPos, yPos, PXS_COLOUR_WHITE );
}

//===============================================================================================//
//  Description:
//      Draw the static's icon onto the specified device context
//
//  Parameters:
//      hdc - the device context
//
//  Remarks:
//      DrawIcon uses "system large icon" size.
//
//  Returns:
//    void
//===============================================================================================//
void StaticControl::DrawStaticIcon( HDC hdc )
{
    int x1   = m_rBounds.left;
    int y1   = m_rBounds.top;
    int x2   = m_rBounds.right;
    int y2   = m_rBounds.bottom;
    int xPos = x1 + m_nPadding;
    int yPos = y1 + m_nPadding;

    if ( m_rBounds.top == 408 )
    {
        m_rBounds.top =408;
    }
    if ( ( hdc == nullptr ) || ( m_hIcon == nullptr ) )
    {
       return;
    }

    // x Position
    if ( m_uAlignmentX == PXS_CENTER_ALIGNMENT )
    {
        xPos = x1 + ( x2 - x1 - GetSystemMetrics( SM_CXICON ) ) / 2;
    }
    else if ( m_uAlignmentX == PXS_RIGHT_ALIGNMENT )
    {
        xPos = x2 - GetSystemMetrics( SM_CXICON ) - m_nPadding;
    }

    // y Position
    if ( m_uAlignmentY == PXS_CENTER_ALIGNMENT )
    {
        yPos = y1 + ( y2 - y1 - GetSystemMetrics( SM_CYICON ) ) / 2;
    }
    else if ( m_uAlignmentY == PXS_BOTTOM_ALIGNMENT )
    {
        yPos = y2 - GetSystemMetrics( SM_CYICON ) - m_nPadding;
    }

    DrawIcon( hdc, xPos, yPos, m_hIcon );
}

//===============================================================================================//
//  Description:
//      Draw the static's shape
//
//  Parameters:
//      hdc - the device context
//
//  Returns:
//    void
//===============================================================================================//
void StaticControl::DrawStaticShape( HDC hdc )
{
    HPEN    dottedPen = nullptr;
    HGDIOBJ oldPen    = nullptr;

    if ( ( hdc == nullptr ) || ( m_crShapeColour == CLR_INVALID ) )
    {
        return;     // Nothing to do
    }

    // Test if need to create a dotted pen
    if ( ( m_uShape == PXS_SHAPE_TAB_DOTTED       ) ||
         ( m_uShape == PXS_SHAPE_TAB_DOTTED_RTL   ) ||
         ( m_uShape == PXS_SHAPE_RECTANGLE_DOTTED )  )
    {
        dottedPen = CreatePen( PS_DOT, 1, m_crShapeColour );
        if ( dottedPen )
        {
            oldPen = SelectObject( hdc, dottedPen );
        }
    }
    else
    {
        // Solid pen, set its colour
        oldPen = SelectObject( hdc, GetStockObject( DC_PEN ) );
        SetDCPenColor( hdc, m_crShapeColour );
    }

    // Need to reset the DC
    try
    {
        DrawShape( hdc );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }

    // Clean up
    if ( oldPen )
    {
        SelectObject( hdc, oldPen );
    }

    if ( dottedPen )
    {
       DeleteObject( dottedPen );
    }
}

//===============================================================================================//
//  Description:
//      Draw the static's text
//
//  Parameters:
//      hdc - the device context
//
//  Returns:
//    void
//===============================================================================================//
void StaticControl::DrawStaticText( HDC hdc )
{
    int      nSavedDC   = 0;
    UINT     textFormat = 0;
    RECT     bounds     = { 0, 0, 0, 0 };
    SIZE     textSize   = { 0, 0 };
    HFONT    hFont      = nullptr;
    LPCWSTR  pszText    = nullptr;
    HGDIOBJ  oldFont    = nullptr;
    COLORREF textColour = CLR_INVALID;

    pszText = m_Text.c_str();
    if ( pszText == nullptr )
    {
        return;    // Nothing to do
    }
    nSavedDC   = SaveDC( hdc );
    textColour = GetTextColour();
    SetTextColor( hdc, textColour );
    SetMapMode( hdc, MM_TEXT );
    SetBkMode( hdc, TRANSPARENT );
    if ( m_crBackground != CLR_INVALID )
    {
        SetBkColor( hdc, m_crBackground );
    }

    // Set bounds, adjust for frame shapes
    bounds.left   = m_rBounds.left   + m_nPadding;
    bounds.top    = m_rBounds.top    + m_nPadding;
    bounds.right  = m_rBounds.right  - m_nPadding;
    bounds.bottom = m_rBounds.bottom - m_nPadding;
    if ( m_uShape == PXS_SHAPE_FRAME )
    {
        bounds.left   = m_rBounds.left + 10;
        bounds.bottom = m_rBounds.top  + 20;
    }
    else if ( m_uShape == PXS_SHAPE_FRAME_RTL )
    {
        GetTextExtentPoint32( hdc, pszText, lstrlen( pszText ), &textSize );
        bounds.left   = m_rBounds.right - 10 - textSize.cx;
        bounds.bottom = m_rBounds.top + 20;
    }

    // Apply the font
    hFont = m_Font.GetHandle();
    if ( hFont )
    {
        oldFont = SelectObject( hdc, hFont );
    }
    textFormat = GetDrawTextFormat();
    DrawText( hdc, pszText, lstrlen( pszText ), &bounds, textFormat );

    // Reset
    if ( oldFont )
    {
        SelectObject( hdc, oldFont );
    }
    RestoreDC( hdc, nSavedDC );
}

//===============================================================================================//
//  Description:
//      Make the format text value for DrawText
//
//  Parameters:
//      none
//
//  Returns:
//      UINT
//===============================================================================================//
UINT StaticControl::GetDrawTextFormat()
{
    UINT textFormat = DT_NOPREFIX | DT_EXPANDTABS | DT_END_ELLIPSIS | DT_TOP | DT_LEFT;

    if ( m_bSingleLine )
    {
        textFormat |= DT_SINGLELINE;
        if ( m_uAlignmentY == PXS_CENTER_ALIGNMENT )
        {
            textFormat |= DT_VCENTER;
        }
        if ( m_uAlignmentY == PXS_BOTTOM_ALIGNMENT )
        {
            textFormat |= DT_BOTTOM;
        }
    }
    else
    {
        textFormat |= DT_WORDBREAK;
    }

    if ( m_uAlignmentX == PXS_CENTER_ALIGNMENT )
    {
        textFormat |= DT_CENTER;
    }

    if ( m_uAlignmentX == PXS_RIGHT_ALIGNMENT  )
    {
        textFormat |= DT_RIGHT;
    }

    return textFormat;
}

//===============================================================================================//
//  Description:
//      Determine the text colour for this static
//
//  Parameters:
//      none
//
//  Returns:
//      COLORREF
//===============================================================================================//
COLORREF StaticControl::GetTextColour()
{
    COLORREF textColour = m_crForeground;

    if ( m_bEnabled )
    {
        textColour = m_crForeground;
        if ( textColour == CLR_INVALID )
        {
            if ( m_bHyperlink )
            {
                textColour = GetSysColor( COLOR_HOTLIGHT );
            }
            else
            {
                textColour = GetSysColor( COLOR_WINDOWTEXT );
            }
        }
    }
    else
    {
        textColour = GetSysColor( COLOR_GRAYTEXT );
    }

    return textColour;
}

//===============================================================================================//
//  Description:
//      Determine if the static's shape is a closed path
//
//  Parameters:
//      none
//
//  Returns:
//      true if closed, otherwise false
//===============================================================================================//
bool StaticControl::IsClosedShape()
{
    DWORD ClosedShapes[] = { PXS_SHAPE_ARROW_DOWN,
                             PXS_SHAPE_ARROW_UP,
                             PXS_SHAPE_CROSS,
                             PXS_SHAPE_CIRCLE,
                             PXS_SHAPE_DIAMOND,
                             PXS_SHAPE_FIVE_SIDES,
                             PXS_SHAPE_HOME,
                             PXS_SHAPE_RECTANGLE,
                             PXS_SHAPE_RECTANGLE_DOTTED,
                             PXS_SHAPE_RECTANGLE_HOLLOW,
                             PXS_SHAPE_REFRESH,
                             PXS_SHAPE_TAB,
                             PXS_SHAPE_TAB_DOTTED,
                             PXS_SHAPE_TAB_DOTTED_RTL,
                             PXS_SHAPE_TAB_RTL,
                             PXS_SHAPE_TRIANGLE_LEFT,
                             PXS_SHAPE_TRIANGLE_RIGHT };

    return PXSIsInUInt32Array( m_uShape,
                               ClosedShapes, ARRAYSIZE( ClosedShapes ) );
}
