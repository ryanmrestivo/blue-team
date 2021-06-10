///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Tree View Item Class Implementation
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
#include "PxsBase/Header Files/TreeViewItem.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Exception.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor. Keep values in sync with the Reset method
TreeViewItem::TreeViewItem()
             :m_bExpanded( false ),
              m_bNode( false ),
              m_bSelected( false ),
              m_uIndent( 0 ),
              m_hBitmap( nullptr ),
              m_Label(),
              m_StringData()
{
}

// Copy constructor
TreeViewItem::TreeViewItem( const TreeViewItem& oTreeViewItem )
             :m_bExpanded( false ),
              m_bNode( false ),
              m_bSelected( false ),
              m_uIndent( 0 ),
              m_hBitmap( nullptr ),
              m_Label(),
              m_StringData()
{
    *this = oTreeViewItem;
}

// Destructor
TreeViewItem::~TreeViewItem()
{
    // Clean up
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
TreeViewItem& TreeViewItem::operator= ( const TreeViewItem& oTreeViewItem )
{
    if ( this == &oTreeViewItem ) return *this;

    m_bExpanded  = oTreeViewItem.m_bExpanded;
    m_bNode      = oTreeViewItem.m_bNode;
    m_bSelected  = oTreeViewItem.m_bSelected;
    m_uIndent    = oTreeViewItem.m_uIndent;
    m_Label      = oTreeViewItem.m_Label;
    m_StringData = oTreeViewItem.m_StringData;

    // Free any existing bitmap then copy
    if ( m_hBitmap )
    {
        DeleteObject( m_hBitmap );
        m_hBitmap = nullptr;
    }

    if ( oTreeViewItem.m_hBitmap )
    {
       m_hBitmap = (HBITMAP)CopyImage( oTreeViewItem.m_hBitmap, IMAGE_BITMAP, 0, 0, 0 );
    }

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the bitmap for this tree view item
//
//  Parameters:
//      None
//
//  Returns:
//      Handle to the bitmap
//===============================================================================================//
HBITMAP TreeViewItem::GetBitmap() const
{
    return m_hBitmap;
}

//===============================================================================================//
//  Description:
//      Get the indent level for this tree view item
//
//  Parameters:
//      None
//
//  Returns:
//      BYTE representing the indent level
//===============================================================================================//
BYTE TreeViewItem::GetIndent() const
{
    return m_uIndent;
}

//===============================================================================================//
//  Description:
//      Get the labe; for this tree view item
//
//  Parameters:
//      None
//
//  Returns:
//      Constant string pointer
//===============================================================================================//
LPCWSTR TreeViewItem::GetLabel() const
{
    return m_Label.c_str();
}

//===============================================================================================//
//  Description:
//      Get the string data associated with this tree view item
//
//  Parameters:
//      None
//
//  Returns:
//      Constant string pointer
//===============================================================================================//
LPCWSTR TreeViewItem::GetStringData() const
{
    return m_StringData.c_str();
}

//===============================================================================================//
//  Description:
//      Get if the this tree view item is expanded, applies to nodes only.
//
//  Parameters:
//      None
//
//  Returns:
//      true if node expanded, else false
//===============================================================================================//
bool TreeViewItem::IsExpanded() const
{
    return m_bExpanded;
}

//===============================================================================================//
//  Description:
//      Get if the this tree view item is a node.
//
//  Parameters:
//      None
//
//  Returns:
//      true if this item is a node, else it is a leaf
//===============================================================================================//
bool TreeViewItem::IsNode() const
{
    return m_bNode;
}

//===============================================================================================//
//  Description:
//      Get if the this tree view item is selected.
//
//  Parameters:
//      None
//
//  Returns:
//      true if this item is selected
//===============================================================================================//
bool TreeViewItem::IsSelected() const
{
    return m_bSelected;
}

//===============================================================================================//
//  Description:
//      Reset the values of this class
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void TreeViewItem::Reset()
{
    m_bExpanded  = false;
    m_bNode      = false;
    m_bSelected  = false;
    m_uIndent    = 0;
    m_Label      = PXS_STRING_EMPTY;
    m_StringData = PXS_STRING_EMPTY;

    // Delete the bitmap
    if ( m_hBitmap )
    {
        DeleteObject( m_hBitmap );
        m_hBitmap = nullptr;
    }
}

//===============================================================================================//
//  Description:
//      Set the bitmap for this tree view item
//
//  Parameters:
//      pszBitmapName - pointer to resource ID for the bitmap
//      transparent   - nominated transparent colour
//      background    - background colour onto which the bitmap will be drawn
//
//  Returns:
//      void
//===============================================================================================//
void TreeViewItem::SetBitmap( LPCWSTR pszBitmapName, COLORREF transparent, COLORREF background )
{
    // Release any existing bitmap
    if ( m_hBitmap )
    {
        DeleteObject( m_hBitmap );
        m_hBitmap = nullptr;
    }

    if ( pszBitmapName == nullptr )
    {
        return;     // Nothing to do
    }

    m_hBitmap = LoadBitmap( GetModuleHandle( nullptr ), pszBitmapName );
    if ( m_hBitmap )
    {
        // Replace colour to make the bitmap transparent
        if ( ( transparent != CLR_INVALID ) && ( background != CLR_INVALID ) )
        {
            PXSReplaceBitmapColour( m_hBitmap, transparent, background );
        }
    }
}

//===============================================================================================//
//  Description:
//      Set if the this tree view item is expanded, applies to nodes only.
//
//  Parameters:
//      expanded - true if node expanded, else false
//
//  Returns:
//      void
//===============================================================================================//
void TreeViewItem::SetExpanded( bool expanded )
{
    m_bExpanded = expanded;
}

//===============================================================================================//
//  Description:
//      Set the indent level for this tree view item
//
//  Parameters:
//      indent - the indent level
//
//  Returns:
//      void
//===============================================================================================//
void TreeViewItem::SetIndent( BYTE indent )
{
    m_uIndent = indent;
}

//===============================================================================================//
//  Description:
//      Set the label for this tree view item
//
//  Parameters:
//      Label - the caption string
//
//  Returns:
//      void
//===============================================================================================//
void TreeViewItem::SetLabel( const String& Label )
{
    m_Label = Label;
}

//===============================================================================================//
//  Description:
//      Set if the this tree view item is a node.
//
//  Parameters:
//      node - true if this item is a node, else it is a leaf
//
//  Returns:
//      void
//===============================================================================================//
void TreeViewItem::SetNode( bool node )
{
    m_bNode = node;
}

//===============================================================================================//
//  Description:
//      Set the data string for this tree view item
//
//  Parameters:
//      StringData - the data string
//
//  Returns:
//      void
//===============================================================================================//
void TreeViewItem::SetStringData( const String& StringData )
{
    m_StringData = StringData;
}

//===============================================================================================//
//  Description:
//      Set if the this tree view item is selected.
//
//  Parameters:
//      selected - true if this item is selected
//
//  Returns:
//      void
//===============================================================================================//
void TreeViewItem::SetSelected( bool selected )
{
    m_bSelected = selected;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
