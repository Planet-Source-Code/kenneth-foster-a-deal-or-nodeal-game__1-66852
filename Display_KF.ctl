VERSION 5.00
Begin VB.UserControl Display 
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   870
   ScaleWidth      =   4800
   Begin VB.PictureBox Display 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   0
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   307
      TabIndex        =   0
      Top             =   0
      Width           =   4665
   End
End
Attribute VB_Name = "Display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum BorderStyles
None = 0
Fixed = 1
End Enum

Const m_def_BackgroundColor = vbBlue
Const m_def_TextColor = vbYellow
Const m_def_GridColor = &HC00000
Const m_def_Border = None
Const m_def_FontBold = True

Dim m_BackgroundColor As OLE_COLOR
Dim m_TextColor As OLE_COLOR
Dim m_GridColor As OLE_COLOR
Dim m_Text As String
Dim m_Font As StdFont
Dim m_Border As BorderStyles
Dim m_FontBold As Boolean

Private Sub UserControl_Initialize()
   Set Font = UserControl.Font
   TextColor = m_def_TextColor
   BackgroundColor = m_def_BackgroundColor
   Border = Fixed
   GridColor = m_def_GridColor
End Sub

Private Sub UserControl_InitProperties()
    m_Text = Extender.Name
    m_FontBold = m_def_FontBold
    m_Border = m_def_Border
    m_BackgroundColor = m_def_BackgroundColor
    m_GridColor = m_def_GridColor
End Sub

Public Property Get BackgroundColor() As OLE_COLOR
   BackgroundColor = m_BackgroundColor
End Property

Public Property Let BackgroundColor(NewBackgroundColor As OLE_COLOR)
   m_BackgroundColor = NewBackgroundColor
   PropertyChanged "BackgroundColor"
   Display.BackColor = m_BackgroundColor
   Draw
End Property

Public Property Get Border() As BorderStyles
   Border = m_Border
End Property

Public Property Let Border(NewBorder As BorderStyles)
   m_Border = NewBorder
   PropertyChanged "Border"
   Display.BorderStyle = m_Border
   Draw
End Property

Public Property Get FontBold() As Boolean
   FontBold = m_FontBold
End Property

Public Property Let FontBold(NewFontBold As Boolean)
   m_FontBold = NewFontBold
   PropertyChanged "FontBold"
    If FontBold = True Then
      Font.Bold = True
   Else
      Font.Bold = False
   End If
   Draw
End Property
Public Property Get TextColor() As OLE_COLOR
   TextColor = m_TextColor
End Property

Public Property Let TextColor(NewTextColor As OLE_COLOR)
   m_TextColor = NewTextColor
   PropertyChanged "TextColor"
   Draw
End Property

Public Property Get GridColor() As OLE_COLOR
   GridColor = m_GridColor
End Property

Public Property Let GridColor(NewGridColor As OLE_COLOR)
   m_GridColor = NewGridColor
   PropertyChanged "GridColor"
   Draw
End Property

Public Property Get Text() As String
   Text = m_Text
End Property

Public Property Let Text(NewText As String)
   m_Text = NewText
   PropertyChanged "Text"
   Draw
End Property

Public Property Get Font() As StdFont
   Set Font = m_Font
End Property

Public Property Set Font(ByRef NewFont As StdFont)
   Set m_Font = NewFont
   PropertyChanged "Font"
   Display.Font = NewFont
   If Font.Bold = True Then
      FontBold = True
   Else
      FontBold = False
   End If
   Draw
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   BackgroundColor = PropBag.ReadProperty("BackgroundColor", m_def_BackgroundColor)
   Border = PropBag.ReadProperty("Border", m_def_Border)
   TextColor = PropBag.ReadProperty("TextColor", m_def_TextColor)
   GridColor = PropBag.ReadProperty("GridColor", m_def_GridColor)
   Text = PropBag.ReadProperty("Text", Extender.Name)
   Set Font = PropBag.ReadProperty("Font", UserControl.Font)
   FontBold = PropBag.ReadProperty("FontBold", m_def_FontBold)
End Sub

Private Sub UserControl_Resize()
    Display.Width = UserControl.Width
    Display.Height = UserControl.Height
    Draw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
   Call .WriteProperty("BackgroundColor", m_BackgroundColor, m_def_BackgroundColor)
   Call .WriteProperty("Border", m_Border, m_def_Border)
   Call .WriteProperty("TextColor", m_TextColor, m_def_TextColor)
   Call .WriteProperty("GridColor", m_GridColor, m_def_GridColor)
   Call .WriteProperty("Text", m_Text, Extender.Name)
   Call .WriteProperty("Font", m_Font, UserControl.Font)
   Call .WriteProperty("FontBold", m_FontBold, m_def_FontBold)
   End With
End Sub

Private Sub Draw()
Dim x As Integer
Dim y As Integer

Display.Cls                                                          ' clear for next text message
Display.CurrentX = 2                                             ' positions where the text is displayed
Display.CurrentY = -4
Display.Font.Bold = Font.Bold
Display.FontSize = Font.Size
Display.ForeColor = TextColor                                ' display text color
Display.Print Text
' Draw the grid lines
For x = 0 To Display.Height Step 2
   Display.Line (0, x)-(Display.Width, x), GridColor
Next x
For y = 0 To Display.Width Step 2
   Display.Line (y, 0)-(y, Display.Height), GridColor
Next y
End Sub
