VERSION 5.00
Begin VB.UserControl OutLineText 
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2130
   ScaleHeight     =   675
   ScaleWidth      =   2130
   Begin VB.Label Letter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   525
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   330
   End
   Begin VB.Label Letter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   330
   End
   Begin VB.Label Letter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Letter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   3
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   330
   End
   Begin VB.Label Letter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   4
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "OutLineText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const m_def_Caption = "OutlineText"
Const m_def_OutlineCol = vbBlack
Const m_def_TextCol = vbYellow
Const m_def_BackCol = &H8000000F

Dim m_Caption As String
Dim m_OutlineCol As OLE_COLOR
Dim m_TextCol As OLE_COLOR
Dim m_BackCol As OLE_COLOR

Private Sub UserControl_Initialize()
Dim x As Integer
For x = 0 To 4
Letter(x).Caption = m_def_Caption
Next x
UserControl.BackColor = m_def_BackCol
End Sub

Private Sub UserControl_InitProperties()
m_Caption = m_def_Caption
m_OutlineCol = m_def_OutlineCol
m_TextCol = m_def_TextCol
m_BackCol = m_def_BackCol
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
m_OutlineCol = PropBag.ReadProperty("OutlineCol", m_def_OutlineCol)
m_TextCol = PropBag.ReadProperty("TextCol", m_def_TextCol)
m_BackCol = PropBag.ReadProperty("BackCol", m_def_BackCol)
Set Font = PropBag.ReadProperty("Font", Ambient.Font)

Letter(0).ForeColor = TextCol
Letter(0).Caption = Caption
Letter(1).ForeColor = OutlineCol
Letter(1).Caption = Caption
Letter(2).ForeColor = OutlineCol
Letter(2).Caption = Caption
Letter(3).ForeColor = OutlineCol
Letter(3).Caption = Caption
Letter(4).ForeColor = OutlineCol
Letter(4).Caption = Caption
UserControl.BackColor = m_def_BackCol
UserControl_Resize
End Sub

Private Sub UserControl_Resize()

UserControl.Width = Letter(0).Width + 100
UserControl.Height = Letter(0).Height + 50
UserControl.BackColor = m_BackCol

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Caption", m_Caption, m_def_Caption
PropBag.WriteProperty "OutlineCol", m_OutlineCol, m_def_OutlineCol
PropBag.WriteProperty "TextCol", m_TextCol, m_def_TextCol
PropBag.WriteProperty "BackCol", m_BackCol, m_def_BackCol
PropBag.WriteProperty "Font", Font, Ambient.Font

End Sub
Public Property Get Caption() As String
Caption = m_Caption
UserControl_Resize
End Property
Public Property Let Caption(ByVal NewCaption As String)
Dim x As Integer

m_Caption = NewCaption
For x = 0 To 4
   Letter(x).Caption = m_Caption
Next x
PropertyChanged "Caption"

End Property
Public Property Get OutlineCol() As OLE_COLOR
OutlineCol = m_OutlineCol
End Property
Public Property Let OutlineCol(ByVal NewOutlineCol As OLE_COLOR)
Dim x As Integer

m_OutlineCol = NewOutlineCol
For x = 1 To 4
   Letter(x).ForeColor = m_OutlineCol
Next x
PropertyChanged "OutlineCol"

End Property

Public Property Get TextCol() As OLE_COLOR
TextCol = m_TextCol
End Property

Public Property Let TextCol(ByVal NewTextCol As OLE_COLOR)
m_TextCol = NewTextCol
Letter(0).ForeColor = m_TextCol
PropertyChanged "TextCol"
End Property

Public Property Get BackCol() As OLE_COLOR
BackCol = m_BackCol
End Property
Public Property Let BackCol(ByVal NewBackCol As OLE_COLOR)
m_BackCol = NewBackCol
UserControl.BackColor = m_BackCol
PropertyChanged "BackCol"
End Property
Public Property Get Font() As Font
Dim x As Integer
For x = 0 To 4
  Set Font = Letter(x).Font
Next x
End Property
Public Property Set Font(ByVal NewFont As Font)
Dim x As Integer

For x = 0 To 4
  Set Letter(x).Font = NewFont
Next x
End Property
