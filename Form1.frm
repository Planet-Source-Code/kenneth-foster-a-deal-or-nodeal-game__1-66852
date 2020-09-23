VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8505
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Offer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   495
      Top             =   5865
   End
   Begin VB.TextBox txtPrevOffer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   3225
      MultiLine       =   -1  'True
      TabIndex        =   64
      Text            =   "Form1.frx":0000
      Top             =   2415
      Visible         =   0   'False
      Width           =   1695
   End
   Begin Project1.OutLineText OT4 
      Height          =   600
      Left            =   375
      TabIndex        =   63
      Top             =   5325
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   1058
      Caption         =   "0"
      TextCol         =   16711680
      BackCol         =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.ccXPButton cmdExit 
      Height          =   435
      Left            =   7755
      TabIndex        =   62
      Top             =   5790
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   767
      Caption         =   "Exit"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   990
      ScaleHeight     =   1185
      ScaleWidth      =   6675
      TabIndex        =   4
      Top             =   5010
      Width           =   6705
      Begin Project1.Display Display 
         Height          =   615
         Left            =   30
         TabIndex        =   66
         Top             =   30
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   1085
         Border          =   1
         Text            =   "       Deal or No Deal"
      End
      Begin Project1.ccXPButton cmdStart 
         Height          =   465
         Left            =   2655
         TabIndex        =   7
         Top             =   690
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   820
         Caption         =   "START"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.ccXPButton cmdNoDeal 
         Height          =   465
         Left            =   4200
         TabIndex        =   6
         Top             =   690
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   820
         Caption         =   "NO DEAL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.ccXPButton cmdDeal 
         Height          =   465
         Left            =   1170
         TabIndex        =   5
         Top             =   690
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   820
         Caption         =   "DEAL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer DelayTimer 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   75
      Top             =   5865
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   4740
      Left            =   135
      ScaleHeight     =   4710
      ScaleWidth      =   2475
      TabIndex        =   2
      Top             =   195
      Width           =   2505
      Begin Project1.OutLineText OT3 
         Height          =   405
         Left            =   1080
         TabIndex        =   61
         Top             =   330
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   714
         Caption         =   "NO DEAL"
         TextCol         =   49152
         BackCol         =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.OutLineText OT2 
         Height          =   405
         Left            =   795
         TabIndex        =   60
         Top             =   150
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   714
         Caption         =   "or"
         TextCol         =   12583104
         BackCol         =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.OutLineText OT1 
         Height          =   405
         Left            =   30
         TabIndex        =   59
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   714
         Caption         =   "DEAL"
         BackCol         =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape Shape2 
         Height          =   750
         Left            =   15
         Top             =   15
         Width           =   2445
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   25
         Left            =   1260
         TabIndex        =   57
         Top             =   4395
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   24
         Left            =   1260
         TabIndex        =   56
         Top             =   4095
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   23
         Left            =   1260
         TabIndex        =   55
         Top             =   3795
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   22
         Left            =   1260
         TabIndex        =   54
         Top             =   3495
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   21
         Left            =   1260
         TabIndex        =   53
         Top             =   3195
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   20
         Left            =   1260
         TabIndex        =   52
         Top             =   2895
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   19
         Left            =   1260
         TabIndex        =   51
         Top             =   2595
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   18
         Left            =   1260
         TabIndex        =   50
         Top             =   2295
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   17
         Left            =   1260
         TabIndex        =   49
         Top             =   1995
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   16
         Left            =   1260
         TabIndex        =   48
         Top             =   1695
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   15
         Left            =   1260
         TabIndex        =   47
         Top             =   1395
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   14
         Left            =   1260
         TabIndex        =   46
         Top             =   1095
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   13
         Left            =   1260
         TabIndex        =   45
         Top             =   795
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   12
         Left            =   30
         TabIndex        =   44
         Top             =   4395
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   11
         Left            =   30
         TabIndex        =   43
         Top             =   4095
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   10
         Left            =   30
         TabIndex        =   42
         Top             =   3795
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   30
         TabIndex        =   41
         Top             =   3495
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   30
         TabIndex        =   40
         Top             =   3195
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   30
         TabIndex        =   39
         Top             =   2895
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   30
         TabIndex        =   38
         Top             =   2595
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   30
         TabIndex        =   37
         Top             =   2295
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   30
         TabIndex        =   36
         Top             =   1995
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   30
         TabIndex        =   35
         Top             =   1695
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   30
         TabIndex        =   34
         Top             =   1395
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   30
         TabIndex        =   33
         Top             =   1095
         Width           =   1200
      End
      Begin VB.Label lblAmounts 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   30
         TabIndex        =   3
         Top             =   795
         Width           =   1200
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   4740
      Left            =   2700
      ScaleHeight     =   4710
      ScaleWidth      =   5610
      TabIndex        =   0
      Top             =   195
      Width           =   5640
      Begin VB.Line Line16 
         X1              =   4545
         X2              =   4545
         Y1              =   2895
         Y2              =   3630
      End
      Begin VB.Line Line15 
         X1              =   3645
         X2              =   3645
         Y1              =   2910
         Y2              =   3660
      End
      Begin VB.Line Line14 
         X1              =   2745
         X2              =   2745
         Y1              =   2910
         Y2              =   3660
      End
      Begin VB.Line Line13 
         X1              =   1845
         X2              =   1845
         Y1              =   2910
         Y2              =   3645
      End
      Begin VB.Line Line12 
         X1              =   945
         X2              =   945
         Y1              =   2910
         Y2              =   3645
      End
      Begin VB.Line Line11 
         X1              =   4995
         X2              =   4995
         Y1              =   0
         Y2              =   2895
      End
      Begin VB.Line Line10 
         X1              =   495
         X2              =   495
         Y1              =   -15
         Y2              =   2910
      End
      Begin VB.Line Line9 
         X1              =   4095
         X2              =   4095
         Y1              =   -45
         Y2              =   2910
      End
      Begin VB.Line Line8 
         X1              =   3195
         X2              =   3195
         Y1              =   -60
         Y2              =   2910
      End
      Begin VB.Line Line7 
         X1              =   2295
         X2              =   2295
         Y1              =   0
         Y2              =   2910
      End
      Begin VB.Line Line6 
         X1              =   1395
         X2              =   1395
         Y1              =   0
         Y2              =   2910
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   -15
         X2              =   5610
         Y1              =   3645
         Y2              =   3645
      End
      Begin VB.Line Line4 
         X1              =   -45
         X2              =   5625
         Y1              =   2895
         Y2              =   2895
      End
      Begin VB.Line Line3 
         X1              =   495
         X2              =   5010
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Line Line2 
         X1              =   495
         X2              =   5010
         Y1              =   1485
         Y2              =   1485
      End
      Begin VB.Line Line1 
         X1              =   495
         X2              =   4995
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Label lblSelectedNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2775
         TabIndex        =   58
         Top             =   4110
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Image imgSelectedCase 
         Height          =   660
         Left            =   2400
         Picture         =   "Form1.frx":0012
         Top             =   3840
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   25
         Left            =   4890
         TabIndex        =   32
         Top             =   3210
         Width           =   255
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   24
         Left            =   3990
         TabIndex        =   31
         Top             =   3210
         Width           =   240
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   23
         Left            =   3090
         TabIndex        =   30
         Top             =   3210
         Width           =   240
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   22
         Left            =   2205
         TabIndex        =   29
         Top             =   3210
         Width           =   240
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   21
         Left            =   1305
         TabIndex        =   28
         Top             =   3210
         Width           =   240
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   20
         Left            =   405
         TabIndex        =   27
         Top             =   3210
         Width           =   240
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   19
         Left            =   4440
         TabIndex        =   26
         Top             =   2505
         Width           =   255
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   18
         Left            =   3540
         TabIndex        =   25
         Top             =   2505
         Width           =   240
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   17
         Left            =   2625
         TabIndex        =   24
         Top             =   2490
         Width           =   255
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   16
         Left            =   1725
         TabIndex        =   23
         Top             =   2490
         Width           =   255
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   15
         Left            =   840
         TabIndex        =   22
         Top             =   2490
         Width           =   240
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   25
         Left            =   4590
         Picture         =   "Form1.frx":0A61
         Top             =   2940
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   24
         Left            =   3690
         Picture         =   "Form1.frx":14B0
         Top             =   2940
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   23
         Left            =   2790
         Picture         =   "Form1.frx":1EFF
         Top             =   2940
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   22
         Left            =   1890
         Picture         =   "Form1.frx":294E
         Top             =   2940
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   21
         Left            =   990
         Picture         =   "Form1.frx":339D
         Top             =   2940
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   20
         Left            =   90
         Picture         =   "Form1.frx":3DEC
         Top             =   2940
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   19
         Left            =   4140
         Picture         =   "Form1.frx":483B
         Top             =   2220
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   18
         Left            =   3240
         Picture         =   "Form1.frx":528A
         Top             =   2220
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   17
         Left            =   2340
         Picture         =   "Form1.frx":5CD9
         Top             =   2220
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   16
         Left            =   1440
         Picture         =   "Form1.frx":6728
         Top             =   2220
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   15
         Left            =   540
         Picture         =   "Form1.frx":7177
         Top             =   2220
         Width           =   825
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   14
         Left            =   4455
         TabIndex        =   21
         Top             =   1785
         Width           =   240
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   13
         Left            =   3555
         TabIndex        =   20
         Top             =   1785
         Width           =   240
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   12
         Left            =   2655
         TabIndex        =   19
         Top             =   1785
         Width           =   240
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   1740
         TabIndex        =   18
         Top             =   1785
         Width           =   255
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   10
         Left            =   840
         TabIndex        =   17
         Top             =   1785
         Width           =   255
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   9
         Left            =   4440
         TabIndex        =   16
         Top             =   1080
         Width           =   255
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   14
         Left            =   4140
         Picture         =   "Form1.frx":7BC6
         Top             =   1515
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   13
         Left            =   3240
         Picture         =   "Form1.frx":8615
         Top             =   1515
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   12
         Left            =   2340
         Picture         =   "Form1.frx":9064
         Top             =   1515
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   11
         Left            =   1440
         Picture         =   "Form1.frx":9AB3
         Top             =   1515
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   10
         Left            =   540
         Picture         =   "Form1.frx":A502
         Top             =   1515
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   9
         Left            =   4155
         Picture         =   "Form1.frx":AF51
         Top             =   810
         Width           =   825
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   3600
         TabIndex        =   15
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   2700
         TabIndex        =   14
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   1800
         TabIndex        =   13
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   900
         TabIndex        =   12
         Top             =   1065
         Width           =   135
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   4500
         TabIndex        =   11
         Top             =   390
         Width           =   135
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   3600
         TabIndex        =   10
         Top             =   390
         Width           =   120
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   2700
         TabIndex        =   9
         Top             =   375
         Width           =   120
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   1800
         TabIndex        =   8
         Top             =   375
         Width           =   120
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   8
         Left            =   3240
         Picture         =   "Form1.frx":B9A0
         Top             =   810
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   7
         Left            =   2340
         Picture         =   "Form1.frx":C3EF
         Top             =   810
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   6
         Left            =   1440
         Picture         =   "Form1.frx":CE3E
         Top             =   810
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   5
         Left            =   540
         Picture         =   "Form1.frx":D88D
         Top             =   810
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   4
         Left            =   4140
         Picture         =   "Form1.frx":E2DC
         Top             =   105
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   3
         Left            =   3240
         Picture         =   "Form1.frx":ED2B
         Top             =   105
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   2
         Left            =   2340
         Picture         =   "Form1.frx":F77A
         Top             =   105
         Width           =   825
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   1
         Left            =   1440
         Picture         =   "Form1.frx":101C9
         Top             =   105
         Width           =   825
      End
      Begin VB.Shape Shape1 
         Height          =   765
         Left            =   2250
         Top             =   3795
         Width           =   1125
      End
      Begin VB.Label lblCaseNum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   1
         Top             =   360
         Width           =   135
      End
      Begin VB.Image imgBriefCase 
         Height          =   660
         Index           =   0
         Left            =   540
         Picture         =   "Form1.frx":10C18
         Top             =   105
         Width           =   825
      End
   End
   Begin Project1.OutLineText OutLineText1 
      Height          =   315
      Left            =   45
      TabIndex        =   65
      Top             =   5070
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      Caption         =   "Round"
      TextCol         =   16711680
      BackCol         =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      X1              =   15
      X2              =   15
      Y1              =   15
      Y2              =   6375
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00FF8080&
      X1              =   8430
      X2              =   8550
      Y1              =   6330
      Y2              =   6465
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00C00000&
      X1              =   60
      X2              =   -90
      Y1              =   60
      Y2              =   -90
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00C00000&
      BorderWidth     =   6
      X1              =   45
      X2              =   8535
      Y1              =   6390
      Y2              =   6390
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00C00000&
      BorderWidth     =   6
      X1              =   8490
      X2              =   8490
      Y1              =   30
      Y2              =   6435
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      X1              =   0
      X2              =   8490
      Y1              =   15
      Y2              =   15
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SelectingMyCase As Boolean
Dim NumOfCasesLeft As Integer
Dim sum As Long                               ' the bankers offer amount
Dim MyCase As Integer                       ' users own case
Dim CasePicked As Integer                  ' case selected by user, not his own
Dim pickcount As Integer                     ' keeps track of number of cases left to pick
Dim bCount As Integer                         ' keeps track of number of amounts still showing
Dim PriceArray(26) As String                'holds the money amounts
Dim PM(26)                                         'randomized number list used for prize money index
Dim PrevArray(9) As String                    'stores previous offers
Dim pc As Integer                                  ' index for above array
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Private Sub Form_Load()
   Picture1.Enabled = False                'lock gameboard until ready to play
   Display.Text = "   Press Start to Begin..."
End Sub

Private Sub LoadAmtArray()
   Dim x As Integer
   
   PriceArray(0) = ".01"
   PriceArray(1) = "1"
   PriceArray(2) = "5"
   PriceArray(3) = "10"
   PriceArray(4) = "25"
   PriceArray(5) = "50"
   PriceArray(6) = "75"
   PriceArray(7) = "100"
   PriceArray(8) = "200"
   PriceArray(9) = "300"
   PriceArray(10) = "400"
   PriceArray(11) = "500"
   PriceArray(12) = "750"
   PriceArray(13) = "1000"
   PriceArray(14) = "5000"
   PriceArray(15) = "10000"
   PriceArray(16) = "25000"
   PriceArray(17) = "50000"
   PriceArray(18) = "75000"
   PriceArray(19) = "100000"
   PriceArray(20) = "200000"
   PriceArray(21) = "300000"
   PriceArray(22) = "400000"
   PriceArray(23) = "500000"
   PriceArray(24) = "750000"
   PriceArray(25) = "1000000"
   'load the prize amounts into the windows
   For x = 0 To 25
      If PriceArray(x) >= 100 Then
          lblAmounts(x).Caption = Format(PriceArray(x), "#,##0")
      Else
          lblAmounts(x).Caption = Format(PriceArray(x), "#,##0.00")
      End If
   Next x
End Sub

Private Sub LoadPrizeMoneyArray()
   Dim MaxNumber As Integer
   Dim seq As Integer
   Dim MainLoop As Integer
   Dim ChosenNumber As Integer
   Dim A(26)
   Dim x As Integer
   
   MaxNumber = 25                                            'Sets the maximum number to pick
   
   For seq = 0 To MaxNumber
      A(seq) = seq
   Next seq
   'Main Loop (mix em all up)
   Randomize
   For MainLoop = MaxNumber To 0 Step -1
      ChosenNumber = Int(MainLoop * Rnd)
      PM(MaxNumber - MainLoop) = A(ChosenNumber)
      A(ChosenNumber) = A(MainLoop)
   Next MainLoop
End Sub

Private Sub cmdStart_Click()
Dim x As Integer

   MyCase = 0
   CasePicked = 0
   OT4.Caption = "1"                                            ' set Round number to beginning
   pickcount = 6
   pc = 0
   txtPrevOffer.Visible = False
   txtPrevOffer.Text = "Previous Offers" & vbCrLf
   NumOfCasesLeft = 26
   imgSelectedCase.Visible = False                       'hide bottom briefcase from previous game
   lblSelectedNum.Visible = False
  Display.Text = "         Pick a Case"
   For x = 0 To 25                                                  'load all briefcases
      imgBriefCase(x).Visible = True
      lblCaseNum(x).Visible = True
      lblAmounts(x).Visible = True
      lblCaseNum(x).Caption = x + 1
   Next x
   'set some values
   LoadAmtArray
   LoadPrizeMoneyArray
   SelectingMyCase = True
   Picture1.Enabled = True                                        ' allow gameplay
   cmdStart.Visible = False
   Picture1.BackColor = &HFF0000                            'change background color back to blue
   Picture2.BackColor = &HFF0000
   Picture4.BackColor = &HFF0000
   OT1.BackCol = &HFF0000
   OT2.BackCol = &HFF0000
   OT3.BackCol = &HFF0000
End Sub

Private Sub cmdDeal_Click()
Dim x As Integer
   OT4.Visible = False                                                            ' hide Round number
  Display.Text = "      Deal is " & Format(sum, "$#,##0") & "   "
   lblSelectedNum.Caption = lblAmounts(PM(MyCase)).Caption
   For x = 0 To 25
      If imgBriefCase(x).Visible = True Then                                 ' show amounts for remaining briefcases
         If lblAmounts(PM(x)).Caption > 50 Then
            lblCaseNum(x).Caption = Format(lblAmounts(PM(x)).Caption, "#,##0")
         Else
            lblCaseNum(x).Caption = Format(lblAmounts(PM(x)).Caption, "#,##0.00")
         End If
      End If
   Next x
   cmdStart.Visible = True
   cmdDeal.Visible = False
   cmdNoDeal.Visible = False
   txtPrevOffer.Visible = False
End Sub

Private Sub cmdNoDeal_Click()
   Dim x As Integer
   
    cmdDeal.Visible = False
    cmdNoDeal.Visible = False
    Offer.Enabled = False                                                       ' in case you press no deal button before timer is finished
    If NumOfCasesLeft = 19 Then pickcount = 5
    If NumOfCasesLeft = 14 Then pickcount = 4
    If NumOfCasesLeft = 10 Then pickcount = 3
    If NumOfCasesLeft = 7 Then pickcount = 2
    If NumOfCasesLeft = 5 Then pickcount = 1
    If NumOfCasesLeft = 4 Then pickcount = 1
    If NumOfCasesLeft = 3 Then pickcount = 1
    If NumOfCasesLeft = 2 Then pickcount = 1
    If NumOfCasesLeft = 1 Then
     Display.Text = " You Won " & Format(lblAmounts(PM(MyCase)).Caption, "$#,##0.00") & " "
      lblSelectedNum.Caption = lblAmounts(PM(MyCase)).Caption   ' show users prize money
      For x = 0 To 25                                                                     ' open remaining briefcases and show prize money
         If imgBriefCase(x).Visible = True Then
               lblCaseNum(x).Caption = lblAmounts(PM(x)).Caption
         End If
      Next x
      cmdStart.Visible = True
      Picture1.Enabled = False                                     ' lock gameboard so you can't screw around
      txtPrevOffer.Visible = False
      Exit Sub
    End If
      If pickcount = 1 Then
         Display.Text = "         Pick " & pickcount & " Case      "
      Else
         Display.Text = "         Pick " & pickcount & " Cases     "
      End If
      pickcount = pickcount - 1
      Picture1.Enabled = True                                        'enable game borad to continue game
      Picture1.BackColor = &HFF0000                            'change background color back to blue
      Picture2.BackColor = &HFF0000
      Picture4.BackColor = &HFF0000
      OT1.BackCol = &HFF0000
      OT2.BackCol = &HFF0000
      OT3.BackCol = &HFF0000
      OT4.Caption = OT4.Caption + 1                            ' advance Round number
      txtPrevOffer.Visible = False
End Sub

Private Sub cmdExit_Click()
   DelayTimer.Enabled = False
   Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     FormDrag Me
End Sub

Private Sub imgBriefCase_Click(Index As Integer)   'in case user clicks on image and not label
   PickingCases Index
End Sub

Private Sub lblCaseNum_Click(Index As Integer)     'in case user clicks on label and not image
   PickingCases Index
End Sub

Private Sub PickingCases(Index As Integer)
 If SelectingMyCase = True Then
      imgBriefCase(Index).Visible = False                      ' place selected briefcase at bottom of gameboard
      lblCaseNum(Index).Visible = False
      imgSelectedCase.Visible = True
      lblSelectedNum.Visible = True
      lblSelectedNum.Caption = Index + 1
      MyCase = Index                                             'save selected case index number
      NumOfCasesLeft = NumOfCasesLeft - 1
     Display.Text = "         Pick " & pickcount & " Cases"
      pickcount = pickcount - 1
      SelectingMyCase = False
      OT4.Visible = True                                        ' show Round number
      Exit Sub
   End If
   If Index = MyCase Then Exit Sub                         'if user clicks on selected case it will not disappear
   DelayTimer.Enabled = True                                  ' hide case that was clicked on
   If pickcount <> 0 Then
        Display.Text = "         Pick " & pickcount & " Cases"
   End If
   CasePicked = Index
    If PriceArray(PM(Index)) >= 100 Then
       lblCaseNum(Index).Caption = Format(PriceArray(PM(Index)), "#,##0")
    Else
       lblCaseNum(Index).Caption = Format(PriceArray(PM(Index)), "#,##0.00")
    End If
    If pickcount = 0 Then
      Picture1.Enabled = False                               'so briefcases cannot be clicked on at end of round
      Picture1.BackColor = vbRed                           ' offer time so change background colors
      Picture2.BackColor = vbRed
      Picture4.BackColor = vbRed
      OT1.BackCol = vbRed
      OT2.BackCol = vbRed
      OT3.BackCol = vbRed
      Calculate                                                         ' time to make an offer
     Display.Text = "     It's The Banker..."
      Offer.Enabled = True
      Exit Sub
   End If
   pickcount = pickcount - 1
End Sub

Private Sub DelayTimer_Timer()
   Dim x As Integer
   'hide case after being selected
   For x = 0 To 25
      If lblCaseNum(CasePicked).Caption = lblAmounts(x).Caption Then
         lblAmounts(x).Visible = False
     End If
   Next x
   imgBriefCase(CasePicked).Visible = False
   lblCaseNum(CasePicked).Visible = False
   NumOfCasesLeft = NumOfCasesLeft - 1
   DelayTimer.Enabled = False
End Sub

Private Sub Offer_Timer()
Display.Text = " His Offer is " & Format(sum, "$#,##0")
txtPrevOffer.Visible = True
Offer.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   DelayTimer.Enabled = False
   Unload Me
End Sub

Private Sub Calculate()
Dim x As Integer

bCount = 0
sum = 0

For x = 0 To 25                                      ' just to make sure selected amounts are not showing in amounts columns
    If lblCaseNum(CasePicked).Caption = lblAmounts(x).Caption Then
         lblAmounts(x).Visible = False
   End If
   If lblAmounts(x).Visible = True Then       ' count the number of remaining amounts
     bCount = bCount + 1
   End If
Next x

For x = 0 To 25                                      ' now add all amounts remaining , together
   If lblAmounts(x).Visible = True Then
      sum = sum + lblAmounts(x).Caption
   End If
Next x

  ' I'm using simple averaging for the math... nothing complicated, but seems to work okay
   If NumOfCasesLeft = 20 Then sum = (sum / bCount) / 2
   If NumOfCasesLeft = 15 Then sum = (sum / bCount) / 1.7
   If NumOfCasesLeft = 11 Then sum = (sum / bCount) / 1.5
   If NumOfCasesLeft = 8 Then sum = (sum / bCount) / 1.3
   If NumOfCasesLeft = 6 Then sum = (sum / bCount)
   If NumOfCasesLeft = 5 Then sum = (sum / bCount)
   If NumOfCasesLeft = 4 Then sum = (sum / bCount)
   If NumOfCasesLeft = 3 Then sum = (sum / bCount)
   If NumOfCasesLeft = 2 Then sum = (sum / bCount)
   If NumOfCasesLeft = 1 Then sum = (sum / bCount)
   
   ' round off the offer by removing the last three digits and replacing them with 000
   If sum >= 1000 Then
      sum = (sum / 1000)
      sum = sum & "000"
   End If
   PrevArray(pc) = Format(sum, "$#,##0")
   txtPrevOffer.Text = txtPrevOffer.Text & PrevArray(pc) & vbCrLf
   pc = pc + 1
   cmdDeal.Visible = True
   cmdNoDeal.Visible = True
End Sub

Private Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
