VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "WinXP Frames & TextBoxes"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin Project1.XPFrame XPFrame5 
      Height          =   2715
      Left            =   3225
      TabIndex        =   13
      Top             =   1950
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   4789
      Caption         =   "Normal Flat ListBox"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontBold        =   0   'False
      FontItalic      =   0   'False
      Align           =   2
      Begin VB.CommandButton Command1 
         Caption         =   "How to Use?"
         Height          =   315
         Left            =   1350
         TabIndex        =   17
         Top             =   2250
         Width           =   1890
      End
      Begin VB.Label Label1 
         Caption         =   $"XPFrame.frx":0000
         Height          =   1965
         Left            =   75
         TabIndex        =   18
         Top             =   225
         Width           =   3240
      End
   End
   Begin Project1.XPFrame XPFrame4 
      Height          =   2715
      Left            =   75
      TabIndex        =   9
      Top             =   1950
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   3995
      Caption         =   "Flat List Box (using XPListBox)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontBold        =   0   'False
      FontItalic      =   0   'False
      Align           =   2
      Begin Project1.XPSimpleFrame XPSimpleFrame4 
         Height          =   2040
         Left            =   225
         TabIndex        =   10
         Top             =   450
         Width           =   2640
         _ExtentX        =   4524
         _ExtentY        =   2910
         Begin Project1.XPListBox XPListBox1 
            Height          =   1950
            Left            =   45
            TabIndex        =   11
            Top             =   45
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   3440
            Begin VB.ListBox List1 
               Appearance      =   0  'Flat
               Height          =   1980
               ItemData        =   "XPFrame.frx":01B7
               Left            =   -15
               List            =   "XPFrame.frx":01F7
               TabIndex        =   12
               Top             =   -15
               Width           =   2580
            End
         End
      End
   End
   Begin Project1.XPFrame XPFrame2 
      Height          =   1815
      Left            =   3225
      TabIndex        =   1
      Top             =   75
      Width           =   3390
      _ExtentX        =   5318
      _ExtentY        =   3201
      Caption         =   "Multi Line"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   8421504
      FontName        =   "Tahoma"
      FontSize        =   8.25
      FontBold        =   0   'False
      FontItalic      =   0   'False
      Begin Project1.XPSimpleFrame XPSimpleFrame1 
         CausesValidation=   0   'False
         Height          =   990
         Left            =   225
         TabIndex        =   2
         Top             =   600
         Width           =   2865
         _ExtentX        =   4392
         _ExtentY        =   1746
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   45
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Text            =   "XPFrame.frx":02E2
            Top             =   45
            Width           =   2775
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Multi Line Text :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   150
         TabIndex        =   8
         Top             =   375
         Width           =   2040
      End
   End
   Begin Project1.XPFrame XPFrame1 
      Height          =   1815
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   3201
      Caption         =   "XP Frame"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Verdana"
      FontSize        =   8.25
      FontBold        =   0   'False
      FontItalic      =   0   'False
      Begin VB.OptionButton Option1 
         Caption         =   "Center"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   1875
         TabIndex        =   16
         Top             =   900
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Right"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1875
         TabIndex        =   15
         Top             =   600
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1875
         TabIndex        =   14
         Top             =   300
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Font Italic"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   7
         Top             =   675
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Font Bold"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   375
         Width           =   1215
      End
      Begin Project1.XPSimpleFrame XPSimpleFrame2 
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Top             =   1050
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   45
            TabIndex        =   5
            Text            =   "XP Frame"
            Top             =   45
            Width           =   1500
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
    XPFrame1.FontBold = Check1.Value
End Sub

Private Sub Check2_Click()
    XPFrame1.FontItalic = Check2.Value
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub Option1_Click(Index As Integer)
    XPFrame1.Align = Index
End Sub

Private Sub Text2_Change()
    XPFrame1.Caption = Text2.Text
End Sub

