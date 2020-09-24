VERSION 5.00
Begin VB.Form Dialog1 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ishihara Test"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton e 
      BackColor       =   &H8000000E&
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton d 
      BackColor       =   &H8000000E&
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton c 
      BackColor       =   &H8000000E&
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton None 
      BackColor       =   &H8000000E&
      Caption         =   "None"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton a 
      BackColor       =   &H8000000E&
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton b 
      BackColor       =   &H8000000E&
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The numerals in parenthesis show that they can be read but they are comparatively unclear."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Which Number did you see?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   4005
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub a_Click()
    Ishahara.Picture1.Visible = False
    Ishahara.List4.AddItem (Trim(a.Caption))
    Unload Me

End Sub

Private Sub b_Click()
Dim x As String
    Ishahara.Picture1.Visible = False
    'x = Replace((Trim(b.Caption)), ")", "0")
    'x = Replace(x, "(", "0")
    Ishahara.List4.AddItem Trim(b.Caption)
    Unload Me

End Sub

Private Sub c_Click()
    Ishahara.Picture1.Visible = False
    Ishahara.List4.AddItem Trim(c.Caption)
    Unload Me

End Sub

Private Sub d_Click()
    Ishahara.Picture1.Visible = False
    Ishahara.List4.AddItem (Trim(d.Caption))
    Unload Me

End Sub

Private Sub e_Click()
Dim x As String
    Ishahara.Picture1.Visible = False
    'x = Replace((Trim(e.Caption)), ")", "0")
    'x = Replace(x, "(", "0")
    Ishahara.List4.AddItem Trim(e.Caption)
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
 Ishahara.Enabled = True
End Sub

Private Sub None_Click()
    Ishahara.Picture1.Visible = False
    Ishahara.List4.AddItem ("0")
    Unload Me

End Sub



