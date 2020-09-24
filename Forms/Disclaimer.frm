VERSION 5.00
Begin VB.Form Disclaimer 
   BackColor       =   &H80000012&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Disclaimer"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Disclaimer.frx":0000
   ScaleHeight     =   4425
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Ishihara.CandyButton CandyButton1 
      Height          =   330
      Left            =   5235
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Accept"
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Ishihara.CandyButton CandyButton2 
      Height          =   330
      Left            =   2760
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Decline"
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H00A40404&
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "Disclaimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
Open App.Path & "\DonotShow" For Output As #1
Close #1
End Sub


Private Sub CandyButton1_Click()
On Error Resume Next
Select Case Index
    Case 0
        Open App.Path & "\Accept" For Output As #1
        Close #1
        Kill App.Path & "\Decline"
        Unload Me
        Load AfibChad2
        AfibChad2.Show
    Case 1
        Open App.Path & "\Decline" For Output As #1
        Close #1
        Kill App.Path & "\Accept"
        Unload Me
End Select

End Sub

Private Sub CandyButton2_Click()
On Error Resume Next
        Open App.Path & "\Decline" For Output As #1
        Close #1
        Kill App.Path & "\Accept"
        Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
End Sub

Private Sub Form_Load()
On Error Resume Next
SetTopMostWindow Me.hwnd, True
If Dir(App.Path & "\Accept") <> "" Then
    Unload Me
    Exit Sub
End If
CandyButton1.Style = 2
CandyButton2.Style = 2

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
If Dir(App.Path & "\Accept") <> "" Then
    Load Ishahara
    Ishahara.Show
    Ishahara.Enabled = True
Else
    CloseAll
End If
Set Disclaimer = Nothing
End Sub
Sub CloseAll()
    On Error Resume Next
    Dim intFrmNum As Integer
    intFrmNum = Forms.Count


    Do Until intFrmNum = 0
        Unload Forms(intFrmNum - 1)
        Set Forms(intFrmNum - 1) = Nothing
        intFrmNum = intFrmNum - 1
    Loop
End Sub
