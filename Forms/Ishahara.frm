VERSION 5.00
Begin VB.Form Ishahara 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Visionary"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   8400
   Enabled         =   0   'False
   Icon            =   "Ishahara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   8400
   StartUpPosition =   1  'CenterOwner
   Begin Ishihara.cmdopen CmDlg 
      Left            =   360
      Top             =   6360
      _ExtentX        =   661
      _ExtentY        =   635
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8070
      Left            =   8400
      Picture         =   "Ishahara.frx":08CA
      ScaleHeight     =   8070
      ScaleWidth      =   9420
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   9420
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8070
      Left            =   8280
      MouseIcon       =   "Ishahara.frx":9E06
      MousePointer    =   99  'Custom
      Picture         =   "Ishahara.frx":A6D0
      ScaleHeight     =   8070
      ScaleWidth      =   8385
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   8385
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6240
         Left            =   7230
         ScaleHeight     =   6240
         ScaleWidth      =   1170
         TabIndex        =   10
         Top             =   1065
         Width           =   1170
         Begin Ishihara.CandyButton CandyButton3 
            Height          =   210
            Left            =   240
            TabIndex        =   11
            Top             =   3480
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   370
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Undo"
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            Checked         =   0   'False
            ColorButtonHover=   16760976
            ColorButtonUp   =   15309136
            ColorButtonDown =   15309136
            BorderBrightness=   0
            ColorBright     =   16772528
            DisplayHand     =   0   'False
            ColorScheme     =   0
         End
         Begin Ishihara.CandyButton CandyButton4 
            Height          =   210
            Left            =   240
            TabIndex        =   12
            Top             =   3840
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   370
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Save"
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            Checked         =   0   'False
            ColorButtonHover=   16760976
            ColorButtonUp   =   15309136
            ColorButtonDown =   15309136
            BorderBrightness=   0
            ColorBright     =   16772528
            DisplayHand     =   0   'False
            ColorScheme     =   0
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   8070
         Left            =   8250
         ScaleHeight     =   8070
         ScaleWidth      =   8385
         TabIndex        =   9
         Top             =   2460
         Visible         =   0   'False
         Width           =   8385
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5220
      Left            =   8400
      ScaleHeight     =   5220
      ScaleWidth      =   7260
      TabIndex        =   2
      Top             =   3765
      Visible         =   0   'False
      Width           =   7260
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   165
      TabIndex        =   5
      Top             =   105
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   1050
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List3 
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List4 
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   135
      Visible         =   0   'False
      Width           =   855
   End
   Begin Ishihara.CandyButton CandyButton2 
      Height          =   330
      Left            =   420
      TabIndex        =   1
      Top             =   7560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Start Again"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Ishihara.CandyButton CandyButton1 
      Height          =   330
      Left            =   6930
      TabIndex        =   0
      Top             =   7590
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Next >>"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Image Image1 
      Height          =   8055
      Index           =   2
      Left            =   15
      Picture         =   "Ishahara.frx":14E30
      Stretch         =   -1  'True
      Top             =   30
      Visible         =   0   'False
      Width           =   8355
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   24
      Left            =   405
      Picture         =   "Ishahara.frx":1C406
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   23
      Left            =   2025
      Picture         =   "Ishahara.frx":23BF1
      Stretch         =   -1  'True
      Top             =   1110
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   22
      Left            =   1995
      Picture         =   "Ishahara.frx":2B0FA
      Stretch         =   -1  'True
      Top             =   615
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   21
      Left            =   2520
      Picture         =   "Ishahara.frx":31928
      Stretch         =   -1  'True
      Top             =   1290
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   20
      Left            =   420
      Picture         =   "Ishahara.frx":38B17
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   19
      Left            =   0
      Picture         =   "Ishahara.frx":3EA45
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   18
      Left            =   405
      Picture         =   "Ishahara.frx":45C83
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   17
      Left            =   0
      Picture         =   "Ishahara.frx":4CBC8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   16
      Left            =   405
      Picture         =   "Ishahara.frx":53847
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   15
      Left            =   2025
      Picture         =   "Ishahara.frx":5AF70
      Stretch         =   -1  'True
      Top             =   1110
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   14
      Left            =   1995
      Picture         =   "Ishahara.frx":6159B
      Stretch         =   -1  'True
      Top             =   615
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   13
      Left            =   2520
      Picture         =   "Ishahara.frx":68673
      Stretch         =   -1  'True
      Top             =   1290
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   12
      Left            =   420
      Picture         =   "Ishahara.frx":6EB49
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   11
      Left            =   0
      Picture         =   "Ishahara.frx":74892
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   10
      Left            =   405
      Picture         =   "Ishahara.frx":7A978
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   9
      Left            =   0
      Picture         =   "Ishahara.frx":80FED
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   8
      Left            =   405
      Picture         =   "Ishahara.frx":87C47
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   7
      Left            =   2025
      Picture         =   "Ishahara.frx":8EB56
      Stretch         =   -1  'True
      Top             =   1110
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   6
      Left            =   1995
      Picture         =   "Ishahara.frx":95A04
      Stretch         =   -1  'True
      Top             =   615
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   5
      Left            =   2520
      Picture         =   "Ishahara.frx":9CBF6
      Stretch         =   -1  'True
      Top             =   1290
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   4
      Left            =   420
      Picture         =   "Ishahara.frx":A33A2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2340
      Index           =   3
      Left            =   0
      Picture         =   "Ishahara.frx":A9E88
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuOpenAmsler 
         Caption         =   "Open Amsler"
      End
      Begin VB.Menu mnuSavee 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuPrintAmsler 
         Caption         =   "Print Blank Amsler Chart"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuReport 
         Caption         =   "Color Blindness Report"
      End
      Begin VB.Menu trywrtg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColorBlindness 
         Caption         =   "Color Blindness Charts"
      End
      Begin VB.Menu mnuCBInstr 
         Caption         =   "Color Blindness Instructions"
      End
      Begin VB.Menu gretrwt 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAmsler 
         Caption         =   "Amsler Chart"
         Begin VB.Menu mnuChart 
            Caption         =   "Chart"
         End
         Begin VB.Menu mnuInstructions 
            Caption         =   "Instructions"
         End
      End
      Begin VB.Menu werfdesfde 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRedDesat 
         Caption         =   "Red Desaturation"
         Begin VB.Menu mnuRedPage 
            Caption         =   "Red Page"
         End
         Begin VB.Menu mnuInstr 
            Caption         =   "Instructions"
         End
      End
      Begin VB.Menu mnuSnellen 
         Caption         =   "Near and Far Visual Acuity"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuDisclaimer 
         Caption         =   "Disclaimer"
      End
      Begin VB.Menu mnuInstruct 
         Caption         =   "Instructions"
         Begin VB.Menu mnuCBlind 
            Caption         =   "Color Blindness"
         End
         Begin VB.Menu mnuAAmsler 
            Caption         =   "Amsler"
         End
         Begin VB.Menu mnuRedD 
            Caption         =   "Red Chart"
         End
      End
      Begin VB.Menu wfwfwef 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProgramInstruct 
         Caption         =   "Program Instructions"
      End
      Begin VB.Menu vcxdsvdsdv 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuArchive 
      Caption         =   "Archive"
   End
End
Attribute VB_Name = "Ishahara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Saveflag As Boolean

Private Sub CandyButton1_Click()
On Error Resume Next
Select Case j
    Case 2
        Dialog.CancelButton.Caption = "3": Dialog.OKButton.Caption = "8"
        Ishahara.Enabled = False: Dialog.Show
        Picture1.Visible = True
    Case 3
       Dialog.CancelButton.Caption = "70": Dialog.OKButton.Caption = "29"
        Ishahara.Enabled = False: Dialog.Show
        Picture1.Visible = True
    Case 4
       Dialog.CancelButton.Caption = "2": Dialog.OKButton.Caption = "5"
        Ishahara.Enabled = False: Dialog.Show
        Picture1.Visible = True
    Case 5
       Dialog.CancelButton.Caption = "5": Dialog.OKButton.Caption = "3"
        Ishahara.Enabled = False: Dialog.Show
        Picture1.Visible = True
    Case 6
       Dialog.CancelButton.Caption = "17": Dialog.OKButton.Caption = "15"
        Ishahara.Enabled = False: Dialog.Show
        Picture1.Visible = True
    Case 7
       Dialog.CancelButton.Caption = "21": Dialog.OKButton.Caption = "74"
        Ishahara.Enabled = False: Dialog.Show
        Picture1.Visible = True
    Case 8
       Dialog.CancelButton.Visible = False: Dialog.OKButton.Caption = "6"
        Ishahara.Enabled = False: Dialog.Show
        Picture1.Visible = True
    Case 9
       Dialog.CancelButton.Visible = False: Dialog.OKButton.Caption = "45"
        Ishahara.Enabled = False: Dialog.Show
        Picture1.Visible = True
    Case 10
       Dialog.CancelButton.Visible = False: Dialog.OKButton.Caption = "5"
        Ishahara.Enabled = False: Dialog.Show
        Picture1.Visible = True
    Case 11
       Dialog.CancelButton.Visible = False: Dialog.OKButton.Caption = "7"
        Ishahara.Enabled = False: Dialog.Show
        Picture1.Visible = True
    Case 12
       Dialog.CancelButton.Visible = False: Dialog.OKButton.Caption = "16"
        Ishahara.Enabled = False: Dialog.Show
        Picture1.Visible = True
    Case 13
       Dialog.CancelButton.Visible = False: Dialog.OKButton.Caption = "73"
        Ishahara.Enabled = False: Dialog.Show
        Picture1.Visible = True
    Case 14
       Dialog.CancelButton.Caption = "5": Dialog.OKButton.Visible = False
        Ishahara.Enabled = False: Dialog.Show
        Picture1.Visible = True
    Case 15
       Dialog.CancelButton.Caption = "45": Dialog.OKButton.Visible = False
        Ishahara.Enabled = False: Dialog.Show
        Picture1.Visible = True
    Case 16
        Dialog1.a.Caption = "26"
        Dialog1.b.Caption = "6"
        Dialog1.c.Caption = "(2)6"
        Dialog1.d.Caption = "2"
        Dialog1.e.Caption = "2(6)"
        Ishahara.Enabled = False: Dialog1.Show
        Picture1.Visible = True
    Case 17
        Dialog1.a.Caption = "42"
        Dialog1.b.Caption = "2"
        Dialog1.c.Caption = "(4)2"
        Dialog1.d.Caption = "2"
        Dialog1.e.Caption = "4(2)"
        Ishahara.Enabled = False: Dialog1.Show
        Picture1.Visible = True
End Select


j = j + 1
If j > 17 Then CandyButton1.Visible = False: mnuReport.Enabled = True: mnuSave.Enabled = True: Ishahara.mnuSavee.Enabled = True: Saveflag = False: Exit Sub
Image1(j - 1).Visible = False
Image1(j).Visible = True
Me.Caption = "Color Visionary: Plate " & j

End Sub

Private Sub CandyButton2_Click()
On Error Resume Next
Me.Caption = "Color Visionary: Plate 2"
j = 2
For i = 3 To 24
    Image1(i).Visible = False
Next
mnuSave.Enabled = False
mnuSavee.Enabled = False
List4.Clear
CandyButton1.Visible = True
mnuReport.Enabled = False
Image1(2).Visible = True
Saveflag = True
End Sub

Private Sub CandyButton3_Click()
Picture2.Picture = Picture4.Picture
If Dir(App.Path & "\Amsler*.bmp") = "" Then mnuOpenAmsler.Enabled = False

End Sub

Private Sub CandyButton4_Click()
mnuOpenAmsler.Enabled = True
SavePicture Picture2, App.Path & "\Amsler " & Format(Now, "ddmmyyhhmmss") & ".bmp"

End Sub

Private Sub Form_Initialize()
On Error Resume Next
Dim i As Integer
CandyButton1.Style = 2
CandyButton2.Style = 2
CandyButton3.Style = 2
CandyButton4.Style = 2
Picture2.DrawWidth = 3
Picture4.Picture = Picture2.Image
Me.Caption = "Color Visionary: Plate 2"
If Dir(App.Path & "\Amsler*.bmp") = "" Then mnuOpenAmsler.Enabled = False
For i = 3 To 18
    Image1(i).Visible = False
    Image1(i).Top = 0
    Image1(i).Left = 0
    Image1(i).Height = Image1(2).Height
    Image1(i).Width = Image1(2).Width
Next
Picture1.Left = 0
Picture1.Top = 0
Picture1.Height = Image1(2).Height
Picture1.Width = Image1(2).Width
Picture2.Left = 270
Picture2.Top = -120
Picture3.Left = -915
Picture3.Top = 135
j = 2
HelpFiles
mnuArchive.Visible = False
SetTopMostWindow Dialog.hWnd, True
SetTopMostWindow Dialog1.hWnd, True
CandyButton2_Click
End Sub

Private Sub Form_Load()
Dim i As Integer

List1.AddItem 3
List1.AddItem 70
List1.AddItem 2
List1.AddItem 5
List1.AddItem 17
List1.AddItem 21
List1.AddItem 0
List1.AddItem 0
List1.AddItem 0
List1.AddItem 0
List1.AddItem 0
List1.AddItem 0
List1.AddItem 5
List1.AddItem 45

List2.AddItem 0
List2.AddItem 0
List2.AddItem 0
List2.AddItem 0
List2.AddItem 0
List2.AddItem 0
List2.AddItem 0
List2.AddItem 0
List2.AddItem 0
List2.AddItem 0
List2.AddItem 0
List2.AddItem 0
List2.AddItem 0
List2.AddItem 0

List3.AddItem 8
List3.AddItem 29
List3.AddItem 5
List3.AddItem 3
List3.AddItem 15
List3.AddItem 74
List3.AddItem 6
List3.AddItem 45
List3.AddItem 5
List3.AddItem 7
List3.AddItem 16
List3.AddItem 73
List3.AddItem 0
List3.AddItem 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim intSave As String
If Saveflag = False Then
    intSave = MsgBox("Do you want to Save the Results?", _
                     vbYesNoCancel + vbExclamation)
    Select Case intSave
      Case vbYes
        Cancel = True
        mnuSave_Click
        Exit Sub
      Case vbCancel
        Cancel = True
        Exit Sub
    End Select
End If
Unload Me
CloseAll
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
Private Sub Image1_Click(Index As Integer)
'Image1(3).Visible = True
'Image1(2).Visible = False
End Sub

Private Sub List4_Click()
'MsgBox List4.List(List4.ListIndex)
End Sub

Private Sub mnuAAmsler_Click()
    'StartDoc App.Path & "\Amsler.pdf"
    StartDoc tempPath & "\Amsler.pdf"
End Sub

Private Sub mnuAbout_Click()
Load frmAbout
frmAbout.Show
End Sub

Private Sub mnuArchive_Click()
    'On Error Resume Next
    Dim nFiles As Long
    'nFiles = SaveFilesToArchiv(ProjectDir, App.Path & "\Archival.dat", "*.*")
    Open App.Path & "\e.rc" For Output As #1
        Print #1, "0" & "   CUSTOM " & App.Path & "\Ishihara.pdf"
        Print #1, "1" & "   CUSTOM " & App.Path & "\Amsler.pdf"
        Print #1, "2" & "   CUSTOM " & App.Path & "\Near.pdf"
        Print #1, "3" & "   CUSTOM " & App.Path & "\Ishiharas.pdf"
    Close #1
    Open App.Path & "\bat.bat" For Output As #1
        Print #1, App.Path & "\RC.exe /r /fo " & App.Path & "\Res.RES " & _
        App.Path & "\e.rc"
    Close #1
    'Kill "c:\Res.RES"
    'a = ShellExecute(Me.hwnd, "", App.Path & "\bat.bat", "", _
    "c:\", 10)
    'Shell App.Path & "\bat.bat", vbMinimizedFocus
    
    'mnuCompileRC.Enabled = True
Exit Sub
ooops:
End Sub
Public Function SaveFilesToArchiv( _
  ByVal sPath As String, _
  ByVal sArchiv As String, _
  Optional ByVal sPattern As String = "*.*") As Long

  Dim F As Integer
  Dim n As Integer
  Dim nLenFileName As Integer
  Dim nLenFileData As Long
  Dim DirName As String
  Dim FileData As String
  Dim File() As String
  Dim nFiles As Long
  Dim i As Long
  Dim lngUBound As Long

  ' Add backslash to the path
  If Right$(sPath, 1) <> "\" Then sPath = sPath + "\"

  ' Get all files in the directory
  nFiles = 0

  DirName = Dir(sPath & sPattern, vbNormal)
  While DirName <> ""
    If DirName <> "." And DirName <> ".." Then
      nFiles = nFiles + 1
      'Get files
      If nFiles > lngUBound Then lngUBound = 2 * nFiles
      ReDim Preserve File(lngUBound)
      File(nFiles) = DirName
    End If
    DirName = Dir
  Wend
  ReDim Preserve File(nFiles)

  ' If archiv exists already, delete it
  If Dir(sArchiv) <> "" Then Kill sArchiv

  ' Now save all files to the archive
  F = FreeFile
  Open sArchiv For Binary As #F

  ' Set number of files
  Put #F, , nFiles

  For i = 1 To nFiles
    ' Save filename
    nLenFileName = Len(File(i))
    Put #F, , nLenFileName
    Put #F, , File(i)

    ' Read filedata
    n = FreeFile
    Open sPath + File(i) For Binary As #n
    FileData = Space$(LOF(n))
    Get #n, , FileData
    Close #n

    ' Save filedata to the archive
    nLenFileData = Len(FileData)
    Put #F, , nLenFileData
    Put #F, , FileData
    
    ' Progress
    DoEvents
  Next i
  Close #F
  
  SaveFilesToArchiv = nFiles
End Function
Public Sub ResToFile(Filename As String, ResID As Variant, ResType As Variant, Optional Overwrite As Boolean = False)
    Dim Buffer() As Byte
    Dim Filenum As Integer


    If Dir(Filename) <> Empty Then 'Check if output file already exists
        If Overwrite Then Kill Filename Else Err.Raise 58
    End If
    Buffer = LoadResData(ResID, ResType) 'Load the resource into a byte array
    Filenum = FreeFile
    Open Filename For Binary Access Write As Filenum
    Put Filenum, , Buffer 'Write the entire array into the file
    Close Filenum
End Sub
Private Sub mnuCBInstr_Click()
mnuCBlind_Click
End Sub

Private Sub mnuCBlind_Click()
StartDoc tempPath & "\Ishihara.pdf"

'StartDoc App.Path & "\Ishihara.pdf"
End Sub

Private Sub mnuChart_Click()
Me.Caption = "Amsler Chart: Draw What You See and Don't See"
Picture3.Visible = False
Picture1.Visible = False
Picture2.Visible = True
End Sub

Private Sub mnuColorBlindness_Click()
CandyButton2_Click

Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False

End Sub

Private Sub mnuDisclaimer_Click()
On Error Resume Next
Kill App.Path & "\Accept"
Ishahara.Enabled = False
Load Disclaimer
Disclaimer.Show
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuInstr_Click()
    mnuRedd_Click
End Sub

Private Sub mnuInstructions_Click()
Me.Caption = "Amsler Chart Instructions"
Picture3.Visible = True
Picture1.Visible = False
Picture2.Visible = False

End Sub

Private Sub mnuNew_Click()
CandyButton2_Click
End Sub
Private Sub HelpFiles()
tempPath = IIf(Environ$("tmp") <> "", Environ$("tmp"), Environ$("temp")) & "\Ishahara"
MkDir tempPath
ResToFile tempPath & "\Ishihara.pdf", "0", "CUSTOM", True
ResToFile tempPath & "\Amsler.pdf", 1, "CUSTOM", True
ResToFile tempPath & "\Near.pdf", 2, "CUSTOM", True
ResToFile tempPath & "\Ishiharas.pdf", 3, "CUSTOM", True
End Sub
Private Sub mnuOpen_Click()
On Error Resume Next
Dim i As Integer
Dim X As String
List4.Clear
Dim Filenamme As String
    With Ishahara
        .CmDlg.InitialDir = App.Path
        .CmDlg.CancelError = True 'Set cancel error to true
        .CmDlg.MultiSelect = False   'True 'Allow multi select
        .CmDlg.DialogTitle = "Open Study" 'Set dialog title
        '.CmDlg.DefaultFilename = Format(Now, "ddmmyyhhmmss") & ".ish"
        .CmDlg.Filter = "ish Files (*.ish)" & Chr$(0) & "*.ish" & Chr$(0)
        
        '.CmDlg.FilterIndex = 1 'Set filter index
        .CmDlg.ShowOpen
    End With
    If Ishahara.CmDlg.cFileName(1) = "" Then Exit Sub
    Filenamme = Ishahara.CmDlg.cFileName(1)
    Open Filenamme For Input As #1
        For i = 1 To 16
            Line Input #1, X
            List4.AddItem X
        Next
    Close #1
    mnuReport.Enabled = True
    mnuReport_Click
End Sub

Private Sub mnuOpenAmsler_Click()
On Error Resume Next
Dim i As Integer
Dim X As String
Dim Filenamme As String
    With Ishahara
        .CmDlg.InitialDir = App.Path
        .CmDlg.CancelError = True 'Set cancel error to true
        .CmDlg.MultiSelect = False   'True 'Allow multi select
        .CmDlg.DialogTitle = "Open Amsler" 'Set dialog title
        '.CmDlg.DefaultFilename = Format(Now, "ddmmyyhhmmss") & ".ish"
        .CmDlg.Filter = "Amsler Files (*.bmp)" & Chr$(0) & "*.bmp" & Chr$(0)
        
        '.CmDlg.FilterIndex = 1 'Set filter index
        .CmDlg.ShowOpen
    End With
    If Ishahara.CmDlg.cFileName(1) = "" Then Exit Sub
    Filenamme = Ishahara.CmDlg.cFileName(1)
    Picture2.Picture = LoadPicture(Filenamme)
     mnuChart_Click
End Sub

Public Sub mnuOptions_Click()

End Sub

Private Sub mnuPrintAmsler_Click()
    StartDoc tempPath & "\Amsler.pdf"
End Sub

Private Sub mnuProgramInstruct_Click()
 StartDoc tempPath & "\Ishiharas.pdf"
End Sub

Private Sub mnuRedd_Click()
    MsgBox "Red desaturation: Colour vision can be estimated by the patient " & vbCrLf & _
    "looking at a red object (e.g. a red pen / page) with each eye. If there is an " & vbCrLf & _
    "optic nerve or tract lesion on one side the color looks pink, dull or " & vbCrLf & _
    "washed out with that eye. This is ‘red desaturation'."
End Sub

Private Sub mnuRedPage_Click()
Me.Caption = "Red Desaturation Chart"
Picture3.Visible = False
Picture2.Visible = False
Picture1.Visible = True
End Sub

Public Sub mnuReport_Click()
Dim i As Integer, i1 As Integer, i2 As Integer, i3 As Integer
Dim Colrr As String
Dim Colrr1 As String
Colrr = "": Colrr1 = ""
i1 = 0
'Red-Green
For i = 0 To 13
    If Val(List4.List(i)) = Val(List1.List(i)) Then
        i1 = i1 + 1
    End If
Next
i2 = 0
'Total
For i = 0 To 13
    If Val(List4.List(i)) = Val(List2.List(i)) Then
        i2 = i2 + 1
    End If
Next
'For i = 1 To 14
'    If Val(List4.List(i)) = Val(List3.List(i)) Then
'        i3 = i3 + 1
'    End If
'Next
i3 = 0
'Normal
For i = 0 To 15
    If Val(List4.List(i)) = Val(List3.List(i)) Then
        i3 = i3 + 1
    End If
Next
For i = 15 To 16
    Select Case i
        Case 15
            If List4.List(i - 1) = "6" Then Colrr = "; Protan Strong "
            If List4.List(i - 1) = "(2)6" Then Colrr = "; Protan Mild "
            If List4.List(i - 1) = "2" Then Colrr = "; Deutan Strong "
            If Trim(List4.List(i - 1)) = "2(6)" Then Colrr = "; Deutan Mild "
            'MsgBox List4.List(i - 1)
        Case 16
            If List4.List(i - 1) = "2" Then Colrr1 = "; Protan Strong "
            If List4.List(i - 1) = "(4)2" Then Colrr1 = "; Protan Mild "
            If List4.List(i - 1) = "4" Then Colrr1 = "; Deutan Strong "
            If List4.List(i - 1) = "4(2)" Then Colrr1 = "; Deutan Mild "
    End Select
Next
'MsgBox Str(i1) & " " & Str(i2) & " " & Str(i3)
If Colrr1 <> "" And Colrr <> "" Then
    If Colrr1 <> Colrr Then Colrr = Colrr & Replace(Colrr1, ";", "vs")
End If
If i1 = 14 Then
    MsgBox "You Definitely have Red Green Color Deficiency" & Colrr & "!"
    Reported = "You Definitely have Red Green Color Deficiency" & Colrr & "!"
    GoTo Here
End If
If i2 = 14 Then
    MsgBox "You are Totally Color Blind!"
    Reported = "You are Totally Color Blind!"
    GoTo Here
End If
If i3 = 14 Then
    MsgBox "You are Normal!"
    Reported = "You are Normal!"
    GoTo Here
End If
MsgBox "You Probably have Red Green Color Deficiency" & Colrr & "!"
Reported = "You Probably have Red Green Color Deficiency" & Colrr & "!"
Here:

Load Report
Report.Show
End Sub

Private Sub mnuSave_Click()
On Error Resume Next
Dim i As Integer
Dim Filenamme As String
    With Ishahara
        .CmDlg.InitialDir = App.Path
        .CmDlg.CancelError = True 'Set cancel error to true
        .CmDlg.MultiSelect = False   'True 'Allow multi select
        .CmDlg.DialogTitle = "Save As" 'Set dialog title
        .CmDlg.DefaultFilename = Format(Now, "ddmmyyhhmmss") & ".ish"
        .CmDlg.Filter = "ish Files (*.ish)" & Chr$(0) & "*.ish" & Chr$(0)
        
        '.CmDlg.FilterIndex = 1 'Set filter index
        .CmDlg.ShowSave
    End With
    If Ishahara.CmDlg.cFileName(1) = "" Then Exit Sub
    Filenamme = Ishahara.CmDlg.cFileName(1)
    If Right(Filenamme, 4) <> ".ish" Then Filenamme = Filenamme & ".ish"
    Open Filenamme For Output As #1
        For i = 0 To List4.ListCount - 1
            Print #1, List4.List(i)
        Next
    Close #1
    Saveflag = True
End Sub

Private Sub mnuSavee_Click()
    Open Format(Now, "ddmmyyhhmmss") & ".ish" For Output As #1
        For i = 0 To List4.ListCount - 1
            Print #1, List4.List(i)
        Next
    Close #1
    Saveflag = True
End Sub

Private Sub mnuSnellen_Click()
    StartDoc tempPath & "\Near.pdf"
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        Picture2.CurrentX = X
        Picture2.CurrentY = Y
    End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
    Picture2.Line (Picture2.CurrentX, Picture2.CurrentY)-(X, Y), QBColor(1)
    End If
    Picture2.Picture = Picture2.Image

End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture2.Picture = Picture2.Image

End Sub

