Attribute VB_Name = "modIshahara"
Option Explicit
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetDesktopWindow Lib "User32" () As Long

Public Const SW_SHOWDEFAULT = 10
Public Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetSystemDirectoryB Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal Path As String, ByVal cbBytes As Long) As Long
Private Const MAX_LENGTH = 512
Global j As Integer, tempPath As String, Reported As String
Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
 If Topmost = True Then 'Make the window topmost
  SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
 Else
  SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
  SetTopMostWindow = False
 End If
End Function
Public Function GetWindowsSystemDirectory() As String
   Dim s As String
   Dim c As Long
   s = String$(MAX_LENGTH, 0)
   c = GetSystemDirectoryB(s, MAX_LENGTH)
   If c > 0 Then
       If c > Len(s) Then
           s = Space$(c + 1)
           c = GetSystemDirectoryB(s, MAX_LENGTH)
       End If
   End If
   GetWindowsSystemDirectory = IIf(c > 0, Left$(s, c), "")
End Function

Public Function StartDoc(DocName As String) As Long
On Error Resume Next

Dim Scr_hDC As Long
Scr_hDC = GetDesktopWindow()
StartDoc = ShellExecute(Scr_hDC, "", DocName, _
"", "C:\", -1)
End Function

