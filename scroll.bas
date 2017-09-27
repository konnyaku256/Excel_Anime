Attribute VB_Name = "Module1"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
(ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Sub Scroll()
  Application.Calculation = xlCalculationManual
  
  Dim SoundFile As String, rc As Long
    SoundFile = "D:\excel_anime\kemofure.mp3"
    If Dir(SoundFile) = "" Then
        MsgBox SoundFile & vbCrLf & "‚ª‚ ‚è‚Ü‚¹‚ñB", vbExclamation
        Exit Sub
    End If
    rc = mciSendString("Play " & SoundFile, "", 0, 0)

  Dim myCnt As Long
  myCnt = 0
  Dim IMG_CNT As Long '‰æ‘œ–‡”
  IMG_CNT = 692
    Do Until myCnt > IMG_CNT
        ActiveWindow.SmallScroll Down:=360, ToLeft:=0
        'If myCnt Mod 2 = 0 Then Call Sleep(10)
        Call Sleep(51)
        If myCnt Mod 40 = 0 Then DoEvents
        myCnt = myCnt + 1
        
    Loop
  
  Application.Calculation = xlCalculationAutomatic
End Sub
