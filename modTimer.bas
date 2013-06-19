Attribute VB_Name = "MODtIMER"
Option Explicit
Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public bEsperando As Boolean
Public ID_TIM As Integer
Dim idTimer As Long
Public Sub subTimer(iSeg As Integer)
bEsperando = True
idTimer = SetTimer(0, 0, iSeg * 1000, AddressOf TimerProc)
End Sub
Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
bEsperando = False
KillTimer 0, idTimer
MsgBox "Listo"
End Sub



