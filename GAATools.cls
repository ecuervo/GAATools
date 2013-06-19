VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GAATools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long

Public Function bInternet() As Boolean
    Dim flags As Long
    On Local Error Resume Next
    bInternet = CBool(InternetGetConnectedState(flags, 0&))
End Function
Public Function fnExisteArchivo(sRuta As String) As Boolean
On Local Error GoTo errH
fnExisteArchivo = False
GetAttr sRuta
fnExisteArchivo = True
Exit Function
errH:
fnExisteArchivo = False
End Function
Public Function fnExisteDirectorio(sRuta As String) As Boolean
On Local Error GoTo errH
fnExisteDirectorio = False
GetAttr sRuta
fnExisteDirectorio = True
Exit Function
errH:
fnExisteDirectorio = False
End Function
Public Function fnArrVacioByte(ByRef Arr() As Byte) As Boolean
On Local Error GoTo errH
Dim i As Integer
i = UBound(Arr)
fnArrVacioByte = False
Exit Function
errH:
fnArrVacioByte = True
End Function
Public Function fnArrVacioStr(ByRef Arr() As String) As Boolean
On Local Error GoTo errH
Dim i As Integer
i = UBound(Arr)
fnArrVacioStr = False
Exit Function
errH:
fnArrVacioStr = True
End Function
Public Function fnArrVacioInt(ByRef Arr() As Integer) As Boolean
On Local Error GoTo errH
Dim i As Integer
i = UBound(Arr)
fnArrVacioInt = False
Exit Function
errH:
fnArrVacioInt = True
End Function
Public Function fnArrVacioCls(ByRef Arr() As clsZK) As Boolean
On Local Error GoTo errH
Dim i As Integer
i = UBound(Arr)
If i = -1 Then
    fnArrVacioCls = True
Else
    fnArrVacioCls = False
End If
Exit Function
errH:
fnArrVacioCls = True
End Function

