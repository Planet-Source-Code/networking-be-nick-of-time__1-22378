VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type
Private Freq As Long
Private LI As LARGE_INTEGER
Public Enabled As Boolean

Public Property Get Frequency() As Long

    Frequency = Freq

End Property

Public Property Get Count() As Long

    Dim Retval As Long
    Retval = QueryPerformanceCounter(LI)

    If Retval = 0 Then
        Count = -1
    Else
        Count = LI.lowpart
    End If

End Property

Private Sub Class_Initialize()

    Dim Retval As Long
    
    Retval = QueryPerformanceFrequency(LI)

    If Retval = 0 Then
        MsgBox "The system does not support high frequency timers"
    Else
        Freq = LI.lowpart
    End If

End Sub
