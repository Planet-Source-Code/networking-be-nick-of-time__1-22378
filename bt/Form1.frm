VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "TimerTest"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   3345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblConclusion 
      Height          =   1095
      Left            =   360
      TabIndex        =   12
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label11 
      Caption         =   "Conclusion:"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label lblACPC 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label lblDNPC 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblACTC 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblDNTC 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Accuricy:"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "# different numbers"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "QueryPerformanceCounter"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Accuricy:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "# different numbers"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "GetTickCount:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub Command1_Click()
    
    Command1.Enabled = False

    Dim BT As New Class1
    
    ' this is the frequency the timer run on
    ' the amount of times per second
    f = BT.Frequency
    
    Dim lStart As Long
    Dim DiffTickCount As Long
    Dim DiffPC As Long
    Dim PrevVal As Long
    Dim ThisVal As Long

    ' do thest using gettickcount
    PrevVal = 0
    ThisVal = 0
    lStart = GetTickCount
    Do Until ThisVal > (lStart + 5000) ' add 5 secs
        ThisVal = GetTickCount
        If ThisVal <> PrevVal Then
            PrevVal = ThisVal
            DiffTickCount = DiffTickCount + 1
        End If
    Loop

    ' do test using queryperformancecount
    PrevVal = 0
    ThisVal = 0
    lStart = BT.Count
    Do Until ThisVal > (lStart + (5 * f)) ' add 5 secs
        ThisVal = BT.Count
        If ThisVal <> PrevVal Then
            PrevVal = ThisVal
            DiffPC = DiffPC + 1
        End If
    Loop

    Dim AccTC As Double
    Dim AccPC As Double
    
    ' 5000 is the number of milliseconds they ran
    AccTC = Round(5000 / DiffTickCount, 4)
    AccPC = Round(5000 / DiffPC, 4)
    
    ' display statistics
    lblDNTC.Caption = DiffTickCount
    lblACTC.Caption = AccTC & " ms"
    lblDNPC.Caption = DiffPC
    lblACPC.Caption = AccPC & " ms"

    Dim strCon As String
    If DiffTickCount > DiffPC Then
        ' it would be odd if you ever came here
        strCon = "GetTickCount is more precise as QueryPerformanceCount" & vbCrLf & _
                 Round(DiffTickCount / DiffPC * 100, 4) & " %"
    Else
        strCon = "QueryperformanceCount is more precise as GetTickCount" & vbCrLf & _
                 Round(DiffPC / DiffTickCount * 100, 4) & " %"
    End If

    lblConclusion.Caption = strCon
    Command1.Enabled = True

End Sub

Private Sub Label10_Click()

End Sub
