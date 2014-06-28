VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9045
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   9045
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text1 
      Height          =   4815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   8895
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub getpi(Optional ByVal nums As Long = 10000)
    nums = nums \ 5
    Dim max As Long, laptime As Single, result() As String
    Dim i, j As Long, d As Long, t, g, r, k As Long, f()
    laptime = Timer
    max = 10 * nums
    ReDim f(1 To max)
    ReDim result(nums)
    result(0) = vbCrLf
    For i = 1 To max
        DoEvents
        f(i) = 30000
    Next
    For j = max To 1 Step -10
        t = 0
        DoEvents
        For i = j To 1 Step -1
            DoEvents
            If j = max Then
                t = t + f(i) * 1000000
            Else
                t = t + f(i) * 100000
            End If
            r = 8 * i * (2 * i + 1)
            f(i) = t - Int(t / r) * r
            d = 2 * i - 1
            t = Int(t / r) * d * d
        Next
        k = k + 1
        result(k) = Format(Int(g + t / 100000) Mod 100000, "00000")
        If k Mod 20 = 0 Then result(k) = result(k) & vbCrLf
        If k Mod 200 = 0 Then result(k) = result(k) & "---[" & k * 5 & "]---" & vbCrLf
        g = t Mod 100000
    Next
    Text1.Text = "3." & Join(result, " ")
    Open "C:\1.txt" For Output As #1
        Print #1, Text1.Text
    Close #1
End Sub

Private Sub Form_Load()
getpi 100000
End Sub
