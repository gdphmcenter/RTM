VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   8325
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Txt2 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Txt1 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "空间差分格式f="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Label Lbl1 
      AutoSize        =   -1  'True
      Caption         =   "近似阶数："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Base 0
Dim WL() As Double, WN() As Double

Private Sub Command1_Click()
Dim i, j, k, ii, kk As Integer
Dim M As Integer
Dim T1 As Double, T2 As Double
Dim T3 As Double, T4 As Double, T5 As Double, T6 As Double, T7 As Double
Dim E1 As Single, E2 As Single
Dim Pi As Double
Dim Bo As Double
Dim delta As Double
Dim Error As Double, ErrorT As Double, ErrorTemM As Double
Dim Res As Double, ResT As Double, Imax As Double, Tmax As Double
Dim f As Single
Dim u(1 To 8) As Double, w(1 To 8) As Double
Dim ka As Double
Dim N As Integer
Dim NK As Integer
Dim Sum As Double, Sum1 As Double, Sum2 As Double
Dim Sigam As Double
Dim F1 As Double, F2 As Double
Dim uu1 As Double, uu2 As Double
Dim TT As Double
Dim Iter As Integer, Iter1 As Integer, Iter2 As Integer
Dim kavg() As Double
Dim Davg() As Double
Dim AbspMediam(1 To 999) As Double
Dim CaluT As Double
Dim TemM_End() As Double, TemM_Parallel() As Double, TemM_End_R() As Double, TemM_Parallel_R() As Double, L_End() As Double, L_Parallel() As Double, TemM As Double, TemMR As Double
Dim Tcal As Double
Dim Tg As Double
Dim L As Double
Dim XH2O As Double, XCO As Double, XCO2 As Double
Dim Intensity(0 To 1000, 1 To 8) As Double, Intensity0(0 To 1000, 1 To 8) As Double
Dim T0(0 To 1000) As Double, T(0 To 1000) As Double
Dim G(0 To 1000) As Double, E(0 To 1000) As Double
Dim Iupper(1 To 8) As Double, Idown(1 To 8) As Double

Pi = 3.1415926535
Bo = 5.6703 * 10 ^ (-8)
E1 = 1
E2 = 1
T2 = 1030 + 273
Tg = 1041 + 273
f = 0.5

Dim xlConn As New ADODB.Connection
Dim xlRs As New ADODB.Recordset
Dim strConn As String
Dim xlCnt As Single, xlCnt1 As Single

 strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\SNB_H2O.xls;Extended Properties='Excel 8.0;HDR=yes;IMEX=1'"
 xlConn.Open strConn
 xlRs.Open "select * from [k$]", xlConn, adOpenStatic, adLockReadOnly
 xlCnt = xlRs.RecordCount
 ReDim WL(xlCnt), WN(xlCnt), kavg(1 To xlCnt)
 For i = 1 To xlCnt
     WN(i) = xlRs("波数")
     WL(i) = 10000 / xlRs("波数")
     kavg(i) = (Tg - 1300) * (xlRs("1400K") - xlRs("1300K")) / 100 + xlRs("1300K")
     xlRs.MoveNext
Next i
 xlRs.Close
 
 xlRs.Open "select * from [delta$]", xlConn, adOpenStatic, adLockReadOnly
 xlCnt = xlRs.RecordCount
 ReDim Davg(1 To xlCnt)
 For i = 1 To xlCnt
     Davg(i) = (Tg - 1300) * (xlRs("1400K") - xlRs("1300K")) / 100 + xlRs("1300K")
     xlRs.MoveNext
Next i
 xlRs.Close
 xlConn.Close
 
 strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Peter.xls;Extended Properties='Excel 8.0;HDR=yes;IMEX=1'"
 xlConn.Open strConn
 xlRs.Open "select * from [Sheet1$]", xlConn, adOpenStatic, adLockReadOnly
 xlCnt1 = xlRs.RecordCount
 ReDim TemM_End(1 To xlCnt1), TemM_Parallel(1 To xlCnt1), TemM_End_R(1 To xlCnt1), TemM_Parallel_R(1 To xlCnt1), L_End(1 To xlCnt1), L_Parallel(1 To xlCnt1)
 For i = 1 To xlCnt1

     TemM_End(i) = xlRs("End wall") + 273
     TemM_End_R(i) = xlRs("End_Revised") + 273
     TemM_Parallel(i) = xlRs("Parallel wall") + 273
     TemM_Parallel_R(i) = xlRs("Parallel_Revised") + 273
     L_End(i) = xlRs("L_End") * 100
     L_Parallel(i) = xlRs("L_Parallel") * 100
     xlRs.MoveNext
     
Next i

 xlRs.Close
 xlConn.Close
 Set xlRs = Nothing
 Set xlConn = Nothing

N = 8
NK = 1
u(1) = 0.1422555: u(2) = 0.5773503: u(3) = 0.8040087: u(4) = 0.9795543
u(5) = -0.1422555: u(6) = -0.5773503: u(7) = -0.8040087: u(8) = -0.9795543
w(1) = 2.1637144: w(2) = 2.6406988: w(3) = 0.7938272: w(4) = 0.6849436
w(5) = 2.1637144: w(6) = 2.6406988: w(7) = 0.7938272: w(8) = 0.6849436
Open App.Path & "\Absorp.txt" For Output As #3
Open App.Path & "\Tem.txt" For Output As #4


Dim Ravg As Double, Ravg1 As Double, Ravg2 As Double
Dim Absp(1 To 449) As Double

Dim X As Integer
Dim C As Single

C = 0.8

For ii = 1 To 31


TemM = 1.01 * TemM_End(ii)
TemMR = TemM_End_R(ii)
L = L_End(ii)

delta = L / 500

For X = 10 To 10 Step 1
    XH2O = X * 0.1
    Ravg = 0.462 * XH2O * 296 / Tg + 0.0792 * (296 / Tg) ^ 0.5
    ka = 0
    For k = 380 To 420
        Absp(k) = 2 * Ravg * Davg(k) * ((1 + 1.9 * L * XH2O * kavg(k) / Ravg / Davg(k)) ^ 0.5 - 1) / 1.9 / L
        ka = ka + Absp(k)
    Next k
    ka = ka / (420 - 380 + 1)
For j = 1 To N Step 1
    Intensity0(1000, j) = 0
    Intensity0(0, j) = 0
Next j
For i = 1 To 999 Step 1
    T0(i) = Tg
Next i

T1 = Tg
ErrorTemM = 1
Do While ErrorTemM > 0.1
            Iter = 0
            Error = 100
            Do While Error > 10 ^ -6
                Sum1 = 0: Sum2 = 0
                For j = 1 To N Step 1
                    If u(j) > 0 Then
                        Sum2 = Sum2 + w(j) * u(j) * Intensity0(1000, j)
                    ElseIf u(j) < 0 Then
                        Sum1 = Sum1 + w(j) * u(j) * Intensity0(0, j)
                    End If
                Next j

                For j = 1 To N Step 1
                    If u(j) > 0 Then
                        Intensity(0, j) = E1 * Function1(T1) + (1 - E1) * Abs(Sum1) / Pi
                    ElseIf u(j) < 0 Then
                        Intensity(1000, j) = E2 * Function1(T2) + (1 - E2) * Sum2 / Pi
                    End If
                Next j

                j = 1
                Do While j <= N
                    If u(j) > 0 Then
                        i = 1
                        Iupper(j) = Intensity(0, j)
                        Do While i < 1000
                            Intensity(i, j) = (Abs(u(j)) * Iupper(j) + f * ka * delta * Function1(T0(i))) / (Abs(u(j)) + f * ka * delta)
                            If Intensity(i, j) < 0 Then Intensity(i, j) = f * ka * delta * Function1(T0(i)) / (Abs(u(j)) + f * ka * delta)
                            Idown(j) = (Intensity(i, j) - (1 - f) * Iupper(j)) / f
                            If Idown(j) < 0 Then Idown(j) = 0
                            Intensity(i - 1, j) = Iupper(j)
                            Intensity(i + 1, j) = Idown(j)
                            Iupper(j) = Idown(j)
                            i = i + 2
                        Loop
                    ElseIf u(j) < 0 Then
                        i = 999
                        Iupper(j) = Intensity(1000, j)
                        Do While i > 0
                            Intensity(i, j) = (Abs(u(j)) * Iupper(j) + f * ka * delta * Function1(T0(i))) / (Abs(u(j)) + f * ka * delta)
                            If Intensity(i, j) < 0 Then Intensity(i, j) = f * ka * delta * Function1(T0(i)) / (Abs(u(j)) + f * ka * delta)
                            Idown(j) = (Intensity(i, j) - (1 - f) * Iupper(j)) / f
                            If Idown(j) < 0 Then Idown(j) = 0
                            Intensity(i + 1, j) = Iupper(j)
                            Intensity(i - 1, j) = Idown(j)
                            Iupper(j) = Idown(j)
                            i = i - 2
                        Loop
                    End If
                    j = j + 1
                Loop

            Iter = Iter + 1
            Res = 0
            Imax = 0
            For i = 1 To 999 Step 2
                For j = 1 To N
                    If Intensity(i, j) > Imax Then Imax = Intensity(i, j)
                    If Abs(Intensity(i, j) - Intensity0(i, j)) > Res Then Res = Abs(Intensity(i, j) - Intensity0(i, j))
                Next j
            Next

            Error = Res / Imax

            For i = 0 To 1000 Step 1
                For j = 1 To N Step 1
                   Intensity0(i, j) = Intensity(i, j)
                Next j
            Next i
            Loop

        E(999) = 0
        For j = 1 To N Step 1
            If u(j) > 0 Then E(999) = E(999) + w(j) * Intensity(999, j) * u(j)
        Next j
        T3 = 100
        T4 = 2000
        Error = 100
        Do While Error > 10 ^ -2
            T5 = (T3 + T4) / 2
            Sum = Abs(E(999) - Pi * Function1(T5))
            If (T4 - T3) / 2 < 10 ^ -6 Or Sum < 10 ^ -8 Then
                Exit Do
            End If
            If (E(999) - Pi * Function1(T5)) * (E(999) - Pi * Function1(T3)) < 0 Then
                T4 = T5
            Else
                T3 = T5
            End If
        Loop
        Tcal = T5
        ErrorTemM = Abs(TemM - Tcal)
        T1 = T1 + C * (TemM - Tcal)
    Loop
    Print #4, Format(TemM - 273, "0.0000"); Spc(4); Format(TemMR - 273, "0.0000"); Spc(4); Format(T1 - 273, "0.0000"); Spc(4)
Next X
Close #3
Next ii
Close #4
End Sub
Public Function Function1(ByVal Tem As Double) As Double
Dim Bo As Double
Dim Pi As Double
Dim NN As Integer
Dim u1 As Double, u2 As Double
Dim Fu As Double, Fl As Double, Sum As Double
Pi = 3.1415926535
Bo = 5.6703 * 10 ^ (-8)
u1 = 14388 / 1.05 / Tem: u2 = 14388 / 0.95 / Tem
Sum = 0
For NN = 1 To 10 Step 1
    Sum = Sum + (u1 ^ 3 + 3 * u1 ^ 2 / NN + 6 * u1 / NN ^ 2 + 6 / NN ^ 3) * Exp(-u1 * NN) / NN
Next NN
Fu = 15 * Sum / Pi ^ 4

Sum = 0
For NN = 1 To 10 Step 1
    Sum = Sum + (u2 ^ 3 + 3 * u2 ^ 2 / NN + 6 * u2 / NN ^ 2 + 6 / NN ^ 3) * Exp(-u2 * NN) / NN
Next NN
Fl = 15 * Sum / Pi ^ 4

Sum = Bo * Tem ^ 4 * (Fu - Fl)
Function1 = Bo * Tem ^ 4 * (Fu - Fl) / Pi
End Function
