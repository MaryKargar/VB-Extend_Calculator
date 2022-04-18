VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "www.SourceCodes.ir"
   ClientHeight    =   6045
   ClientLeft      =   4620
   ClientTop       =   3090
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   5790
   Begin VB.CommandButton f_e 
      Caption         =   "F-E"
      Height          =   495
      Left            =   4800
      TabIndex        =   47
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton mod 
      Caption         =   "mod"
      Height          =   495
      Left            =   2640
      TabIndex        =   46
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton a 
      Caption         =   "1"
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   45
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton bs 
      Caption         =   "BackSpace"
      Height          =   495
      Left            =   3600
      TabIndex        =   44
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton int 
      Caption         =   "Int"
      Height          =   495
      Left            =   2880
      TabIndex        =   43
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton c 
      Caption         =   "Not"
      Height          =   495
      Index           =   9
      Left            =   4080
      TabIndex        =   42
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton s 
      Caption         =   "And"
      Height          =   495
      Index           =   6
      Left            =   4080
      TabIndex        =   41
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton s 
      Caption         =   "Or"
      Height          =   495
      Index           =   8
      Left            =   4080
      TabIndex        =   40
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton s 
      Caption         =   "Xor"
      Height          =   495
      Index           =   7
      Left            =   4080
      TabIndex        =   39
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton xp 
      Caption         =   "exp"
      Height          =   495
      Left            =   4200
      TabIndex        =   38
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton tavan_3 
      Caption         =   "x^3"
      Height          =   495
      Left            =   4920
      TabIndex        =   36
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton tavan_2 
      Caption         =   "x^2"
      Height          =   495
      Left            =   4920
      TabIndex        =   35
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton s 
      Caption         =   "x^y"
      Height          =   495
      Index           =   0
      Left            =   4920
      TabIndex        =   34
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton div_1 
      Caption         =   "1/x"
      Height          =   495
      Left            =   480
      TabIndex        =   33
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton pi 
      Caption         =   "pi"
      Height          =   495
      Left            =   3240
      TabIndex        =   32
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton fa 
      Caption         =   "n!"
      Height          =   495
      Left            =   3480
      TabIndex        =   31
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton c 
      Caption         =   "rnd"
      Height          =   495
      Index           =   7
      Left            =   2280
      TabIndex        =   30
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton c 
      Caption         =   "abs"
      Height          =   495
      Index           =   6
      Left            =   1080
      TabIndex        =   29
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton c 
      Caption         =   "tan"
      Height          =   495
      Index           =   4
      Left            =   3000
      TabIndex        =   28
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton c 
      Caption         =   "sqrt"
      Height          =   495
      Index           =   3
      Left            =   1680
      TabIndex        =   27
      Top             =   4680
      Width           =   495
   End
   Begin VB.CommandButton c 
      Caption         =   "cos"
      Height          =   495
      Index           =   0
      Left            =   2400
      TabIndex        =   26
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton c 
      Caption         =   "log"
      Height          =   495
      Index           =   2
      Left            =   3600
      TabIndex        =   25
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton c 
      Caption         =   "sin"
      Height          =   495
      Index           =   1
      Left            =   1800
      TabIndex        =   24
      Top             =   5280
      Width           =   495
   End
   Begin VB.CommandButton M 
      Caption         =   "M+"
      Height          =   495
      Left            =   960
      TabIndex        =   23
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton MS 
      Caption         =   "MS"
      Height          =   495
      Left            =   1560
      TabIndex        =   22
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton MR 
      Caption         =   "MR"
      Height          =   495
      Left            =   2160
      TabIndex        =   21
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton mc 
      Caption         =   "MC"
      Height          =   495
      Left            =   2760
      TabIndex        =   20
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Ce 
      Caption         =   "CE"
      Height          =   495
      Left            =   4920
      TabIndex        =   19
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Clear 
      Caption         =   "C"
      Height          =   495
      Left            =   4920
      TabIndex        =   17
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Cp 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   16
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton eq 
      Caption         =   "="
      Height          =   495
      Left            =   1680
      TabIndex        =   15
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton s 
      Caption         =   "-"
      Height          =   495
      Index           =   5
      Left            =   3240
      TabIndex        =   14
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton s 
      Caption         =   "/"
      Height          =   495
      Index           =   3
      Left            =   3240
      TabIndex        =   13
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton s 
      Caption         =   "%"
      Height          =   495
      Index           =   1
      Left            =   3240
      TabIndex        =   12
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton s 
      Caption         =   "+"
      Height          =   495
      Index           =   4
      Left            =   2640
      TabIndex        =   11
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton s 
      Caption         =   "*"
      Height          =   495
      Index           =   2
      Left            =   2640
      TabIndex        =   10
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton pm 
      Caption         =   "-/+"
      Height          =   495
      Index           =   0
      Left            =   2640
      TabIndex        =   9
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton a 
      Caption         =   "0"
      Default         =   -1  'True
      Height          =   495
      Index           =   9
      Left            =   1080
      TabIndex        =   8
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton a 
      Caption         =   "9"
      Height          =   495
      Index           =   8
      Left            =   1680
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton a 
      Caption         =   "8"
      Height          =   495
      Index           =   7
      Left            =   1080
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton a 
      Caption         =   "7"
      Height          =   495
      Index           =   6
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton a 
      Caption         =   "6"
      Height          =   495
      Index           =   5
      Left            =   1680
      TabIndex        =   4
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton a 
      Caption         =   "5"
      Height          =   495
      Index           =   4
      Left            =   1080
      TabIndex        =   3
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton a 
      Caption         =   "4"
      Height          =   495
      Index           =   3
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton a 
      Caption         =   "3"
      Height          =   495
      Index           =   2
      Left            =   1680
      TabIndex        =   1
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton a 
      Caption         =   "2"
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   0
      Top             =   3240
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00808000&
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   735
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   2655
      Left            =   3960
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label b1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   480
      TabIndex        =   37
      Top             =   1080
      Width           =   375
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   3015
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   5295
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      Height          =   2655
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   2655
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label L 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   600
      TabIndex        =   18
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public p, q, mem, z, n, d As Double
Public flag, sf, fm As Boolean
Public sign, cs, e As String
Public i As Integer
Public f, fi, j As Double

Private Sub a_Click(Index As Integer)

If flag = True Then L.Caption = ""

If L.Caption = "0" Or L.Caption = "error" Then
    L.Caption = ""
End If
L.Caption = L.Caption + a(Index).Caption

flag = False
End Sub

Private Sub bs_Click()
L.Caption = Left(L.Caption, Len(L.Caption) - 1)
If Len(L.Caption) = 0 Then L.Caption = "0"
End Sub

Private Sub C_Click(Index As Integer)
cs = c(Index).Caption
Select Case (cs)
Case "sin"
           L.Caption = (L.Caption * 3.14) / 180
           L.Caption = Math.Sin(Val(L.Caption))
           flag = True
Case "cos"
           L.Caption = (L.Caption * 3.14) / 180
           L.Caption = Math.Cos(Val(L.Caption))
           flag = True
Case "tan"
           L.Caption = (L.Caption * 3.14) / 180
           L.Caption = Math.Tan(Val(L.Caption))
           flag = True
Case "sqrt"
           L.Caption = Math.Sqr(Val(L.Caption))
           flag = True
Case "abs"
           L.Caption = Math.Abs(Val(L.Caption))
           flag = True
Case "rnd"
           L.Caption = Math.Rnd(Val(L.Caption))
           flag = True
Case "log"
           L.Caption = Math.Log(Val(L.Caption))
           flag = True
Case "Not"
           d = Val(L.Caption)
           L.Caption = Not (Val(L.Caption))
           flag = True
End Select
End Sub

Private Sub Ce_Click()
L.Caption = "0"
mem = 0
b1.Caption = ""
End Sub

Private Sub Clear_Click()
L.Caption = "0"
sign = ""
flag = True
sf = False
fm = False
End Sub

Private Sub Cp_Click()
If fm = False Then
If Len(L.Caption) = 0 Then
L.Caption = "0."
Else
L.Caption = L.Caption + "."
End If
Else
Exit Sub
End If

fm = True
End Sub

Private Sub div_1_Click()
If L.Caption = "0" Then
  L.Caption = "error"
Else
  L.Caption = 1 / Val(L.Caption)
End If
End Sub

Private Sub eq_Click()
If (flag = True) Then Exit Sub

q = Val(L.Caption)
Select Case (sign)
   Case "+"
     p = p + q
   
   Case "-"
     p = p - q
   
   Case "*"
     p = p * q
   
   Case "/"
     If q = 0 Then
       L.Caption = "error"
       sf = False
       sign = ""
       Exit Sub
     Else
       p = p / q
     End If
   Case "%"
     p = p * (q / 100)
   
   Case "x^y"
        p = Exp(q * Log(p))
   Case "And"
        p = p And q
     
   Case "Xor"
        p = p Xor q
   
   Case "Or"
        p = p Or q
   
   Case ""
     Exit Sub
End Select

L.Caption = p

sf = False
flag = True
sign = ""
fm = False
End Sub

Private Sub f_e_Click()
n = Val(L.Caption)
For z = 1 To n
  d = n / z
  If n Mod z = 0 Then
    e = e + Str(z)
  End If
Next z
L.Caption = e
End Sub

Private Sub fa_Click()
If Sgn(Val(L.Caption)) = -1 Then
  L.Caption = "error"
  sf = False
  sign = ""
Else
  f = 1
  For fi = Val(L.Caption) To 2 Step -1
     f = f * fi
  Next fi
  L.Caption = f
  flag = True
End If
End Sub

Private Sub Form_Load()
L.Caption = "0"
sign = ""
flag = True
sf = False
End Sub

Private Sub int_Click()
L.Caption = Int(Val(L.Caption))
End Sub



Private Sub Label2_Click()

End Sub

Private Sub M_Click()
mem = mem + Val(L.Caption)
End Sub

Private Sub mc_Click()
mem = 0
b1.Caption = ""
End Sub

Private Sub MR_Click()
L.Caption = mem
End Sub

Private Sub MS_Click()
mem = Val(L.Caption)
b1.Caption = "M"

End Sub



Private Sub pi_Click()
L.Caption = "3.1415926535897932384626433832795"
End Sub

Private Sub pm_Click(Index As Integer)
L.Caption = -Val(L.Caption)
End Sub

Private Sub s_Click(Index As Integer)

If (flag = True And sign <> "") Then Exit Sub

flag = True

If (sf = False) Then
    p = Val(L.Caption)
    sf = True
Else
Select Case (sign)
          Case "+"
                p = p + Val(L.Caption)
   
          Case "-"
                p = p - Val(L.Caption)
   
          Case "*"
                p = p * Val(L.Caption)
   
          Case "/"
                If L.Caption = "0" Then
                   L.Caption = "error"
                     sf = False
                     sign = ""
                     Exit Sub
                Else
                   p = p / Val(L.Caption)
                End If
          
          Case "%"
                p = p * (Val(L.Caption) / 100)
          Case "x^y"
                p = Exp(q * Log(p))
             Case "And"
        p = p And q
     
   Case "Xor"
        p = p Xor q
   
   Case "Or"
        p = p Or q
  
          
          Case ""
                Exit Sub
    End Select
End If
L.Caption = p
sign = s(Index).Caption
End Sub

Private Sub tavan_2_Click()
L.Caption = Val(L.Caption) * Val(L.Caption)
flag = True
End Sub

Private Sub tavan_3_Click()
L.Caption = Val(L.Caption) * Val(L.Caption) * Val(L.Caption)
flag = True
End Sub



Private Sub xp_Click()
L.Caption = Exp(Val(L.Caption))
End Sub
