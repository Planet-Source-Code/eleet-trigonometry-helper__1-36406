VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"frmMain.frx":0000
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10425
   Icon            =   "frmMain.frx":0089
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNextPage 
      Caption         =   "Next Page"
      Height          =   375
      Left            =   9000
      TabIndex        =   55
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6840
      TabIndex        =   54
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame Frame5 
      Caption         =   "Tangent Functions (Double Precision)"
      Height          =   1575
      Left            =   120
      TabIndex        =   42
      Top             =   3960
      Width           =   6615
      Begin VB.TextBox txtTangentOpp 
         Height          =   285
         Left            =   1680
         TabIndex        =   48
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtTangentAdj 
         Height          =   285
         Left            =   1680
         TabIndex        =   47
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton cmdTangentSolveTheta 
         Caption         =   "Solve for Theta"
         Height          =   255
         Left            =   3360
         TabIndex        =   46
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdTangentSolveOpp 
         Caption         =   "Solve for Opp."
         Height          =   255
         Left            =   3360
         TabIndex        =   45
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdTangentSolveAdj 
         Caption         =   "Solve for Adj."
         Height          =   255
         Left            =   3360
         TabIndex        =   44
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtTangentTheta 
         Height          =   285
         Left            =   1680
         TabIndex        =   43
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblTangentTheta 
         BackStyle       =   0  'Transparent
         Caption         =   "Theta (angle in degrees)"
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblTanThetaSymbol 
         BackStyle       =   0  'Transparent
         Caption         =   "q"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1320
         TabIndex        =   52
         Top             =   480
         Width           =   150
      End
      Begin VB.Label lblTangentOpp 
         Caption         =   "Opposite (length)"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblTangentAdj 
         Caption         =   "Adjacent (length)"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblTangentRemeber 
         BackStyle       =   0  'Transparent
         Caption         =   "Remember:"
         Height          =   255
         Left            =   5280
         TabIndex        =   49
         Top             =   360
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   690
         Left            =   5040
         Picture         =   "frmMain.frx":0F53
         Top             =   600
         Width           =   1230
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cosine Functions (Double Precision)"
      Height          =   2655
      Left            =   6840
      TabIndex        =   30
      Top             =   1560
      Width           =   3495
      Begin VB.TextBox txtCosineAdj 
         Height          =   285
         Left            =   1680
         TabIndex        =   36
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtCosineHyp 
         Height          =   285
         Left            =   1680
         TabIndex        =   35
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton cmdCosineSolveTheta 
         Caption         =   "Solve for Theta"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdCosineSolveAdj 
         Caption         =   "Solve for Adj."
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton cmdCosineSolveHyp 
         Caption         =   "Solve for Hyp."
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtCosineTheta 
         Height          =   285
         Left            =   1680
         TabIndex        =   31
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblCosineTheta 
         BackStyle       =   0  'Transparent
         Caption         =   "Theta (angle in degrees)"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblCosThetaSymbol 
         BackStyle       =   0  'Transparent
         Caption         =   "q"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1320
         TabIndex        =   40
         Top             =   480
         Width           =   150
      End
      Begin VB.Label lblCosineAdj 
         Caption         =   "Adacent (length)"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblCosineHyp 
         Caption         =   "Hypotenuse (length)"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblCosineRemember 
         BackStyle       =   0  'Transparent
         Caption         =   "Remember:"
         Height          =   255
         Left            =   2040
         TabIndex        =   37
         Top             =   1560
         Width           =   855
      End
      Begin VB.Image imgCosine 
         Height          =   690
         Left            =   1800
         Picture         =   "frmMain.frx":48DF
         Top             =   1800
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sine, Cosine, Tangent.  In/output as degrees"
      Height          =   1335
      Left            =   6840
      TabIndex        =   22
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmdGetTan 
         Caption         =   "tan"
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdGetCos 
         Caption         =   "cos"
         Height          =   255
         Left            =   1560
         TabIndex        =   26
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtSinCosTanAngle 
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   25
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdGetSin 
         Caption         =   "sin"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtSinCosTanAnswer 
         Height          =   285
         Left            =   720
         TabIndex        =   23
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblSinCosTanAnswer 
         BackStyle       =   0  'Transparent
         Caption         =   "Answer"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sine Functions (Double Precision)"
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   6615
      Begin VB.TextBox txtSineTheta 
         Height          =   285
         Left            =   1680
         TabIndex        =   21
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdSineSolveHyp 
         Caption         =   "Solve for Hyp."
         Height          =   255
         Left            =   3360
         TabIndex        =   20
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdSineSolveOpp 
         Caption         =   "Solve for Opp."
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdSineSolveTheta 
         Caption         =   "Solve for Theta"
         Height          =   255
         Left            =   3360
         TabIndex        =   18
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtSineHyp 
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtSineOpp 
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Top             =   840
         Width           =   1575
      End
      Begin VB.Image imgSine 
         Height          =   960
         Left            =   4800
         Picture         =   "frmMain.frx":833D
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblSineRemember 
         BackStyle       =   0  'Transparent
         Caption         =   "Remember:"
         Height          =   255
         Left            =   5280
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblSineHyp 
         Caption         =   "Hypotenuse (length)"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblSineOpp 
         Caption         =   "Opposite (length)"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblTheta 
         BackStyle       =   0  'Transparent
         Caption         =   "q"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1320
         TabIndex        =   12
         Top             =   480
         Width           =   150
      End
      Begin VB.Label lblSineTheta 
         BackStyle       =   0  'Transparent
         Caption         =   "Theta (angle in degrees)"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pi, (80/pi) and (pi/180) Values (Everything in Double Precision)"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtPiOver180 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton cmdPiOver180 
         Caption         =   "(pi/180)"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt180OverPi 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton cmd180OverPi 
         Caption         =   "(180/pi)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtPi 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton cmdGetPi 
         Caption         =   "Get Pi"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblFrame13Desc 
         BackStyle       =   0  'Transparent
         Caption         =   "What you use to convert your angle in Degrees to an angle in Radians."
         Height          =   495
         Left            =   3120
         TabIndex        =   10
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label lblFrame11Desc 
         BackStyle       =   0  'Transparent
         Caption         =   "Pi.  Easy as pie. ;)"
         Height          =   255
         Left            =   3120
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblFrame12Desc 
         BackStyle       =   0  'Transparent
         Caption         =   "What you use to convert you answer from VB's functions from Radians to Degrees."
         Height          =   495
         Left            =   3120
         TabIndex        =   8
         Top             =   840
         Width           =   3375
      End
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "I assume you know how to use these functions, so if you put the wrong numbers in, the program will crash."
      Height          =   615
      Left            =   6840
      TabIndex        =   29
      Top             =   4320
      Width           =   3375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Const Name as Type = Value
Const Pi As Double = 3.14159265358979 'pi.  Easy as pie. ;)
Const RadToDegFuncX As Double = 57.2957795130824 'what you multiply your trig
                                                 'function answer by.  So:
                                                 'Sin (90) * RadToDegFuncX = the
                                                 'answer in degrees
Const DegToRadAngleX As Double = 1.74532925199433E-02 'what you multiply your
                                                'angle by so that you can use it
                                                'in a trig function like Sin/Cos
                                                'Tan.  This will convert the angle
                                                'to radians, then you use it in
                                                'function, and then convert the
                                                'radian answer back to degrees.

Private Sub cmd180OverPi_Click()

txt180OverPi.Text = CStr(RadToDegFuncX)

End Sub

Private Sub cmdCosineSolveAdj_Click()
On Error GoTo PROC_Err

Dim theta As Double, hyp As Double, adjacent As Double

theta = CDbl(txtCosineTheta.Text)
hyp = CDbl(txtCosineHyp.Text)

adjacent = CosineSolveAdj(theta, hyp)

txtCosineAdj.Text = CStr(adjacent)


PROC_Exit:
Exit Sub
PROC_Err:
adjacent = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdCosineSolveHyp_Click()
On Error GoTo PROC_Err

Dim theta As Double, hypotenuse As Double, adj As Double

theta = CDbl(txtCosineTheta.Text)
adj = CDbl(txtCosineAdj.Text)

hypotenuse = CosineSolveHyp(theta, adj)

txtCosineHyp.Text = CStr(hypotenuse)


PROC_Exit:
Exit Sub
PROC_Err:
hypotenuse = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdCosineSolveTheta_Click()
On Error GoTo PROC_Err

Dim hyp As Double, adj As Double
Dim Number As Double, theta As Double

'declares the values of opposite and hypotenuse sides as double-precision
adj = CDbl(txtCosineAdj.Text)
hyp = CDbl(txtCosineHyp.Text)

'declares the number to be applied to InvSinForTheta  (Cos-1(adj/hyp))
Number = adj / hyp

theta = InvCosForTheta(Number)

txtCosineTheta.Text = CStr(theta)


PROC_Exit:
Exit Sub
PROC_Err:
theta = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdGetCos_Click()
On Error GoTo PROC_Err

txtSinCosTanAnswer.Text = CStr(CosineX(CDbl(txtSinCosTanAngle.Text)))


PROC_Exit:
Exit Sub
PROC_Err:
txtSinCosTanAngle.Text = "0"
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdGetPi_Click()

txtPi.Text = CStr(Pi)

End Sub

Private Sub cmdGetSin_Click()
On Error GoTo PROC_Err

txtSinCosTanAnswer.Text = CStr(SineX(CDbl(txtSinCosTanAngle.Text)))


PROC_Exit:
Exit Sub
PROC_Err:
txtSinCosTanAngle.Text = "0"
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdGetTan_Click()
On Error GoTo PROC_Err

txtSinCosTanAnswer.Text = CStr(TangentX(CDbl(txtSinCosTanAngle.Text)))


PROC_Exit:
Exit Sub
PROC_Err:
txtSinCosTanAngle.Text = "0"
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdNextPage_Click()

Unload Me
Unload frmAbout
Load frmMain2
frmMain2.Show

End Sub

Private Sub cmdPiOver180_Click()

txtPiOver180.Text = CStr(DegToRadAngleX)

End Sub

Private Sub cmdQuit_Click()

Unload Me
Unload frmAbout
Unload frmMain
Unload frmTut
End

End Sub

Private Sub cmdSineSolveHyp_Click()
On Error GoTo PROC_Err

Dim theta As Double, hypotenuse As Double, opp As Double

theta = CDbl(txtSineTheta.Text)
opp = CDbl(txtSineOpp.Text)

hypotenuse = SineSolveHyp(theta, opp)

txtSineHyp.Text = CStr(hypotenuse)


PROC_Exit:
Exit Sub
PROC_Err:
hypotenuse = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdSineSolveOpp_Click()
On Error GoTo PROC_Err

Dim theta As Double, hyp As Double, opposite As Double

theta = CDbl(txtSineTheta.Text)
hyp = CDbl(txtSineHyp.Text)

opposite = SineSolveOpp(theta, hyp)

txtSineOpp.Text = CStr(opposite)


PROC_Exit:
Exit Sub
PROC_Err:
opposite = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdSineSolveTheta_Click()
On Error GoTo PROC_Err

Dim hyp As Double, opp As Double
Dim Number As Double, theta As Double

'declares the values of opposite and hypotenuse sides as double-precision
opp = CDbl(txtSineOpp.Text)
hyp = CDbl(txtSineHyp.Text)

'declares the number to be applied to InvSinForTheta  (Sin-1(opp/hyp))
Number = opp / hyp

theta = InvSinForTheta(Number)

txtSineTheta.Text = CStr(theta)


PROC_Exit:
Exit Sub
PROC_Err:
theta = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdTangentSolveAdj_Click()
On Error GoTo PROC_Err

Dim theta As Double, opp As Double, adjacent As Double

theta = CDbl(txtTangentTheta.Text)
opp = CDbl(txtTangentOpp.Text)

adjacent = TangentSolveAdj(theta, opp)

txtTangentAdj.Text = CStr(adjacent)


PROC_Exit:
Exit Sub
PROC_Err:
adjacent = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdTangentSolveOpp_Click()
On Error GoTo PROC_Err

Dim theta As Double, adj As Double, opposite As Double

theta = CDbl(txtTangentTheta.Text)
adj = CDbl(txtTangentAdj.Text)

opposite = TangentSolveOpp(theta, adj)

txtTangentOpp.Text = CStr(opposite)


PROC_Exit:
Exit Sub
PROC_Err:
opposite = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdTangentSolveTheta_Click()
On Error GoTo PROC_Err

Dim adj As Double, opp As Double
Dim Number As Double, theta As Double

'declares the values of opposite and hypotenuse sides as double-precision
opp = CDbl(txtTangentOpp.Text)
adj = CDbl(txtTangentAdj.Text)

'declares the number to be applied to InvSinForTheta  (Tan-1(opp/adj))
Number = opp / adj

theta = InvTanForTheta(Number)

txtTangentTheta.Text = CStr(theta)


PROC_Exit:
Exit Sub
PROC_Err:
theta = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub mnuFileExit_Click()

Call cmdQuit_Click

End Sub

Private Sub mnuHelpAbout_Click()

Me.Hide
Load frmAbout
frmAbout.Show

End Sub
