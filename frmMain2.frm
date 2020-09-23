VERSION 5.00
Begin VB.Form frmMain2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trigonometric Functions Helper - More Trig"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   Icon            =   "frmMain2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdViewTutorial 
      Caption         =   "View the Trigonometry Tutorial File"
      Height          =   975
      Left            =   7680
      TabIndex        =   26
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7680
      TabIndex        =   25
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Last Page"
      Height          =   375
      Left            =   9360
      TabIndex        =   24
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dont Forget!  ;)"
      Height          =   1575
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   7455
      Begin VB.Label lblFrame3Remember 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain2.frx":0ECA
         Height          =   1215
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   4455
      End
      Begin VB.Image Image1 
         Height          =   1200
         Left            =   4680
         Picture         =   "frmMain2.frx":1018
         Top             =   240
         Width           =   2580
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sine Law (sinA/a = sinB/b = sinC/c)  This will use sinA/a = sinB/b  as a guide.  Sub in any values..."
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton cmdSineLawSolveSinA 
         Caption         =   "Solve Angle A"
         Height          =   255
         Left            =   840
         TabIndex        =   21
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtSineLawAnswer 
         Height          =   285
         Left            =   5760
         TabIndex        =   20
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdSineLawSolvea 
         Caption         =   "Solve Side a"
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdSineLawSolveb 
         Caption         =   "Solve Side b"
         Height          =   255
         Left            =   5160
         TabIndex        =   16
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdSineLawSolveSinB 
         Caption         =   "Solve Angle B"
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtSineLawSinA 
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtSineLawa 
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtSineLawSinB 
         Height          =   285
         Left            =   4080
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtSineLawb 
         Height          =   285
         Left            =   4080
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblSineLawWarning 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain2.frx":65CE
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   7215
      End
      Begin VB.Label lblSineLawAnswer 
         BackStyle       =   0  'Transparent
         Caption         =   "Answer:"
         Height          =   255
         Left            =   6240
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.Line Line2 
         X1              =   4080
         X2              =   5640
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter the side length(s) in the bottom text boxes."
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   2160
         X2              =   3720
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lblSineAngles 
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter the Angle(s) in the top text boxes."
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sin-1, Cos-1, Tan-1.  Use deg."
      Height          =   1815
      Left            =   7680
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtInvSinCosTanAnswer 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CommandButton cmdGetInvSin 
         Caption         =   "sin-1"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtInvSinCosTanAngle 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdGetInvCos 
         Caption         =   "cos-1"
         Height          =   255
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdGetInvTan 
         Caption         =   "tan-1"
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblInvSinCosTanAnswer 
         BackStyle       =   0  'Transparent
         Caption         =   "Answer:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmMain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Pi As Double = 3.14159265358979 'pi, easy as pie. ;)
Const RadToDegFuncX As Double = 57.2957795130824 'what you multiply your trig
'180/pi                                          'function answer by.  So:
                                                 'Sin (90) * RadToDegFuncX = the
                                                 'answer in degrees
Const DegToRadAngleX As Double = 1.74532925199433E-02 'what you multiply your
'pi/180                                         'angle by so that you can use it
                                                'in a trig function like Sin/Cos
                                                'Tan.  This will convert the angle
                                                'to radians, then you use it in
                                                'function, and then convert the
                                                'radian answer back to degrees.

Private Sub cmdExit_Click()

Unload Me
Unload frmAbout
Unload frmMain
Unload frmTut
End

End Sub

Private Sub cmdGetInvCos_Click()
On Error GoTo PROC_Err

Dim Number As Double

Number = CDbl(txtInvSinCosTanAngle.Text)

If Number = 0 Then
    txtInvSinCosTanAnswer.Text = "90"
ElseIf Number = 1 Then
    txtInvSinCosTanAnswer.Text = "0"
Else
    txtInvSinCosTanAnswer.Text = CStr(InvCos(Number))
End If


PROC_Exit:
Exit Sub
PROC_Err:
Number = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdGetInvSin_Click()
On Error GoTo PROC_Err

Dim Number As Double

Number = CDbl(txtInvSinCosTanAngle.Text)

If Number = 1 Then
    txtInvSinCosTanAnswer.Text = "90"
Else
    txtInvSinCosTanAnswer.Text = CStr(InvSin(Number))
End If


PROC_Exit:
Exit Sub
PROC_Err:
Number = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdGetInvTan_Click()
Dim Number As Double

Number = CDbl(txtInvSinCosTanAngle.Text)

txtInvSinCosTanAnswer.Text = CStr(InvTan(Number))

End Sub

Private Sub cmdPrevious_Click()

Unload Me
Load frmMain
frmMain.Show

End Sub

Private Sub cmdSineLawSolvea_Click()
On Error GoTo PROC_Err

'variables to hold, convert and undergoe Sin for the two angles
Dim sinA As Double, dblSinA As Double
Dim sinB As Double, dblSinB As Double
Dim sideb As Double

'holds the answer which is converted to degrees
Dim answer As Double

sinA = CDbl(txtSineLawSinA.Text) * DegToRadAngleX 'converts angle A to radians
sinB = CDbl(txtSineLawSinB.Text) * DegToRadAngleX 'converts angle B to radians
sideb = CDbl(txtSineLawb.Text)                    'holds the side length

dblSinA = (Sin(sinA))
dblSinB = (Sin(sinB))

'does the calculations to solve for side a
answer = CDbl((dblSinA * sideb) / (dblSinB))

'returns the answer
txtSineLawAnswer.Text = CStr(answer)


PROC_Exit:
Exit Sub
PROC_Err:
answer = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdSineLawSolveb_Click()
On Error GoTo PROC_Err

'variables to hold, convert and undergoe Sin for the two angles
Dim sinA As Double, dblSinA As Double
Dim sinB As Double, dblSinB As Double
Dim sidea As Double

'holds the answer which is converted to degrees
Dim answer As Double

sinA = CDbl(txtSineLawSinA.Text) * DegToRadAngleX 'converts angle A to radians
sinB = CDbl(txtSineLawSinB.Text) * DegToRadAngleX 'converts angle B to radians
sidea = CDbl(txtSineLawa.Text)                    'holds the side length

dblSinA = (Sin(sinA))
dblSinB = (Sin(sinB))

'does the calculations to solve for side b
answer = CDbl((dblSinB * sidea) / (dblSinA))

'returns the answer
txtSineLawAnswer.Text = CStr(answer)


PROC_Exit:
Exit Sub
PROC_Err:
answer = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdSineLawSolveSinA_Click()
On Error GoTo PROC_Err

Dim sinB As Double, dblSinB As Double
Dim sidea As Double, sideb As Double
Dim Number As Double, answer As Double

sinB = CDbl(txtSineLawSinB.Text) * DegToRadAngleX
sidea = CDbl(txtSineLawa.Text)
sideb = CDbl(txtSineLawb.Text)

dblSinB = (Sin(sinB))

Number = CDbl((dblSinB * sidea) / (sideb))

answer = CDbl(InvSin(Number))

txtSineLawAnswer.Text = CStr(answer)


PROC_Exit:
Exit Sub
PROC_Err:
answer = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdSineLawSolveSinB_Click()
On Error GoTo PROC_Err

Dim sinA As Double, dblSinA As Double
Dim sidea As Double, sideb As Double
Dim Number As Double, answer As Double

sinA = CDbl(txtSineLawSinA.Text) * DegToRadAngleX
sidea = CDbl(txtSineLawa.Text)
sideb = CDbl(txtSineLawb.Text)

dblSinA = (Sin(sinA))

Number = CDbl((dblSinA * sideb) / (sidea))

answer = CDbl(InvSin(Number))

txtSineLawAnswer.Text = CStr(answer)


PROC_Exit:
Exit Sub
PROC_Err:
answer = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Sub

Private Sub cmdViewTutorial_Click()

Unload Me
Load frmTut
frmTut.Show

End Sub
