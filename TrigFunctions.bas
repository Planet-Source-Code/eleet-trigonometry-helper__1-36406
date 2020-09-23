Attribute VB_Name = "TrigFunctions"
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

Public Function DegAngleToRadAngle(dblDegAngle As Double) As Double
On Error GoTo PROC_Err

Dim dblRadAngle As Double

dblRadAngle = dblDegAngle * DegToRadAngleX

DegAngleToRadAngle = dblRadAngle


PROC_Exit:
Exit Function
PROC_Err:
DegAngleToRadAngle = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function SineX(ByVal angle As Double) As Double
On Error GoTo PROC_Err

Dim RadAngle As Double

'converts your angle (angle) to radians (RadAngle)
RadAngle = angle * DegToRadAngleX

SineX = (Sin(RadAngle))


PROC_Exit:
Exit Function
PROC_Err:
SineX = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function CosineX(ByVal angle As Double) As Double
On Error GoTo PROC_Err

Dim RadAngle As Double

'converts your angle (angle) to radians (RadAngle)
RadAngle = angle * DegToRadAngleX

CosineX = (Cos(RadAngle))


PROC_Exit:
Exit Function
PROC_Err:
CosineX = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function TangentX(ByVal angle As Double) As Double
On Error GoTo PROC_Err

Dim RadAngle As Double

'converts your angle (angle) to radians (RadAngle)
RadAngle = angle * DegToRadAngleX

TangentX = (Tan(RadAngle))


PROC_Exit:
Exit Function
PROC_Err:
TangentX = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function SineSolveOpp(ByVal theta As Double, ByVal hyp As Double) As Double
On Error GoTo PROC_Err

Dim RadTheta As Double, opp As Double

'converts your angle (theta) to radians (RadTheta)
RadTheta = theta * DegToRadAngleX

'calculates your function    'sin(angle) = opp/hyp'
opp = (Sin(RadTheta)) * hyp

'sets the value to return
SineSolveOpp = opp


PROC_Exit:
Exit Function
PROC_Err:
SineSolveOpp = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function SineSolveHyp(ByVal theta As Double, ByVal opp As Double) As Double
On Error GoTo PROC_Err

Dim RadTheta As Double, hyp As Double

'converts your angle (theta) to radians (RadTheta)
RadTheta = theta * DegToRadAngleX

'calculates your function    'sin(angle) = opp/hyp'
hyp = opp / (Sin(RadTheta))

'sets the value to return
SineSolveHyp = hyp


PROC_Exit:
Exit Function
PROC_Err:
SineSolveHyp = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function InvSinForTheta(ByVal Number As Double) As Double
On Error GoTo PROC_Err

Dim theta As Double

theta = Atn(Number / Sqr(-Number * Number + 1)) * RadToDegFuncX

InvSinForTheta = theta


PROC_Exit:
Exit Function
PROC_Err:
InvSinForTheta = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function CosineSolveHyp(ByVal theta As Double, ByVal adj As Double) As Double
On Error GoTo PROC_Err

Dim RadTheta As Double, hyp As Double

'converts your angle (theta) to radians (RadTheta)
RadTheta = theta * DegToRadAngleX

'calculates your function    'sin(angle) = opp/hyp'
hyp = adj / (Cos(RadTheta))

'sets the value to return
CosineSolveHyp = hyp


PROC_Exit:
Exit Function
PROC_Err:
CosineSolveHyp = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function CosineSolveAdj(ByVal theta As Double, ByVal hyp As Double) As Double
On Error GoTo PROC_Err

Dim RadTheta As Double, adj As Double

'converts your angle (theta) to radians (RadTheta)
RadTheta = theta * DegToRadAngleX

'calculates your function    'sin(angle) = opp/hyp'
adj = (Cos(RadTheta)) * hyp

'sets the value to return
CosineSolveAdj = adj


PROC_Exit:
Exit Function
PROC_Err:
CosineSolveAdj = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function InvCosForTheta(ByVal Number As Double) As Double
On Error GoTo PROC_Err

Dim theta As Double

theta = (Atn(-Number / Sqr(-Number * Number + 1)) * RadToDegFuncX) + (2 * Atn(1) * RadToDegFuncX)

InvCosForTheta = theta


PROC_Exit:
Exit Function
PROC_Err:
InvCosForTheta = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function TangentSolveAdj(ByVal theta As Double, ByVal opp As Double) As Double
On Error GoTo PROC_Err

Dim RadTheta As Double, adj As Double

'converts your angle (theta) to radians (RadTheta)
RadTheta = theta * DegToRadAngleX

'calculates your function    'tan(angle) = opp/adj'
adj = opp / (Tan(RadTheta))

'sets the value to return
TangentSolveAdj = adj


PROC_Exit:
Exit Function
PROC_Err:
TangentSolveAdj = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function TangentSolveOpp(ByVal theta As Double, ByVal adj As Double) As Double
On Error GoTo PROC_Err

Dim RadTheta As Double, opp As Double

'converts your angle (theta) to radians (RadTheta)
RadTheta = theta * DegToRadAngleX

'calculates your function    'tan(angle) = opp/adj'
opp = (Tan(RadTheta)) * adj

'sets the value to return
TangentSolveOpp = opp


PROC_Exit:
Exit Function
PROC_Err:
TangentSolveOpp = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function InvTanForTheta(ByVal Number As Double) As Double
On Error GoTo PROC_Err

Dim theta As Double

theta = Atn(Number) * RadToDegFuncX

InvTanForTheta = theta


PROC_Exit:
Exit Function
PROC_Err:
InvTanForTheta = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function InvSin(Number As Double) As Double
On Error GoTo PROC_Err

Dim answer As Double

answer = Atn(Number / Sqr(-Number * Number + 1)) * RadToDegFuncX

InvSin = answer


PROC_Exit:
Exit Function
PROC_Err:
InvSin = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function InvCos(Number As Double) As Double
On Error GoTo PROC_Err

Dim answer As Double

answer = (Atn(-Number / Sqr(-Number * Number + 1)) * RadToDegFuncX) + (2 * Atn(1) * RadToDegFuncX)

InvCos = answer


PROC_Exit:
Exit Function
PROC_Err:
InvCos = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function

Public Function InvTan(Number As Double) As Double
On Error GoTo PROC_Err

Dim answer As Double

answer = Atn(Number) * RadToDegFuncX

InvTan = answer


PROC_Exit:
Exit Function
PROC_Err:
InvTan = 0
MsgBox Err.Description, vbExclamation
Resume PROC_Exit

End Function



