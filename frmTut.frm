VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmTut 
   Caption         =   "Trigonometry Tutorial:  :)"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   Icon            =   "frmTut.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back to the Trig Program"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   4680
      Width           =   4695
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      ExtentX         =   14208
      ExtentY         =   8070
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmTut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()

Unload Me
Load frmMain
frmMain.Show

End Sub

Private Sub Form_Load()

'===========================================================================
'as long as for the Form_Resize() is in this form, all the contents in these
'markers can be disregarded, as it makes no difference
WebBrowser1.Width = frmTut.ScaleWidth

If frmTut.ScaleHeight <= 495 Then
    WebBrowser1.Height = frmTut.ScaleHeight
Else
    WebBrowser1.Height = frmTut.ScaleHeight - 495
End If


If frmTut.ScaleHeight <= 495 Then
    cmdBack.Visible = False
Else
    cmdBack.Visible = True
End If

cmdBack.Top = frmTut.ScaleHeight - 495
cmdBack.Width = frmTut.ScaleWidth
'see the above comment...
'===========================================================================

WebBrowser1.Navigate (App.Path + "\" + "tutorial.htm")

End Sub

Private Sub Form_Resize()

WebBrowser1.Width = frmTut.ScaleWidth

If frmTut.ScaleHeight <= 495 Then
    WebBrowser1.Height = frmTut.ScaleHeight
Else
    WebBrowser1.Height = frmTut.ScaleHeight - 495
End If


If frmTut.ScaleHeight <= 495 Then
    cmdBack.Visible = False
Else
    cmdBack.Visible = True
End If

cmdBack.Top = frmTut.ScaleHeight - 495
cmdBack.Width = frmTut.ScaleWidth

End Sub
