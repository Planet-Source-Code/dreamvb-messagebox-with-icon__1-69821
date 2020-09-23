VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "MessageBox - Custom Icon"
   ClientHeight    =   930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   930
   ScaleWidth      =   4245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2910
      TabIndex        =   2
      Top             =   300
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   300
      Width           =   1215
   End
   Begin VB.CommandButton cmdFormat 
      Caption         =   "Format A:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function InIDE() As Boolean
On Error GoTo IdeErr:
    'Check if running in the IDE
    Debug.Print (1 \ 0)
    'Compiled app.
    InIDE = False
    Exit Function
IdeErr:
    'Running in the IDE
    InIDE = True
End Function
Private Sub cmdAbout_Click()
Dim Ret As VbMsgBoxResult
    'Show About box.
    'Using normal VB Style icon
    Ret = VbMsgBox(frmMain.Caption & vbCrLf & vbTab & "By DreamVB", vbOKOnly Or vbInformation, "About")
    
    If (Ret = vbYes) Then
        'Yes they do so exit.
        Unload frmMain
    End If
    
End Sub

Private Sub cmdExit_Click()
Dim Ret As VbMsgBoxResult
    'Using custom icon resID 102
    'Ask user does they want to exit.
    Ret = VbMsgBox("Do you want to Exit.", vbYesNo, "Exit", True, 102)
    
    If (Ret = vbYes) Then
        'Yes they do so exit.
        Ret = VbMsgBox("Please Vote if you like the code.", vbOKOnly, , True, 103)
        Unload frmMain
    End If
    
End Sub

Private Sub cmdFormat_Click()
Dim Ret As VbMsgBoxResult
    'Using custom icon resID 101
    'Ask user does they want to format Drive A.
Top:
    Ret = VbMsgBox("Can't Read Drive A:", vbRetryCancel, "Format A:", True, 101)
    
    If (Ret = vbRetry) Then
        GoTo Top:
    End If
    
End Sub

Private Sub Form_Load()
    If (InIDE) Then
        MsgBox "Please compile to see the icons."
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMain = Nothing
End Sub
