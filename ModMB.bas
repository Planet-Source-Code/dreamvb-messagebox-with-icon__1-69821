Attribute VB_Name = "ModMB"
Private Declare Function MessageBoxIndirect Lib "user32" Alias "MessageBoxIndirectA" (lpMsgBoxParams As MSGBOXPARAMS) As Long

Private Type MSGBOXPARAMS
    cbSize As Long
    hwndOwner As Long
    hInstance As Long
    lpszText As String
    lpszCaption As String
    dwStyle As Long
    lpszIcon As Long
    dwContextHelpId As Long
    lpfnMsgBoxCallback As Long
    dwLanguageId As Long
End Type

Public Function VbMsgBox(Optional Prompt = "", Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
Optional Title = "", Optional CustomIcon As Boolean, Optional IcoResID As Integer = 101) As VbMsgBoxResult
Dim mParms As MSGBOXPARAMS
    
    With mParms
        .cbSize = Len(mParms)
        .hwndOwner = 0
        .hInstance = App.hInstance
        'Set caption and message.
        .lpszCaption = Title
        .lpszText = Prompt
        'Check if using custom icon.
        If (CustomIcon) Then
            'Set the dialog style VbMsgBoxStyle and custom icon
            .dwStyle = (Buttons Or &H80&)
            'Assign ResID for the custom icon.
            .lpszIcon = IcoResID
        Else
            'Using normal Default VbMsgBoxStyle
            .dwStyle = Buttons
        End If
        
        'Show the messagebox.
        VbMsgBox = MessageBoxIndirect(mParms)
        'Clear string variables.
        .lpszCaption = vbNullString
        .lpszText = vbNullString
    End With
End Function
