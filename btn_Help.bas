Attribute VB_Name = "btn_Help"

'Declare private module constants TESTING3
Private Const varShp As Variant = _Array("Help_Label", "Help_Body", _
    "Help_Send","Help_Cancel","Help_Pane")


'*******************************************************************************
'Show all utility elements.
'*******************************************************************************
Private Sub Help_Initialize()

    'Hide any visible shapes
    Utility.ClearShapes

    'Show utility elements
    Utility.Show(varShp)
End Sub


'*******************************************************************************
'Send help message to server.
'*******************************************************************************
Sub Help_Send()

    'Declare variables
    Dim strFields As String
    Dim strHelp As string

    'Get string from help box
    strHelp = ActiveSheet.TextBoxes("Help_Body").Text

    'Ensure there is data to parse
    If strHelp <> "" Then

        'Send message to database
        Push.SendHelp(strHelp)

        'Complete message
        MsgBox "Message sent!"

        'Clear all shapes from Control Panel
        Utility.ClearShapes

    'No message to Send
    Else

        'Alert user they did not input a message
        msgbox "No message sent."
    End If
End Sub


'*******************************************************************************
'Hide help window
'*******************************************************************************
Sub Help_Cancel()

    'Clear all shapes from Control Panel
    Utility.ClearShapes
End Sub
