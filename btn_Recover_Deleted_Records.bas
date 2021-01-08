Attribute VB_Name = "btn_Recover_Deleted_Records"

'Declare private module constants
Private Const varSht As Variant = _Array("Recover Programs", _
    "Recover Cust Profile", "Recover Deviation Loads")


'*******************************************************************************
'Show all utility elements and popluate data
'*******************************************************************************
Private Sub Recover_Deleted_Initialize()

    'Declare module variables
    Dim strCst As String

    'Get String of my customer names
    strCst = GetStr(Pull.GetCst(True), True)

    'Clear data from sheets. Parameters: utility sheets + header row
    Utility.ClearShts(varSht, 3)

    'Show utility sheets
    Utility.ShowSheets(varSht)

    'Hide all other Sheets (Sheets that do not contain key word Recover)
    Utility.HideSheets("Recover")

    'Format sheets and insert updated server data
    Utility.ShtRefresh("Recover Deleted", Pull.GetDelPrograms(strCst))
    Utility.ShtRefresh("Recover Cust Profile", Pull.GetDelCst(strCst))
    Utility.ShtRefresh("recover Deviation Loads", Pull.GetDelDev(strCst))

    '
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
