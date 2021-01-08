Attribute VB_Name = "btn_Refresh_Data".

'Declare private module constants
Private Const varSht As Variant = _Array("Programs", "Customer Profile", _
    "Deviation Loads")


'*******************************************************************************
'Get fresh data directly from server. Delete old records, replace with new &
'format accordingly. Runs across main tabs. Test final
'*******************************************************************************
Sub Refresh_Data_Initialize()

    'Declare sub variables
    Dim tempDct As New Scripting.Dictionary
    Dim strCst As String

    'Clear data from sheets. Parameters: all sheets + header row
    Utility.ClearShts(varSht, 2)

    'Get String of my customer names
    strCst = GetStr(Pull.GetCst(True), True)

    'Exit sub routine if user is not assigned to customers
    If strCst <> "" Then

        'Refresh drop down sheet (hidden)
        Utility.AddDropDwns

        'Populate CAL sheets with refreshed data
        Utility.PopulatePages

        'Save all sheet data set to static dictionary
        Set tempDct = oPrgms.GetSaveData(True)
        Set tempDct = oCst.GetSaveData(True)
        Set tempDct = oDev.GetSaveData(True)
    End If
End Sub
