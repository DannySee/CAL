Attribute VB_Name = "btn_Refresh_Data".


'*******************************************************************************
'Get fresh data directly from server. Delete old records, replace with new &
'format accordingly. Runs across main tabs. Test final
'*******************************************************************************
Sub Initialize()

    'Declare sub variables
    Dim tempDct As New Scripting.Dictionary
    Dim strCst As String

    'Clear data from sheets
    Utility.ClearShts

    'Get String of my customer names
    strCst = GetStr(Pull.GetCst(True), True)

    'Exit sub routine if user is not assigned to customers
    If strCst <> "" Then

        'Refresh drop
        Utility.AddDropDwns

        'Format sheets and insert updated server data
        Utility.ShtRefresh(oPrgms, Pull.GetPrograms(strCst, "*"))
        Utility.ShtRefresh(oCst, Pull.GetCstProfile(strCst, "*"))
        Utility.ShtRefresh(oDevLds, Pull.GetDevLds(strCst, "*"))

        'Save all sheet data set to static dictionary
        Set tempDct = oPrgms.GetSaveData(True)
        Set tempDct = oCst.GetSaveData(True)
        Set tempDct = oDev.GetSaveData(True)
    End If
End Sub
