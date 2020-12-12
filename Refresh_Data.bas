Attribute VB_Name = "Refresh_Data".


'Declare public project variables
Public cnn As New ADODB.Connection
Public rst As New ADODB.Recordset
Public oPrgms As New clsPrograms
Public oCst As new clsCustProfile
Public oDev As New clsDevLoads
Public oBtnPull As New clsPullCst
Public netID As String
Public i As Integer


'*******************************************************************************
'Get fresh data directly from server. Delete old records, replace with new &
'format accordingly. Runs across main tabs.
'*******************************************************************************
Sub Refresh()

    'Declare sub variables
    Dim tempDct As New Scripting.Dictionary
    Dim strCst As String

    'Set variable to current user Network ID
    netID = Environ("Username")

    'Clear data from sheets
    Format.ClearShts

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Get String of my customer names
    strCst = GetStr(Pull.GetCst(True), True)

    'Exit sub routine if user is not assigned to customers
    If strCst = "" Then Goto NoCst

    'Refresh drop
    Format.AddDropDwns

    'Format sheets and insert updated server data
    Format.ShtRefresh(oPrgms, Pull.GetPrograms(strCst))
    Format.ShtRefresh(oCst, Pull.GetCstProfile(strCst))
    Format.ShtRefresh(oDevLds, Pull.GetDevLds(strCst))

    'Close connection & free connections
    cnn.Close
    Set cnn = Nothing
    Set rst = Nothing

    'Save all sheet data set to static dictionary
    Set tempDct = oPrgms.GetSaveData(True)
    Set tempDct = oCst.GetSaveData(True)
    Set tempDct = oDev.GetSaveData(True)

'Jump to label to skip sub routine
NoCst:
End Sub
