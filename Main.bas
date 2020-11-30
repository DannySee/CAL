Attribute VB_Name = "Main".


'Declare public project variables
Public cnn As New ADODB.Connection
Public rst As New ADODB.Recordset
Public Programs As New Scripting.Dictionary
Public cstProfile As New Scripting.Dictionary
Public devLds As New Scripting.Dictionary
Public oPrgms As New clsPrograms
Public oCst As new clsCustProfile
Public oDev As New clsDevLoads
Public netID As String


'*******************************************************************************
'Get fresh data directly from server. Delete old records, replace with new &
'format accordingly. Runs across main tabs.
'*******************************************************************************
Sub Refresh_Data()

    'Set variable to current user Network ID
    netID = Environ("Username")

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Format sheets and insert updated server data
    Format.ShtRefresh("Programs", Query.GetPrograms)
    Format.ShtRefresh("Customer Profile"), Query.GetCstProfile)
    Format.ShtRefresh("Deviation Loads"), Query.GetDevLds)

    'Add drop down lists file and include data validation on Programs tab
    Format.AddDropDwns(Query.GetDropDwns, Query.GetMyCst, Query.GetOthCst)

    'Save all sheet data set to static dictionary
    Set Programs = Data_Maintenance.dctPrograms(True)
    Set cstProfile = Data_Maintenance.dctCstProfile(True)
    Set devLds = Data_Maintenance.dctDevLds(True)

    'Close connection
    cnn.Close

    'Free Objects
    Set cnn = Nothing
    Set rst = Nothing
End Sub
