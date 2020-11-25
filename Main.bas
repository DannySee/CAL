Attribute VB_Name = "Main"

'*******************************
'Declare public project variables
'*******************************
Public cnn As New ADODB.Connection
Public rst As New ADODB.Recordset
Public oPrgms As New clsPrograms


Sub Refresh_Data()

    Dim netID As String
    Dim appID As String

    netID = Environ("Username")
    appID = Application.Username

    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    Range("A2").CopyFromRecordset Query.Programs(appID, netID)


End Sub
