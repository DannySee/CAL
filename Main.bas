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


Sub Refresh_Data()

    Dim netID As String

    netID = Environ("Username")

    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    Format.ShtRefresh("Programs", Query.GetPrograms(netID))
    Format.ShtRefresh("Customer Profile"), Query.GetCstProfile(netID))
    Format.ShtRefresh("Deviation Loads"), Query.GetDevLds(netID))
    Format.AddDropDwns(Query.GetDropDwns, GetMyCst(netID), _
        GetOthCst(netID))

    Set Programs = Data_Maintenance.dctPrograms(True)
    Set cstProfile = Data_Maintenance.dctCstProfile(True)
    Set devLds = Data_Maintenance.dctDevLds(True)

End Sub
