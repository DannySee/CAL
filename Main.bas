Attribute VB_Name = "Main".


'*******************************************************************************
'Declare public project variables
'*******************************************************************************
Public cnn As New ADODB.Connection
Public rst As New ADODB.Recordset
Public Programs As New Scripting.Dictionary
Public custProfile As New Scripting.Dictionary
Public devLoads As New Scripting.Dictionary
Public oPrgms As New clsPrograms
Public oCst As new clsCustProfile

Sub Refresh_Data()

    Dim netID As String

    netID = Environ("Username")

    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    Format.ShtRefresh("Programs", Query.GetPrograms(netID))

    Set Programs = Data_Maintenance.dctProgrmas(True)

    Format.ShtRefresh("Customer Profile"), Query.GetCustProfile(netID))

    Set custProfile = Data_Maintenance.dctProgrmas(True)


End Sub
