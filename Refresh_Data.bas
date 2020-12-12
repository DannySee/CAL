Attribute VB_Name = "Refresh_Data".


'Declare public project variables
Public cnn As New ADODB.Connection
Public rst As New ADODB.Recordset
Public oPrgms As New clsPrograms
Public oCst As new clsCustProfile
Public oDev As New clsDevLoads
Public oBtnPull As New clsPullCst
Public netID As String


'*******************************************************************************
'Get string of assigned customer names.
'*******************************************************************************
Sub RefreshData()

    'Declare sub variables
    Dim strCst As String
    Dim i As Integer

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Set variable to current user Network ID
    netID = Environ("Username")

    'Query all customer assigned customer names
    rst.Open "SELECT CUSTOMER_NAME " _
        & "FROM UL_Account_Ass " _
        & "WHERE T1_ID = '" & netID & "'", cnn

    'Setup looping Integer
    i=0

    'Assemble customer string
    Do While rst.EOF = False

    Loop


End Sub


'*******************************************************************************
'Get fresh data directly from server. Delete old records, replace with new &
'format accordingly. Runs across main tabs.
'*******************************************************************************
Sub ShtRefresh()

    'Declare sub variables
    Dim tempDct As New Scripting.Dictionary

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Format sheets and insert updated server data
    Format.ShtRefresh("Programs", Pull.GetPrograms)
    Format.ShtRefresh("Customer Profile"), Pull.GetCstProfile)
    Format.ShtRefresh("Deviation Loads"), Pull.GetDevLds)

    'Add drop down lists file and include data validation on Programs tab
    Format.AddDropDwns

    'Save all sheet data set to static dictionary
    Set tempDct = oPrgms.GetSaveData(True)
    Set tempDct = oCst.GetSaveData(True)
    Set tempDct = oDev.GetSaveData(True)

    'Close connection
    cnn.Close

    'Free Objects
    Set cnn = Nothing
    Set rst = Nothing
End Sub
