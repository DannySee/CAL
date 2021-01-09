Attribute VB_Name = "Pull"


'*******************************************************************************
'Query programs tab. Parameter is the user's network ID. Only pulls assigned
'customers. Returns open recordset
'*******************************************************************************
Function GetPrograms(strCst As String, strFlds As String) As ADODB.Recordset

    'Declare function variables
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query program data for assigned customers
    rst.Open "SELECT " & strFlds & " " _
        & "FROM UL_Programs " _
        & "INNER JOIN (" _
            & "SELECT MAX(END_DATE) AS ED, PROGRAM_ID AS PID " _
            & "FROM UL_Programs " _
            & "GROUP BY PROGRAM_ID) AS O " _
        & "ON PROGRAM_ID = O.PID AND END_DATE = O.ED " _
        & "WHERE CUSTOMER IN (" & strCst & ") " _
        & "ORDER BY CUSTOMER, PROGRAM_DESCRIPTION", cnn

    'Return query results
    GetPrograms = rst
End Function


'*******************************************************************************
'Query deleted records. Only pulls assigned customers or records deleted
'by user.
'*******************************************************************************
Function GetDelRecords(strCst As String, strDb As String) As ADODB.Recordset

    'Declare function variables
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query deleted records for assigned customers
    rst.Open "SELECT * " _
        & "FROM " & strDb & " " _
        & "WHERE CUSTOMER IN (" & strCst & ")" _
        & "OR DEL_USER = '" & GetName & "' " _
        & "ORDER BY CUSTOMER, PROGRAM_DESCRIPTION", cnn

    'Return query results
    GetDeletedRecords = rst
End Function


'*******************************************************************************
'Query programs tab. Parameter is the user's network ID. Only pulls assigned
'customers. Returns open recordset.
'*******************************************************************************
Function GetExpPrograms(strCst As String) As ADODB.Recordset

    'Declare function variables
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Pull expiring contracts
    rst.Open "SELECT PROGRAM_DESCRIPTION " _
        & "FROM UL_Programs " _
        & "INNER JOIN (" _
            & "SELECT MAX(END_DATE) AS ED, PROGRAM_ID AS PID " _
            & "FROM UL_Programs " _
            & "WHERE CUSTOMER = '" & strCst & "' " _
            & "GROUP BY PROGRAM_ID) AS O " _
        & "ON PROGRAM_ID = O.PID AND END_DATE = O.ED " _
        & "WHERE VENDOR_NUM <> 1 " _
        & "AND END_DATE < " & Now(), cnn

    'Return query results
    GetExpPrograms = rst
End Function


'*******************************************************************************
'Query Customer Profile tab. Parameter is the user's network ID. Only pulls
'assigned customers. Returns open recordset
'*******************************************************************************
Function GetCstProfile(strCst As String) As ADODB.Recordset

    'Declare function variables
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query customer profile data for assigned customers
    rst.Open "SELECT DISTINCT * " _
        & "FROM UL_Customer_Profile " _
        & "WHERE CUSTOMER_NAME IN (" & strCst & ") " _
        & "ORDER BY CUSTOMER", cnn

    'Return query results
    GetCstProfile = rst
End Function


'*******************************************************************************
'Query Deviation Loads tab. Parameter is the user's network ID. Only pulls
'assigned customers. Returns open recordset
'*******************************************************************************
Function GetDevLds(strCst As String) As ADODB.Recordset

    'Declare function variables
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query deviation load data for assigned customers
    rst.Open "SELECT DISTINCT * " _
        & "FROM UL_Deviation_Loads " _
        & "WHERE CUSTOMER_NAME IN (" & strCst & ") " _
        & "ORDER BY CUSTOMER, PROGRAM", cnn

    'Return query results
    GetDevLds = rst
End Function


'*******************************************************************************
'Query all drop down list data. Returns multidimensional array
'*******************************************************************************
Function GetDropDwns() As Variant

    'Declare function variables
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query drop down option list
    rst.Open "SELECT DROP_DOWN " _
        & "FROM UL_List_Options", cnn

    'Return multidimensional array of drop down list data
    GetDropDowns = rst.GetRows()
End Function


'*******************************************************************************
'Query all  customer names. Parameter is user's network ID. Returns
'array of customer names
'*******************************************************************************
Function GetCst(blMyCst As Boolean) As Variant

    'Declare function variables
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim strVal As String
    Dim strOp As String

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Set equal character
    If blMyCst = False Then
        strOp = "<>"
    Else
        strOp = "="
    End If

    'Query all assigned customer names
    rst.Open "SELECT CUSTOMER_NAME AS CST " _
        & "FROM UL_ACCOUNT_ASS " _
        & "WHERE T1_ID " & strOp & " '" & GetID & "' " _
        & "ORDER BY CUSTOMER_NAME", cnn

    'Assemble string from query results
    Do While rst.EOF = False
        strVal = Append(strVal & "," & rst.Fields("CST").Value)
        rst.MoveNext
    Loop

    'Return Array of assigned customer names
    GetCst = Split(strVal, ",")
End Function


'*******************************************************************************
'Query customer packet. Parameter is SQL string of customer names. Returns
'array of customer packets.
'*******************************************************************************
Function GetCstPacket(strCst As String) As Variant

    'Declare function variables
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query all assigned customer names
    rst.Open "SELECT DISTINCT PACKET AS PKT " _
        & "FROM UL_Customer_Profile " _
        & "WHERE CUSTOMER IN ()" & strCst & ")", cnn

    'Assemble string from query results
    Do While rst.EOF = False
        strVal = Append(strVal & "," & rst.Fields("PKT").Value)
        rst.MoveNext
    Loop

    'Return Array of assigned customer names
    GetCstPacket = Split(strVal, ",")
End Function


'*******************************************************************************
'Query all acount assignments given network ID(s).
'*******************************************************************************
Function GetAssignments(strID As String) As Variant

    'Declare function variables
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim strVal As String

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query all assigned customer names
    rst.Open "SELECT CUSTOMER_NAME AS CST " _
        & "FROM UL_ACCOUNT_ASS " _
        & "WHERE T1_ID IN (" & strID & ")", cnn

    'Assemble string from query results
    Do While rst.EOF = False
        strVal = Append(strVal & "," & rst.Fields("CST").Value)
        rst.MoveNext
    Loop

    'Return Array of assigned customer names
    GetAssignments = strVal
End Function


'*******************************************************************************
'Pull customer ID given customer name
'*******************************************************************************
Function GetCstID(strCst As String) As Long

    'Declare function variables
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query customer ID from customer name
    rst.Open "SELECT TOP 1 CUSTOMER_ID AS CID " _
        & "FROM UL_Account_Ass " _
        & "WHERE CUSTOMER_NAME = '" & strCst & "'", cnn

    'Return query results
    GetCstID = rst.Fields("CID").Value
End Function


'*******************************************************************************
'Pull customer ID & Program ID given customer name
'*******************************************************************************
Function GetCstPgmID(strCst As String) As String

    'Declare function variables
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query customer ID and program ID from customer name
    rst.Open "SELECT TOP 1 CUSTOMER_ID AS CID, MAX(CAST(RIGHT(PROGRAM_ID, " _
        & "CHARINDEX('-', REVERSE(PROGRAM_ID)) - 1) AS INT)) + 1 AS PID " _
        & "FROM UL_Programs WHERE CUSTOMER = '" & strCst & "' " _
        & "GROUP BY CUSTOMER_ID", cnn

    'Return string of customer ID and program ID
    GetCstID = rst.Fields("CID").value & ",'" & rst.Fields("CID").value _
        & "-" & rst.Fields("PID").value & "'"
End Function


'*******************************************************************************
'Return multidimensional array of Excel sheet data.
'*******************************************************************************
Public Function GetXL(strSht As String) As Variant

    'Declare function variables
    Dim stCon as String

    'Save connection string (connection to CAL workbook)
    stCon = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & ThisWorkbook.FullName & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"

    'Query file (from passthrough sheet) and return results in an open recordset
    rst.Open "SELECT * FROM [" & strSht & "$] ORDER BY PRIMARY_KEY", stCon

    'Return multidimensional array of Excel data (from passthrough sheet)
    If Not rst.EOF Then GetXL = rst.GetRows
End Function


'*******************************************************************************
'Pull all accoiate names (excluding user's name).
'*******************************************************************************
Function GetAssName() As Variant

    'Declare function variables
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim strVal As String

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query customer ID from customer name
    rst.Open "SELECT DISTINCT TIER_1 AS ASS " _
        & "FROM UL_Account_Ass " _
        & "WHERE TIER_1 <> '" & GetID & "'", cnn

    'Assemble string from query results
    Do While rst.EOF = False
        strVal = Append(strVal & "," & rst.Fields("ASS").Value)
        rst.MoveNext
    Loop

    'Return Array of associate names
    GetAssName = Split(strVal, ",")
End Function


'*******************************************************************************
'Return all OpCo sites given list of Packets.
'*******************************************************************************
Function GetOpcos(strPacket As String) As Variant

    'Declare function variables
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim iErr As Integer
    Dim strUid As String
    Dim strPwd As String

    'Get username and password
    strUid = get_uid
    strPwd = get_pwd

    'Connect to OpCo
    On Error GoTo OpErr
    cnn.Open "DSN=AS240A;UID=" & strUid & ";PASSWORD=" & strPwd & ";"
    On Error GoTo 0

    'Query OpCo list for each packet
    rst.Open "SELECT DISTINCT TRIM(DVPKGS) AS PCKT, TRIM(DVT500) AS OP " _
        & "FROM SCDBFP10.PMDPDVRF " _
        & "INNER JOIN (" _
            & "SELECT DVPKGS AS PACKET, MAX(LENGTH(TRIM(DVT500))) AS LEN " _
            & "FROM SCDBFP10.PMDPDVRF " _
            & "WHERE TRIM(DVPKGS) IN (" & strPacket & ") " _
            & "GROUP BY DVPKGS) " _
        & "ON DVPKGS = PACKET AND LENGTH(TRIM(DVT500)) = LEN ", cnn

    'Return multidimensional of packet data
    OpCos = rst.GetRows

    'Exit Function before error handler
    Exit Function

'Could not connect to OpCo
OpErr:

    'If catastrophic error
    If InStr(Err.Description, "Catastrophic") > 0 Then
        MsgBox "OBDC overload. " & vblf & vblf _
            & "Please close all open instances of Excel and try again."

    'If invalid password
    ElseIf iErr < 2 Then

        'Prompt user to enter SUS credentials
        UserLog.Show

        'Get updated credentials from server
        strUid = get_uid
        strPwd = get_pwd

        'Iterate login attempt counter and resume login
        iErr = iErr + 1
        Resume

    'All login attempts unsuccessful
    Else
        MsgBox "Could not reach OpCo. Please validate you have access to as240a"
    End If
End Function
s

'*******************************************************************************
'Get associate SUS password
'*******************************************************************************
Function get_pwd() As String

    'Declare function variables
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim i       As Integer
    Dim StrDec  As String
    Dim strPwd  As String
    Dim StrKey  As String

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query uid & password
    rst.Open "SELECT SUS_PWD, CRED_ID " _
        & "FROM Login_Cred " _
        & "WHERE NET_ID = '" & GetID & "'", cnn

    'If records were successfully queried
    If Not rst.EOF Then

        'save value to string
        strPwd = rst.Fields("SUS_PWD").value
        StrKey = rst.Fields("CRED_ID").value

        'Loop through each character in password
        For i = 1 To Len(strPwd)

            'Decrypt character of password
            StrDec = StrDec & Chr(Asc(Mid(strPwd, i, 1)) - Mid(StrKey, i, 1))
        Next

        'Return decryted SUS password
        get_pwd = strDec
    End If
End Function


'*******************************************************************************
'Get associate SUS username
'*******************************************************************************
Function get_uid() As String

    'Declare function variables
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query uid & password
    rst.Open "SELECT SUS_ID AS ID " _
        & "FROM Login_Cred " _
        & "WHERE NET_ID = '" & GetID & "'", cnn

    'If records were successfully queried return SUS ID
    If Not rst.EOF Then get_uid = rst.Fields("ID").value
End Function
