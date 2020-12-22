Attribute VB_Name = "Pull"


'*******************************************************************************
'Query programs tab. Parameter is the user's network ID. Only pulls assigned
'customers. Returns open recordset
'*******************************************************************************
Function GetPrograms(strCst As String, strFlds As String) As ADODB.Recordset

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
        & "WHERE CUSTOMER IN (" & strCst & ")"
        & "ORDER BY CUSTOMER, PROGRAM_DESCRIPTION", cnn

    'Return query results
    GetPrograms = rst

    'Close connection/recordset and free free objects
    FreeObjects
End Function


'*******************************************************************************
'Query programs tab. Parameter is the user's network ID. Only pulls assigned
'customers. Returns open recordset
'*******************************************************************************
Function GetExpPrograms(strCst As String) As ADODB.Recordset

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

    'Close connection/recordset and free free objects
    FreeObjects
End Function


'*******************************************************************************
'Query Customer Profile tab. Parameter is the user's network ID. Only pulls
'assigned customers. Returns open recordset
'*******************************************************************************
Function GetCstProfile(strCst As String) As ADODB.Recordset

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query customer profile data for assigned customers
    rst.Open "SELECT DISTINCT * " _
        & "FROM UL_Customer_Profile " _
        & "WHERE CUSTOMER_NAME IN (" & strCst & ") " _
        & "ORDER BY CUSTOMER", cnn

    'Return query results
    GetCustProfile = rst

    'Close connection/recordset and free free objects
    FreeObjects
End Function


'*******************************************************************************
'Query Deviation Loads tab. Parameter is the user's network ID. Only pulls
'assigned customers. Returns open recordset
'*******************************************************************************
Function GetDevLds(strCst As String) As ADODB.Recordset

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query deviation load data for assigned customers
    rst.Open "SELECT DISTINCT * " _
        & "FROM UL_Deviation_Loads " _
        & "WHERE CUSTOMER_NAME IN (" & strCst & ") " _
        & "ORDER BY CUSTOMER, PROGRAM", cnn

    'Return query results
    GetCustProfile = rst

    'Close connection/recordset and free free objects
    FreeObjects
End Function


'*******************************************************************************
'Query all drop down list data. Returns multidimensional array
'*******************************************************************************
Function GetDropDwns() As Variant

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Query drop down option list
    rst.Open "SELECT DROP_DOWN " _
        & "FROM UL_List_Options", cnn

    'Return multidimensional array of drop down list data
    GetDropDowns = rst.GetRows()

    'Close connection/recordset and free free objects
    FreeObjects
End Function


'*******************************************************************************
'Query all  customer names. Parameter is user's network ID. Returns
'array of customer names
'*******************************************************************************
Function GetCst(blMyCst As Boolean) As Variant

    'Declare function variables
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

    'Set variable to current user Network ID
    netID = Environ("Username")

    'Query all assigned customer names
    rst.Open "SELECT CUSTOMER_NAME AS CST " _
        & "FROM UL_ACCOUNT_ASS " _
        & "WHERE T1_ID " & strOp & " '" & netID & "' " _
        & "ORDER BY CUSTOMER_NAME", cnn

    'Assemble string from query results
    Do While rst.EOF = False
        strVal = Append(strVal & "," & rst.Fields("CST").Value)
        rst.MoveNext
    Loop

    'Return Array of assigned customer names
    GetMyCst = Split(strVal, ",")

    'Close connection/recordset and free free objects
    FreeObjects
End Function


'*******************************************************************************
'Query all acount assignments given network ID(s).
'*******************************************************************************
Function GetAssignments(strID As String) As Variant

    'Declare function variables
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
    GetMyCst = strVal

    'Close connection/recordset and free free objects
    FreeObjects
End Function


'*******************************************************************************
'close and free ADODB objects
'*******************************************************************************
Sub FreeObjects()

    'Close connection/recordset and free objects
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
End Sub
