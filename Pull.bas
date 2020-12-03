Attribute VB_Name = "Pull"


'*******************************************************************************
'Query programs tab. Parameter is the user's network ID. Only pulls assigned
'customers. Returns open recordset
'*******************************************************************************
Function GetPrograms() As ADODB.Recordset

    'Query program data for assigned customers
    rst.Open "SELECT * " _
        & "FROM UL_Programs " _
        & "INNER JOIN (" _
            & "SELECT MAX(END_DATE) AS ED, PROGRAM_ID AS PID " _
            & "FROM UL_Programs " _
            & "GROUP BY PROGRAM_ID) AS O " _
        & "ON PROGRAM_ID = O.PID AND END_DATE = O.ED " _
        & "WHERE CUSTOMER_ID IN (" _
            & "SELECT CUSTOMER_ID " _
            & "FROM UL_Account_Ass " _
            & "WHERE T1_ID = '" & netID & "' " _
            & "OR T2_ID = '" & netID & "')) " _
        & "ORDER BY CUSTOMER, PROGRAM_DESCRIPTION", cnn

    'Return query results
    GetPrograms = rst
End Function


'*******************************************************************************
'Query Customer Profile tab. Parameter is the user's network ID. Only pulls
'assigned customers. Returns open recordset
'*******************************************************************************
Function GetCstProfile() As ADODB.Recordset

    'Query customer profile data for assigned customers
    rst.Open "SELECT DISTINCT * " _
        & "FROM UL_Customer_Profile " _
        & "WHERE CUSTOMER_ID IN (" _
            & "SELECT CUSTOMER_ID " _
            & "FROM UL_Account_Ass " _
            & "WHERE T1_ID = '" & netID & "' " _
            & "OR T2_ID = '" & netID & "') " _
        & "ORDER BY CUSTOMER", cnn

    'Return query results
    GetCustProfile = rst
End Function


'*******************************************************************************
'Query Deviation Loads tab. Parameter is the user's network ID. Only pulls
'assigned customers. Returns open recordset
'*******************************************************************************
Function GetDevLds() As ADODB.Recordset

    'Query deviation load data for assigned customers
    rst.Open "SELECT DISTINCT * " _
        & "FROM UL_Deviation_Loads " _
        & "WHERE CUSTOMER_ID IN (" _
            & "SELECT CUSTOMER_ID " _
            & "FROM UL_Account_Ass " _
            & "WHERE T1_ID = '" & netID & "' " _
            & "OR T2_ID = '" & netID & "') " _
        & "ORDER BY CUSTOMER, PROGRAM", cnn

    'Return query results
    GetCustProfile = rst
End Function


'*******************************************************************************
'Query all drop down list data. Returns multidimensional array
'*******************************************************************************
Function GetDropDwns() As Variant

    'Declare function variables
    Dim var As Variant

    'Query drop down option list
    rst.Open "SELECT DROP_DOWN " _
        & "FROM UL_List_Options", cnn

    'Create multidimensional array from query results
    var = rst.GetRows()

    'Close recordset
    rst.Close

    'Return multidimensional array of drop down list data
    GetDropDowns = var
End Function


'*******************************************************************************
'Query all assigned customer names. Parameter is user's network ID. Returns
'array of customer names
'*******************************************************************************
Function GetMyCst() As Variant

    'Declare function variables
    Dim var As Variant

    'Query all assigned customer names
    rst.Open "SELECT CUSTOMER_NAME " _
        & "FROM UL_ACCOUNT_ASS " _
        & "WHERE T1_ID = '" & netID & "' " _
        & "OR T2_ID = '" & netID & "' " _
        & "ORDER BY CUSTOMER_NAME", cnn

    'Create array from query results
    var = rst.GetRows()

    'Close recordset
    rst.Close

    'Return Array of assigned customer names
    GetMyCst = var
End Function


'*******************************************************************************
'Query all unassigned customer names. Parameter is user's network ID. Returns
'array of customer names
'*******************************************************************************
Function GetOthCst() As Variant

    'Declare function variables
    Dim var As Variant

    'Query all unassigned customer names
    rst.Open "SELECT CUSTOMER_NAME " _
        & "FROM UL_ACCOUNT_ASS " _
        & "WHERE T1_ID <> '" & netID & "' " _
        & "AND T2_ID <> '" & netID & "' " _
        & "ORDER BY CUSTOMER_NAME", cnn

    'Create array from query results
    var = rst.GetRows()

    'Close recordset
    rst.Close

    'Return array of unassigned customer names
    GetOthCst = var
End Function
