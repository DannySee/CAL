Attribute VB_Name = "Query"


Function GetPrograms(netID As String) As ADODB.Recordset

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

    GetPrograms = rst

End Function


Function GetCustProfile(netID As String) As ADODB.Recordset

    rst.Open "SELECT DISTINCT * " _
        & "FROM UL_Customer_Profile " _
        & "WHERE CUSTOMER_ID IN (" _
            & "SELECT CUSTOMER_ID " _
            & "FROM UL_Account_Ass " _
            & "OR T1_ID = '" & netID & "' " _
            & "OR T2_ID = '" & netID & "') " _
        & "ORDER BY CUSTOMER", cnn

    GetCustProfile = rst

End Function
