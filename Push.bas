Attribute VB_Name = "Pull"


'*******************************************************************************
'Executes SQL update statement to CAL database. Paramaters are an array of
'update statements (one element per statement).
'*******************************************************************************
Sub Update(upd As Variant, strDb As String)

    'Declare sub variables
    Dim cnn As New ADODB.Connecton
    Dim i As Integer

    'Establish connection to SQL server
    cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;" _
        & "DATABASE=Pricing_Agreements;"

    'Loop through each update statement in passthrough array
    For i = 0 To UBound(upd)

        'Execute update statement
        cnn.Execute "UPDATE " & strDb & " SET " & upd(i)
    Next
End Sub


'*******************************************************************************
'Executes SQL insert statement to CAL database. Insert deleted customer elements
'into archive table.
'*******************************************************************************
Sub InsertDeleted(strIns As String, strDb As String)

    'Declare sub variables
    Dim cnn As New ADODB.Connecton

    'Establish connection to SQL server
    cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;" _
        & "DATABASE=Pricing_Agreements;"

    'Insert records into archive table
    cnn.Execute "INSERT INTO " & strDb & " VALUES(" & strIns & ")"
End Sub


'*******************************************************************************
'Send help message to server.
'*******************************************************************************
Sub SendHelp(str As String)

    'Declare sub variables
    Dim cnn As New ADODB.Connecton

    'Establish connection to SQL server
    cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;" _
        & "DATABASE=Pricing_Agreements;"

    'Insert help message into SQL Server table
    cnn.Execute("INSERT INTO UL_Help " _
        & "VALUES('" & GetID & "', '" & GetName & "', '" & str & "', '" _
        & Now() & "')")
End Sub


'*******************************************************************************
'Executes SQL Insert statement to CAL database. Select records from archive
'table and return it to main data table
'*******************************************************************************
Sub RecoverDeleted(obj As Object, iPKey As Long)

    'Declare sub variables
    Dim cnn As New ADODB.Connecton
    Dim rst As New ADODB.Recordset
    Dim i As Integer

    'Establish connection to SQL server
    cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;" _
        & "DATABASE=Pricing_Agreements;"

    'Insert records from recovery table to main table
    rst.Open "INSERT INTO " & obj.Db & " " _
        & "SELECT " obj.AllFlds " FROM " & obj.Dbx & " " _
        & "WHERE PRIMARY_KEY = " & iPKey, cnn

    'Remove recovered records from archive
    cnn.Execute("DELETE FROM " & obj.Dbx & " WHERE PRIMARY_KEY = " & iPKey)
End Sub


'*******************************************************************************
'Executes SQL Insert statement to CAL database. Returns recordset of gutter
'fields if inserted lines i.e. Primary Key, Customer ID etc
'*******************************************************************************
Function Insert(strSQL) As ADODB.Recordset

    'Declare sub variables
    Dim cnn As New ADODB.Connecton
    Dim rst As New ADODB.Recordset

    'Establish connection to SQL server
    cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;" _
        & "DATABASE=Pricing_Agreements;"

    'Insert new lines and return specified gutter fields
    rst.Open strSQL, cnn

    'Return recordset of gutter fields
    Set Insert = rst
End Function


'*******************************************************************************
'Executes SQL Delete statement to CAL database. Returns multidimensional array
'of deleted cusotmer elements.
'*******************************************************************************
Function GetDeleted(strDel As String, strDb As String) As Variant

    'Declare sub variables
    Dim cnn As New ADODB.Connecton
    Dim rst As New ADODB.Recordset
    Dim i As Integer

    'Establish connection to SQL server
    cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;" _
        & "DATABASE=Pricing_Agreements;"

    'Loop through each update statement in passthrough array
    rst.Open "DELETE FROM " & strDb & " " _
        & "OUTPUT DELETED.* " _
        & "WHERE PROGRAM_ID IN (" & strDel & ")", cnn

    'Return multidimensional array of deleted elements
    GetDeleted = rst.GetRows
End Function
