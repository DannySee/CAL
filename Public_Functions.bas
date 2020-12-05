Attribute VB_Name = "Custom_Functions"


'*******************************************************************************
'Returns a concatenated string with separator.
'*******************************************************************************
Public Function Append(val As String, sep As String, val2 As String) As String

    'If value is blank
    If val = "" Then

        'Return blank variable
        Append = val2
    Else

        'Return concatenated value and separator
        Append = val & sep & " " & val2
    End If
End Function


'*******************************************************************************
'Returns a concatenated string in SQL syntax. Parameters are a multidimensional
'array with Excel data, the row index to be parsed.
'*******************************************************************************
Public Function AppendRow(var As Variant, iRow As Integer, iStrt As Integer) As String

    'Declare function variables
    Dim strVal As String
    Dim strRow As String

    'Loop through each column of passthrough array
    For iCol = iStrt To UBound(var, 1)

        'Assemble SQL insert string
        strVal = Validate(var(iCol, iRow), iCol, iRow)

        'Assemble array of insert rows
        strRow = Append(strRow, ",", strVal)
    Next

    'Return SQL insert string
    AppendRow = strRow
End Function


'*******************************************************************************
'Return multidimensional array of Excel sheet data.
'*******************************************************************************
Public Function GetXL(strSheet As String) As Variant

    'Declare function variables
    Dim stCon as String
    Dim var As Variant

    'Save connection string (connection to CAL workbook)
    stCon = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & ThisWorkbook.FullName & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"

    'Query file (from passthrough sheet) and return results in an open recordset
    rst.Open "SELECT * FROM [" & strSheet & "$] ORDER BY PRIMARY_KEY", stCon

    'Parse recordset into an multidimensional array
    If Not rst.EOF Then var = rst.GetRows()

    'Close recordset and connection & free objects
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing

    'Return multidimensional array of Excel data (from passthrough sheet)
    GetXL = var
End Function
