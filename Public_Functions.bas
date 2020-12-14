Attribute VB_Name = "Custom_Functions"

'Declare public project variables
Public cnn As New ADODB.Connection
Public rst As New ADODB.Recordset
Public oPrgms As New clsPrograms
Public oCst As new clsCustProfile
Public oDev As New clsDevLoads
Public oBtnPull As New clsPullCst
Public netID As String
Public i As Integer


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
Public Function AppendRow(var As Variant, iRow As Integer, iStrt As Integer) _
    As String

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
Public Function GetXL(strSht As String) As Variant

    'Declare function variables
    Dim stCon as String
    Dim var As Variant

    'Save connection string (connection to CAL workbook)
    stCon = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & ThisWorkbook.FullName & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"

    'Query file (from passthrough sheet) and return results in an open recordset
    rst.Open "SELECT * FROM [" & strSht & "$] ORDER BY PRIMARY_KEY", stCon

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


'*******************************************************************************
'Returns updated dictionary of Excel data (Key = Primary_Key, Value = Array
'of fields). Meant to update the passthrough dictionary with static dictionary.
'*******************************************************************************
Public Function RefreshDct(strSht As String) As Scripting.Dictionary

    'Declare function variables
    Dim dct As New Scripting.Dictionary
    Dim arr As Variant
    Dim var As Variant
    Dim iRow As Integer
    Dim iCol As Integer

    'Save multidimensional array of program data
    var = GetXL(strSht)

    'Create an empty array with an index for each program field'
    ReDim arr(UBound(var, 1))

    'Loop through each row of program data
    For iRow = 0 To UBound(var, 2)

        'Loop through each column of program data & add element to array
        For iCol = 0 To UBound(var, 1)
            arr(iCol) = var(iCol, iRow)
        Next

        'Add line to dictionary with program ID as key & row fields as value
        dct(var(0, iRow)) = arr
    Next

    'Return dictionary of program data
    Set RefreshDct = dct
End Function


'*******************************************************************************
'Get comma delimited string from array. blStr indicates if string should also be
'quote delimited.
'*******************************************************************************
Public Function GetStr(upd As Variant, blStr) As String

    'Declare function variables
    Dim i As Integer
    Dim strVal As String
    Dim sep As String

    'Get comma delimiter if string
    If blStr = True Then sep = "'"

    'Setup looping variable
    i = 0

    'Loop through recordset
    For i = 0 To Ubound(var)

        'Assemble string of customer names
        strVal = Append(strVal, ",", sep & var(i) & sep)
    Next

    'Return string of customer names
    GetStr = strVal
End Function
