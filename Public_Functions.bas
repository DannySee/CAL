Attribute VB_Name = "Public_Functions"

'Declare public project variables
Public oPrgms As New clsPrograms
Public oCst As new clsCustProfile
Public oDev As New clsDevLoads
Public iLRow As Long
Public iLCol As Integer
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
    var = Pull.GetXL(strSht)

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
Public Function GetStr(var As Variant, blStr) As String

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


'*******************************************************************************
'Return last row of active sheet assuming data starts in cell A1
'*******************************************************************************
Public Function LastRow(strSht As String) As Long

    'Return last row of active sheet
    With Sheets(strSht)
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
End Function


'*******************************************************************************
'Return last column of active sheet assuming data starts in cell A1
'*******************************************************************************
Public Function LastCol(strSht As String) As Integer

    'Return last column of active sheet
    With Sheets(strSht)
        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With
End Function


'*******************************************************************************
'Return next month and year
'*******************************************************************************
Public Function NextMonth() As String

    'Declare function variables
    Dim StrMonth As String
    Dim iYear As Integer

    'Get month and year
    strMonth = MonthName(Month(DateAdd("m", 1, Date)))
    iYear = Year(DateSerial(Year(Now), Month(Now) + 1, 1))

    'Return month and year
    NextMonth = strMonth & " " & iYear
End Function


'*******************************************************************************
'Returns user's network ID
'*******************************************************************************
Public Function GetID() As String

    'Declare function variables
    GetID = Environ("Username")
End Function


'*******************************************************************************
'Returns user's name
'*******************************************************************************
Public Function GetName() As String

    'Declare function variables
    GetName = Application.Username
End Function
