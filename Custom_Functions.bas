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
