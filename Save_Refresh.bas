Attribute VB_Name = "IN_PROGRESS"
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim oPrgms As New clsPrograms

'***************************************************************
'Returns a static dicitonary of Program data (Key = Program ID,
'Value = array of each field). Boolean operator indicates
'if the dictionary needs to be updated before it returns
'***************************************************************
Function GetDict(blUpdate As Boolean) As Scripting.Dictionary

    'declare static dictionary to hold program data
    Static dictPrograms As New Scripting.Dictionary

    'Update dictionary before it is returned (if indicated by passthrough boolean)
    If blUpdate = True Then Set dictPrograms = UpdateDict(dictPrograms)

    'Return static dictionary
    Set GetDict = dictPrograms

End Function

Function UpdateDict(dictPrograms As Scripting.Dictionary) As Scripting.Dictionary

    Dim iRow, iCol As Integer
    Dim arr As Variant
    Dim var As Variant
    
    dictPrograms.RemoveAll
    
    var = GetXL("Programs$")
    
    ReDim arr(UBound(var, 1))
    
    For iRow = 0 To UBound(var, 2)
        For iCol = 0 To UBound(var, 1)
            arr(iCol) = var(iCol, iRow)
        Next
        
        dictPrograms(var(2, iRow)) = arr
    Next
    
    Set UpdateDict = dictPrograms

End Function

Function GetXL(strSheet As String) As Variant

    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
        "Data Source=" & ThisWorkbook.FullName & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
        
    rst.Open "SELECT * FROM [" & strSheet & "]", cnn
    
    var = rst.GetRows()
    
    rst.Close
    cnn.Close
    Set rst = Nothing
    Set cnn = Nothing
    
    GetXL = var
    
End Function

Function Compare(old As Scripting.Dictionary, upd As Variant) As Variant

    Dim iRow As Integer
    Dim iCol As Integer
    Dim pID As String
    Dim strInsert As String
    Dim strAssemble As String
    Dim strUpdate As String
    Dim strVal As String
    Dim strInsRows As String
    
    
    For iRow = 0 To UBound(upd, 2)
    
        If IsNull(upd(oPrgms.ColIndex("CUSTOMER_ID"), iRow)) Then
            If upd(oPrgms.ColIndex("CUSTOMER"), iRow) <> "" Then
            
                strInsert = Append(strInsert, "|") & GetInsertString(upd, iRow)
                strInsRows = Append(iRow + 2, "|")
                
            End If
        Else
        
            pID = upd(oPrgms.ColIndex("PROGRAM_ID"), iRow)
            
            strAssemble = ""
        
            For iCol = 0 To UBound(upd, 1)
            
                If (old(pID)(iCol) <> upd(iCol, iRow)) Or _
                    (IsNull(old(pID)(iCol)) <> IsNull(upd(iCol, iRow))) Then
                    
                    strVal = Validate(upd(iCol, iRow), iCol, iRow)
                
                    If strVal <> "'DateErr'" Then _
                        strAssemble = Append(strAssemble, ",") & oPrgms.Cols(iCol) & " = " & strVal
 
                End If
 
            Next
            
            If InStr(strAssemble, "START_DATE") <> 0 Then
            
                    strUpdate = Append(strUpdate, "|") & "END_DATE = '" _
                        & upd(oPrgms.ColIndex("START_DATE"), iRow) - 1 & "' " _
                        & "WHERE PRIMARY_KEY = " & upd(oPrgms.ColIndex("PRIMARY_KEY"), iRow)
                        
                    strInsert = Append(strInsert, "|") & GetInsertString(upd, iRow)
                    strInsRows = Append(iRow + 2, "|")
                
            ElseIf strAssemble <> "" Then
            
                strUpdate = Append(strUpdate, "|") & strAssemble & " WHERE PRIMARY_KEY = " & upd(0, iRow)
                
            End If
 
        End If
    Next
    
    If strUpdate = "" Then strUpdate = "0"
    If strInsert = "" Then strInsert = "0"
    
    Compare = Array(Split(strUpdate, "|"), Split(strInsert, "|"), Split(strInsRows, "|"))
    
End Function

Function Validate(val As Variant, iCol As Integer, iRow As Integer)

    Dim sep As String
    
    sep = oPrgms.ColType(iCol)
    
    If (iCol = oPrgms.ColIndex("START_DATE") Or iCol = oPrgms.ColIndex("END_DATE")) And _
        Not IsDate(val) Then
        
        MsgBox "INVALID DATE: Please correct the " & oPrgms.Cols(iCol) & " field - " _
            & val & " is not a valid date entry"
        
        Rows(iRow + 2).EntireRow.Select
        
        val = "DateErr"
        
    ElseIf iCol = oPrgms.ColIndex("VENDOR_NUM") And IsNull(val) Then
    
        val = "0"
    
    ElseIf InStr(val, "'") <> 0 Then
        
        val = Replace(val, "'", "")
    
    End If
    
    Validate = sep & val & sep
        
End Function

Function Append(str As String, sep As String) As String

    If str = "" Then
        Append = ""
    Else
        Append = str & sep & " "
    End If
    
End Function

Function GetInsertString(var As Variant, iRow As Integer) As String

    Dim strRow As String
    Dim iCol As Integer
    Dim strVal As String
    
    If IsNull(var(oPrgms.ColIndex("CUSTOMER_ID"), iRow)) Then
    
        cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"
        rst.Open "SELECT TOP 1 CUSTOMER_ID AS CID, " _
            & "MAX(CAST(RIGHT(PROGRAM_ID, CHARINDEX('-', REVERSE(PROGRAM_ID)) - 1) AS INT)) + 1 AS PID " _
            & "FROM UL_Programs WHERE CUSTOMER = '" & var(oPrgms.ColIndex("CUSTOMER"), iRow) & "' " _
            & "GROUP BY CUSTOMER_ID", cnn
    
        strRow = rst.Fields("CID").value & ",'" & rst.Fields("CID").value & "-" & rst.Fields("PID").value & "'"
        
        rst.Close
        cnn.Close
        Set rst = Nothing
        Set cnn = Nothing
    Else
        
        strRow = var(oPrgms.ColIndex("CUSTOMER_ID"), iRow) & ",'" & var(oPrgms.ColIndex("PROGRAM_ID"), iRow) & "'"
    End If
    
    For iCol = oPrgms.ColIndex("DAB") To UBound(var, 1)
        
        strVal = Validate(var(iCol, iRow), iCol, iRow)
    
        strRow = strRow & "," & strVal
    Next
    
    GetInsertString = strRow

End Function

Function UploadChanges(upd As Variant)

    Dim i As Integer
    
    For i = 0 To UBound(upd)
    
        cnn.Execute "UPDATE UL_Programs SET " & upd(i)
        
    Next
    
End Function

Function InsertNew(ins As Variant, insRow As Variant)

    Dim i As Integer
    
    For i = 0 To UBound(ins)
    
        rst.Open "INSERT INTO UL_Programs " _
            & "OUTPUT inserted.PRIMARY_KEY AS PKEY, " _
            & "inserted.CUSTOMER_ID AS CID, " _
            & "inserted.PROGRAM_ID AS PID " _
            & "VALUES(" & ins(i) & ")", cnn
            
        Cells(insRow(i), oPrgms.ColIndex("PRIMARY_KEY") + 1).value = rst.Fields("PKEY").value
        Cells(insRow(i), oPrgms.ColIndex("CUSTOMER_ID") + 1).value = rst.Fields("CID").value
        Cells(insRow(i), oPrgms.ColIndex("PROGRAM_ID") + 1).value = rst.Fields("PID").value
        
        rst.Close

    Next

End Function

Sub Fake_Refresh()

    Dim Programs As New Scripting.Dictionary
    
    Set Programs = GetDict(True)
    
End Sub

Sub Fake_Save()

    Dim Programs As New Scripting.Dictionary
    Dim var     As Variant
    Dim varChange As Variant
    
    Set Programs = GetDict(False)
    var = GetXL("Programs$")
    
    varChange = Compare(Programs, var)
    
    If varChange(0)(0) <> 0 Or varChange(1)(0) <> 0 Then
    
        cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"
    
        If varChange(0)(0) <> 0 Then UploadChanges (varChange(0))
        If varChange(1)(0) <> 0 Then RetVal = InsertNew(varChange(1), varChange(2))
         
        cnn.Close
        Set cnn = Nothing
    End If
    
    
    
End Sub
