Attribute VB_Name = "Format"


'*******************************************************************************
'Declare private module constants
'*******************************************************************************
Private const shtPWD = "Dac123am"
Private const shtProperties = _
    "UserInterFaceOnly:=True, " & _
    "AllowFormattingCells:=True, " & _
    "AllowDeletingRows:=True, " & _
    "AllowFormattingRows:=True, " & _
    "AllowInsertingRows:=True, " & _
    "AllowSorting:=False, " & _
    "AllowFiltering:=True"


'*******************************************************************************
'Declare private module variables
'*******************************************************************************
Private iLRow As Long


Sub ShtUnlock(strSheet As String)

    Sheets(strSheet).Unprotect shtPWD
End Sub


Sub ShtLock(strSheet As String)

    Sheets(strSheet).Protect Password:=strPWD, shtProperties
End Sub


Sub ShtRefresh(strSheet As String, upd As ADODB.Recordset)

    ShtLock(strSheet)

    With Sheets(strSheet)
        .Rows(1).AutoFilter
        iLRow = .Range("A" & .Rows.Count).End(xlUp)
        .Range("A2:A" & iLRow + 1).EntireRow.Delete
        .Range("A2").CopyFromRecordset upd
        .Rows(1).AutoFilter
    End With

    ShtUnlock(strSheet)

    upd.Close

End Sub
