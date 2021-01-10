Attribute VB_Name = "btn_Overlap_Validation"

'Declare private module constants
Private Const varShp As Variant = Array("Overlap_Pane", _"Overlap_Opco", _
    "Overlap_VA","Overlap_CA", "Overlap_Textbox_1", "Overlap_Textbox_2", _
    "Overlap_Cancel","Overlap_Validate")
Private Const varHeaders As Variant = Array("PRN/GRP", "SHP", "ITEM")


'*******************************************************************************
'Show all utility elements, update and resize listbox.
'*******************************************************************************
Private Sub Overlap_Validation_Initialize()

    'Hide any visible shapes
    Utility.ClearShapes

    'Show utility elements
    Utility.Show(varShp)
End Sub


'*******************************************************************************
'Pull in selected customer OpCo data. Output formatted report on new workbook.
'*******************************************************************************
Private Sub Overlap_Validation_Select()

    'Declare sub variables
    Dim strOp As String
    Dim strOv1 As String
    Dim strOv2 As String
    Dim blVA As Boolean

    'Get user input
    strOp = Trim(Sheets("Control Panel".TextBoxes("Overlap_OpCo").Text)
    strOv1 = Trim(Sheets("Control Panel".TextBoxes("Overlap_Textbox_1").Text)
    strOv2 = Trim(Sheets("Control Panel".TextBoxes("Overlap_Textbox_2").Text)

    'Get overlap type
    If Sheets("Control Panel").Shapes("Overlap_VA").ControlFormat.Value = _
        xlOn Then blVA = True

    'Ensure there is no missing input data
    If strOp <> "" And strOv1 <> "" And strOv2 <> "" Then

        'Create new workbook
        Utility.CreateWorkbook("Report")

        'Add report headers
        Utility.AddHeaders(varHeaders)

        'If VA overlap
        If blVA = True Then

            'Paste customer overlaps
            Utility.PasteOverlap(Pull.GetVaOvCst(strOv1, strOv2, strOp))
            Range("C2").CopyFromRecordset Pull.GetVaOvItm(strOv1, strOv2, strOp)

        'If CA overlap
        Else

            'Paste customer overlaps
            Utility.PasteOverlap(Pull.GetCaOvCst(strOv1, strOv2, strOp))
            Range("C2").CopyFromRecordset Pull.GetCaOvItm(strOv1, strOv2, strOp)
        End If

        'Add borders
        Utility.AddBorders

        'Clear all utility elements
        Utility.ClearShapes

    'If missing data
    Else

        'Alert user of missing data
        msgbox "Missing Data:" & vbNewLine _
            & "You must include OpCo and two agreement numbers."
    End if
End Sub
