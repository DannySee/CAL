Attribute VB_Name = "btn_Pull_OpCo_List"

'Declare private module constants
Private Const varShp As Variant = _Array("Listbox_Pane", _"Multiuse_Listbox", _
    "Listbox_Cancel","Listbox_Select")
Private Const varHeaders As Variant = _Array("OPCO LIST")


'*******************************************************************************
'Show all utility elements, update and resize listbox.
'*******************************************************************************
Private Sub Pull_OpCo_List_Initialize()

    'Hide any visible shapes
    Utility.ClearShapes

    'Set default listbox view
    Utility.ListboxByCst(True)

    'Show utility elements
    Utility.Show(varShp)

    'Assign Select button to correct routine
    Sheets("Control Panel").Shapes("Listbox_Select").OnAction = _
        "Pull_Unassigned_Customers_Select"
End Sub


'*******************************************************************************
'Pull in selected customer OpCo data. Output formatted report on new workbook.
'*******************************************************************************
Private Sub Pull_OpCo_List_Select()

    'Declare sub variables
    Dim varCst As Variant
    Dim varPacket As String
    Dim strPacket As String
    Dim varOpCo As Variant
    Dim varOpList As Variant
    Dim i As Integer

    'Get Selected customers and selected packet
    varCst = Utility.GetSelection

    'If a customer was selected from list
    If Not IsEmpty(varCst) Then

        'Get array of customer packets
        varPacket = Pull.GetCstPacket(GetStr(varCst, True))

        'If customer packets were located
        If Not IsEmpty(varPacket) Then

            'Get SQL string of customer packets
            strPacket = GetStr(varPacket, True)

            'Get array of servicing OpCos
            varOpCo = Pull.GetOpcos(strPacket)

            'Loop through multidimensional array (rows)
            For i = 0 To UBound(varOpCo, 2)

                'Create new workbook for report population
                Utility.NewSheet(varOpCo(0,i))

                'Add headers to new report
                Utility.AddHeaders(varHeaders)

                'Add OpCo list to sheet
                Utility.PasteList(varOpCo(1, i), 3, 1)

                'Add borders to sheet
            Next

            'Clear utility shapes
            Utility.ClearShapes

        'If no customer packets were found
        Else

            'Alert user and exit sub if no customers were selected
            msgbox "Selected customer(s) has missing packet name. " & vblf _
                & " Please validate the PACKET column on the Customer " _
                & "Profile tab and try again."

    'If no customers were selected from list
    Else

        'Alert user and exit sub if no customers were selected
        msgbox "You must make at least one selection."
    End If


End Sub
