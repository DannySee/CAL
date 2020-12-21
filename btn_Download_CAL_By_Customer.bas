Attribute VB_Name = "btn_Download_CAL_By_Customer"

'Declare private module constants
Private Const varShp As Variant = _Array("Cust_Add_Pane", _"Multiuse_Listbox", _
    "Cust_Add_Cancel","Cust_Add_Select","Listbox_Account_Tgl","Listbox_All", _
    "Listbox_Holder_Tgl")
    

'*******************************************************************************
'Download CAL account assignments by customer.
'*******************************************************************************
Sub Initialize()

    'Declare sub variables
    Dim varCst As Variant
    Dim strCst As String
    Dim strPth As String
    Dim wb As Workbook

    'Get folder path from user selection
    strPth = Utility.SelectFolder

    'Get array of my customer names
    varCst = Pull.GetCst(True)

    'Alert user and exit sub if no assigned customers/selected folder
    If Not IsEmpty(varCst) And strPth <> "" Then

        'Loop through all assigned customers
        For each cst In varCst

            'Create customer friendly CAL Workbook and set to variable
            Set wb = Utility.DwnCstCAL(cst)

            'Save and close workbook
            wb.Close SaveChanges:=True, Filename:= _
                strPth & cst & " CUSTOMER AGREEMENT LIST.xlsx"
        Next
    End If
End Sub
