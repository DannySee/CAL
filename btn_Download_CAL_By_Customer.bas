Attribute VB_Name = "btn_Download_CAL_By_Customer"


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
