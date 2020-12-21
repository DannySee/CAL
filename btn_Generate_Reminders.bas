Attribute VB_Name = "btn_Generate_Reminders"

'Declare private module constants
Private Const varShp As Variant = _Array("Cust_Add_Pane", _"Multiuse_Listbox", _
    "Cust_Add_Cancel","Cust_Add_Select","Listbox_Account_Tgl","Listbox_All", _
    "Listbox_Holder_Tgl")
Private Const strHeader As String = _
    "Hello," & vbLf & vbLf & "Please read this notification in its entirety." _
    & vbLf & vbLf & "Our records indicate that we are still missing the " _
    & "following contract(s): " & vbLf
Private Const strFooter As String = _
    vbLf & vbLf & "If you have recently submitted any of the programs above, " _
    & "please disregard this notification. We may have them in queue yet to " _
    & "be processed. " & vbLf & vbLf & "Please DO NOT REPLY OR FORWARD this " _
    & "message. This notification is auto-generated and responses are not " _
    & "visible to our team. You must send a new email to DPMSupplierContracts" _
    & "@corp.sysco.com if you are submitting or inquiring about a customer " _
    & "contract." & vbLf & vbLf & "All customer/supplier agreements will " _
    & "need to be received by the 20th of the month prior to the start date " _
    & "of the contract in order to be effective by the 1st of the month. Any " _
    & "customer/supplier agreement received after the 20th will be " _
    & "implemented with an effective date of 10 calendar days after the " _
    & "receipt date. (via DPMSupplierContracts@corp.sysco.com)" & vbLf & vbLf _
    & "Thank you," & vbLf & vbLf & "Sysco Pricing & Agreements Team"


'*******************************************************************************
'Send reminder emails to DPM hotline including email body w/ formatted list of
'expiring agreements & attached customer friendly CAL form
'*******************************************************************************
Sub Initialize()

    'Declare sub variables
    Dim wb As Workbook
    Dim varCst As Variant
    Dim strCst As String
    Dim strPth As String
    Dim strSubject As String
    Dim strBody As String
    Dim strTxt As String
    Dim strFile As String

    'Get folder path from user selection
    strPth = Utility.SelectFolder

    'Get array of my customer names
    varCst = Pull.GetCst(True)

    'Alert user and exit sub if no assigned customers/selected folder
    If Not IsEmpty(varCst) And strPth <> "" Then

        'Loop through all assigned customers
        For each cst In varCst

            'Get body of reminder email (bulleted list of expiring agmts)
            strBody = Utility.GetReminderBody(Cst)

            'If customer has any expiring agreements
            If strBody <> "" Then

                'Assemble email subject, email body and attachment path
                strSubject = cst & " " & NextMonth & " Friendly Reminder"
                strTxt = strHeader & strBody & strFooter
                strFile = strPth & cst & " CUSTOMER AGREEMENT LIST.xlsx"

                'Create customer friendly CAL Workbook and set to variable
                Set wb = Utility.DwnCstCAL(cst)

                'Save and close workbook
                wb.Close SaveChanges:=True, Filename:=strFile

                'Send email (pass through text and filename)
                SendReminder(strHeader, strTxt, strFile)
            End If
        Next
    End If
End Sub
