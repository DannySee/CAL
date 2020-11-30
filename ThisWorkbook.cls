VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "clsPrograms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'*******************************************************************************
'Upload/inserts new worksheet records to the SQL server. Refresh static
'dictionary after upload is complete.
'*******************************************************************************
Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    'Delcare sub variables
    Dim varPgrms As Variant
    Dim varCst As Variant
    Dim carDev As Variant

    varPgrms = Compare(dctPrograms(False), GetXL("Programs$"))
    varCst = Compare(dctPrograms(False), GetXL("Programs$"))
    carDev = Compare(dctPrograms(False), GetXL("Programs$"))

    If varChange(0)(0) <> 0 Or varChange(1)(0) <> 0 Then

        cnn.Open _
            "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

        If varChange(0)(0) <> 0 Then UploadChanges (varChange(0))
        If varChange(1)(0) <> 0 Then RetVal = InsertNew(varChange(1), varChange(2))

        cnn.Close
        Set cnn = Nothing
    End If


End Sub
