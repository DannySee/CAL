Attribute VB_Name = "Main".


'Declare public project variables
Public cnn As New ADODB.Connection
Public rst As New ADODB.Recordset
Public oPrgms As New clsPrograms
Public oCst As new clsCustProfile
Public oDev As New clsDevLoads
Public oBtnPull As New clsPullCst
Public netID As String


'*******************************************************************************
'Get fresh data directly from server. Delete old records, replace with new &
'format accordingly. Runs across main tabs.
'*******************************************************************************
Sub Refresh_Data()

    'Declare sub variables
    Dim tempDct As New Scripting.Dictionary

    'Set variable to current user Network ID
    netID = Environ("Username")

    'Establish connection to SQL server
    cnn.Open _
        "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"

    'Format sheets and insert updated server data
    Format.ShtRefresh("Programs", Pull.GetPrograms)
    Format.ShtRefresh("Customer Profile"), Pull.GetCstProfile)
    Format.ShtRefresh("Deviation Loads"), Pull.GetDevLds)

    'Add drop down lists file and include data validation on Programs tab
    Format.AddDropDwns

    'Save all sheet data set to static dictionary
    Set tempDct = oPrgms.GetSaveData(True)
    Set tempDct = oCst.GetSaveData(True)
    Set tempDct = oDev.GetSaveData(True)

    'Close connection
    cnn.Close

    'Free Objects
    Set cnn = Nothing
    Set rst = Nothing
End Sub


Sub INITIALIZE_Pull_Unassigned_Customers

    'Hide any visible shapes
    Utility.Clear_Shapes

    'Update listbox
    Utility.Update_Listbox(False)

    'Show utility elements
    Utility.Show(oBtnPull.GetShapes)

End Sub


Sub SELECT_Pull_Unassigned_Customers

End Sub



Sub Download_CAL_By_Customer

End Sub


Sub Generate_Reminders

End Sub


Sub View_Active_Programs

End Sub


Sub DAB_Receipt_Validation

End Sub


Sub Reover_Deleted_Records

End Sub


Sub Pull_OpCo_List

End Sub


Sub Overlap_Validation

End Sub


Sub Item_Lookup

End Sub


Sub Request_Automation

End Sub
