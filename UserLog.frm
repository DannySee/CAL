VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserLog 
   Caption         =   "SUS Login"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5565
   OleObjectBlob   =   "UserLog.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "UserLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Cancel_Click()

    'Close userform
    Unload Me
    End

End Sub

Private Sub LoginButton_Click()

    'Declare variables
    Dim cnn     As New ADODB.Connection
    Dim strPKey As String
    
    'Establish connection to SQL Server
    cnn.Open "DRIVER=SQL Server;SERVER=MS440CTIDBPC1;DATABASE=Pricing_Agreements;"
    
    'Delete old credentials
    cnn.Execute "DELETE FROM Login_Cred WHERE NET_ID = '" & Environ("Username") & "'"
    
    'Get encrypted password & key
    strPKey = Enc_Pwd(PWD.value)
    
    'Insert new credentials
    cnn.Execute "INSERT INTO Login_Cred (NET_ID, SUS_ID, CRED_ID, SUS_PWD) VALUES('" & Environ("Username") & "','" & UCase(UID.value) & "','" & strPKey
    
    'unload login page
    Unload Me
    
End Sub

Function Enc_Pwd(strOG As String) As String

    'Declare variables
    Dim strPwd  As String
    Dim StrKey  As String
    Dim i       As Integer
    Dim iRnd    As Integer
    
    'iterate through each character of password
    For i = 1 To Len(strOG)
        
        'Get random int from 0 to 9
        Randomize
        iRnd = Int((9 - 1 + 1) * Rnd + 1)
        
        'Assemble encrypted password and key
        StrKey = StrKey & iRnd
        strPwd = strPwd & Chr(Asc(Mid(strOG, i, 1)) + iRnd)
    Next
    
    'Return end of query string
    Enc_Pwd = StrKey & "','" & strPwd & "')"
    
End Function

