VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Login"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8910.001
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Clear_Click()

    Me.Account.Value = ""
    Me.Password.Value = ""
    
    Me.Account.SetFocus
    
    End Sub

Private Sub GuestLogin_Click()

Application.Visible = False

Sheets("Books").Visible = False

Sheets("Publishers").Visible = False

Sheets("adminonly").Visible = True

Sheets("Staffs").Visible = False

Sheets("Members").Visible = False

Sheets("Orders").Visible = False

Sheets("SalesTable").Visible = False

Unload Me

End Sub

Private Sub Login_Click()

Dim acc As String
Dim pw As String
Dim xTitle As String
Dim i As Integer

xTitle = "Login"

acc = Me.Account.Value

pw = Me.Password.Value


Dim xInputBlank As String: xInputBlank = ""

If acc = "" Then xInputBlank = xInputBlank & "Username" & vbNewLine

If pw = "" Then xInputBlank = xInputBlank & "Password" & vbNewLine

 

If xInputBlank <> "" Then

MsgBox "below fields are blank." & vbNewLine & xInputBlank, vbCritical, xTitle

Else

Dim xWs As Worksheet

Set xWs = ThisWorkbook.Sheets("Members")

Dim xizt As Boolean

xizt = False

Dim wPassword As String

wPassword = ""

Dim a As Integer

For a = 2 To xWs.Cells(Rows.Count, 2).End(xlUp).row

If xWs.Range("I" & a).Value = acc Then

xizt = True

wPassword = xWs.Range("H" & a).Value

LoginAc = xWs.Range("I" & a).Value

memberidd = xWs.Range("A" & a).Value

LoginPw = xWs.Range("H" & a).Value


Exit For

End If

Next a

 

If xizt = False Then

MsgBox "Invalid account.", vbCritical, xTitle

acc = ""

pw = ""

Me.Account.SetFocus

 

Else



If pw <> wPassword Then

MsgBox "UserName Or Password Incorrect.", vbCritical, xTitle

acc = ""

pw = ""

Me.Account.SetFocus

Else


Unload Me


Application.Visible = False

Sheets("Books").Visible = False

Sheets("adminonly").Visible = True

Sheets("Staffs").Visible = False

Sheets("Publishers").Visible = False

Sheets("Members").Visible = False

Sheets("Orders").Visible = False

Sheets("SalesTable").Visible = False

Unload Me

End If

End If

End If



End Sub

Private Sub SignUp_Click()

SignUpform.Show
Unload Me

End Sub

Private Sub StaffLogin_Click()

Dim acc As String
Dim pw As String
Dim xTitle As String

xTitle = "Login"

acc = Me.Account.Value

pw = Me.Password.Value



If (acc = "manager" And pw = "manager") Then

Unload Me

Application.Visible = True

Sheets("Books").Visible = True

Sheets("Publishers").Visible = True

Sheets("Staffs").Visible = True

Sheets("Members").Visible = True

Sheets("Orders").Visible = True

Sheets("SalesTable").Visible = True

admincome = True

Exit Sub

End If



Dim xInputBlank As String: xInputBlank = ""

If acc = "" Then xInputBlank = xInputBlank & "Username" & vbNewLine

If pw = "" Then xInputBlank = xInputBlank & "Password" & vbNewLine

 

If xInputBlank <> "" Then

MsgBox "below fields are blank." & vbNewLine & xInputBlank, vbCritical, xTitle

Else

Dim xWs As Worksheet

Set xWs = ThisWorkbook.Sheets("Staffs")

Dim xizt As Boolean

xizt = False

Dim wPassword As String

wPassword = ""

Dim a As Integer

For a = 2 To xWs.Cells(Rows.Count, 2).End(xlUp).row

If xWs.Range("I" & a).Value = acc Then

xizt = True

wPassword = xWs.Range("H" & a).Value


Exit For

End If

Next a

 

If xizt = False Then

MsgBox "Invalid account.", vbCritical, xTitle

acc = ""

pw = ""

Me.Account.SetFocus

 

Else



If pw <> wPassword Then

MsgBox "UserName Or Password Incorrect.", vbCritical, xTitle

acc = ""

pw = ""

Me.Account.SetFocus

Else

Unload Me

Application.Visible = True

Sheets("Books").Visible = True

Sheets("Publishers").Visible = True

Sheets("Staffs").Visible = False

Sheets("Members").Visible = False

Sheets("Orders").Visible = True

Sheets("adminonly").Visible = False

Sheets("SalesTable").Visible = False

admincome = False

Unload Me

End If

End If

End If


End Sub

Private Sub UserForm_Initialize()

    Me.Account.Value = ""
    Me.Password.Value = ""
    
    Me.Account.SetFocus

End Sub
