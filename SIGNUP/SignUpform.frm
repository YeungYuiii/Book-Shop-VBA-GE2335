VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SignUpform 
   Caption         =   "Sign Up"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9330.001
   OleObjectBlob   =   "SignUpform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SignUpform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdcancel_Click()

Unload Me

End Sub

Private Sub cmdSignUp_Click()

Dim X As Long

Dim Y As Worksheet

Dim res As Integer



For res = 1 To Columns.Count

If PhoneBox.Text = shMem.Cells(4, res) Then

MsgBox "phone number registered"

Exit Sub

End If

Next res

For res = 1 To Columns.Count

If EmailBox.Text = shMem.Cells(5, res).Text Then

MsgBox "Email registered"

Exit Sub

End If

Next res

For res = 1 To Columns.Count

If UsernameBox.Text = shMem.Cells(9, res) Then

MsgBox "username registered"

Exit Sub

End If

Next res



Set Y = shMem
 

X = Y.Range("A" & Rows.Count).End(xlUp).row + 1

 

With Y
If X <= 10 Then
        .Cells(X, 1).Value = "M0000" & X - 1
    ElseIf X <= 100 Then
        .Cells(X, 1).Value = "M000" & X - 1
    ElseIf X <= 1000 Then
        .Cells(X, 1).Value = "M00" & X - 1
    ElseIf X <= 10000 Then
        .Cells(X, 1).Value = "M0" & X - 1
    Else
        .Cells(X, 1).Value = "M" & X - 1
    End If

.Cells(X, 2).Value = FirstNameBox.Text

.Cells(X, 3).Value = LastNameBox.Text

.Cells(X, 4).Value = PhoneBox.Text

.Cells(X, 5).Value = EmailBox.Text

.Cells(X, 6).Value = DistrictBox.Text

.Cells(X, 7).Value = Date

.Cells(X, 8).Value = PasswordBox.Text

.Cells(X, 9).Value = UsernameBox.Text

End With

MsgBox "You are now a member!", vbInformation

Unload Me

End Sub


Private Sub EmailBox_Change()

Dim Mrow As Integer

On Error Resume Next

Mrow = Application.WorksheetFunction.Match(CStr(Me.EmailBox), shMem.Range("E:E"), 0)

If Mrow >= 1 Then


MsgBox "Email Already Registered"

End If

End Sub

Private Sub FirstNameBox_Change()

 

On Error Resume Next

Me.FirstNameBox = Format(StrConv(Me.FirstNameBox, vbProperCase))

 

End Sub

Private Sub LastNameBox_Change()

 

On Error Resume Next

Me.LastNameBox = Format(StrConv(Me.LastNameBox, vbProperCase))

 

End Sub

Private Sub PhoneBox_Change()

Dim Mrow As Long

On Error Resume Next

Mrow = WorksheetFunction.Match(Me.PhoneBox.Value, shMem.Range("D:D"), 0)

If Mrow >= 1 Then


MsgBox "Number Already Registered"

End If

End Sub

Private Sub UsernameBox_Change()

Dim Mrow As Variant

On Error Resume Next

Mrow = Application.WorksheetFunction.Match(Me.UsernameBox.Text, shMem.Range("I:I"), 0)

If Mrow >= 1 Then


MsgBox "Username Already Registered"

End If

End Sub

Private Sub PasswordBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)

On Error Resume Next

If Me.PasswordBox = Me.Password2Box And Me.Password2Box <> "" Then

Me.PasswordLab.Caption = ""

Me.cmdSignUp.Enabled = True

Me.cmdSignUp.SetFocus

Else

Me.PasswordLab = "Password Not Matching!"

Me.cmdSignUp.Enabled = False

End If

End Sub

Private Sub Password2Box_Change()
On Error Resume Next

If Me.PasswordBox = Me.Password2Box And Me.Password2Box <> "" Then

Me.PasswordLab.Caption = ""

Me.cmdSignUp.Enabled = True

Me.cmdSignUp.SetFocus

Else

Me.PasswordLab = "Password Not Matching!"

Me.cmdSignUp.Enabled = False

End If
End Sub

Private Sub UserForm_Initialize()


DistrictBox.AddItem "Central & Western"
DistrictBox.AddItem "Eastern"
DistrictBox.AddItem "Islands"
DistrictBox.AddItem "Kowloon City"
DistrictBox.AddItem "Kwai Tsing"
DistrictBox.AddItem "Kwun Tong"
DistrictBox.AddItem "North"
DistrictBox.AddItem "Sai Kung"
DistrictBox.AddItem "Sha Tin"
DistrictBox.AddItem "Sham Shui Po"
DistrictBox.AddItem "Southern"
DistrictBox.AddItem "Tai Po"
DistrictBox.AddItem "Tsuen Wan"
DistrictBox.AddItem "Tuen Mun"
DistrictBox.AddItem "Wan Chai"
DistrictBox.AddItem "Wong Tai Sin"
DistrictBox.AddItem "Yuen Long"
DistrictBox.AddItem "Yau Tsim Wong"

End Sub

