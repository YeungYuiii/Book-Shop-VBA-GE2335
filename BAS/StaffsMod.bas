Attribute VB_Name = "StaffsMod"
Option Explicit

Public Sub StaffsMain()
Attribute StaffsMain.VB_ProcData.VB_Invoke_Func = "q\n14"
    
    Dim Sfrm As New formStaffList
    Sfrm.Show
    Set Sfrm = Nothing
    
End Sub

