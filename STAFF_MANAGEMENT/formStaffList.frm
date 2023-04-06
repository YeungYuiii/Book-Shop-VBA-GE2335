VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formStaffList 
   Caption         =   "Staff"
   ClientHeight    =   6210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12660
   OleObjectBlob   =   "formStaffList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formStaffList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bottonClose_Click()
    Unload Me
End Sub

Private Sub buttonDelete_Click()
    Call DeleteRow(listboxStaff.ListIndex)
End Sub

Private Sub buttonEdit_Click()

End Sub

Private Sub bottonNew_Click()

End Sub

Private Sub listboxStaff_Click()

End Sub

Private Sub UserForm_Initialize()
    Call AddDataToListBox
End Sub

Private Sub AddDataToListBox()

    'Get the Range
    Dim rg As Range
    Set rg = GetRange()
    
    'Link the data to the LsitBox
    With listboxStaff
    
        .RowSource = rg.Address(external:=True)
        .ColumnCount = rg.Columns.Count
        .ColumnWidths = "50;65;65;60;145;70;70;70"
        .ColumnHeads = True
        .ListIndex = 0
        
    End With
    
End Sub
