VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formOrder_forstaff 
   Caption         =   "Create a new order"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5415
   OleObjectBlob   =   "formOrder_forstaff.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formOrder_forstaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdcancel_Click()

    Unload formOrder_forstaff
    Exit Sub
    
End Sub

Private Sub cmdclear_Click()

    UserForm_Initialize
    
End Sub

Private Sub cmdcreate_Click()
    
    Dim lastrow As Long
    Dim cell As Range
    Dim id As String
    Dim out As Boolean

    Worksheets("Orders").Activate
    Worksheets("Books").Activate
    
    For Each cell In Worksheets("Books").UsedRange.Columns("A").Cells
        id = cell.Value
        If id = txtbookid Then
            If Cells(cell.row, 9) < CInt(txtqty) Then
                MsgBox "Not enough storage.", vbCritical, "Error"
            ElseIf Cells(cell.row, 9) = 0 Then
                MsgBox "Out of stock.", vbCritical, "Error"
                out = True
                Unload formOrder_forstaff
            Else
                Cells(cell.row, 9) = Cells(cell.row, 9) - CInt(txtqty)
            End If
        End If
    Next
    
    With Worksheets("Orders")
    
    lastrow = .Cells(Rows.Count, 1).End(xlUp).row + 1

    If lastrow <= 10 Then
        .Cells(lastrow, 1).Value = "O0000" & lastrow - 1
    ElseIf lastrow <= 100 Then
        .Cells(lastrow, 1).Value = "O000" & lastrow - 1
    ElseIf lastrow <= 1000 Then
        .Cells(lastrow, 1).Value = "O00" & lastrow - 1
    ElseIf lastrow <= 10000 Then
        .Cells(lastrow, 1).Value = "O0" & lastrow - 1
    Else
        .Cells(lastrow, 1).Value = "O" & lastrow - 1
    End If

    .Cells(lastrow, 2).Value = txtbookid
            
    .Cells(lastrow, 3).Value = txtmemid
            
    .Cells(lastrow, 4).Value = txtstaffid
            
    .Cells(lastrow, 5).Value = Date
            
    .Cells(lastrow, 6).Value = txtdisc & "%"
            
    .Cells(lastrow, 7).Value = CInt(txtqty)
    End With
    
    If out = False Then
        MsgBox "Order created!", vbInformation, "Success"
    End If
    
    Unload formOrder_forstaff
   
End Sub

Private Sub UserForm_Initialize()

    txtbookid = ""
    txtmemid = ""
    txtstaffid = ""
    txtdisc = ""
    txtqty = ""
    
End Sub
