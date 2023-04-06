VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formOrder_forcus 
   Caption         =   "Create a new order"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8130
   OleObjectBlob   =   "formOrder_forcus.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formOrder_forcus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 
Private Sub cmdcancel_Click()

    Unload formOrder_forcus
    Exit Sub
    
End Sub

Private Sub cmdclear_Click()
    
    UserForm_Initialize
    
End Sub

Private Sub cmdcreate_Click()

    Dim lastrow As Long
    Dim valid As Boolean
    Dim outqe As Boolean
    Dim Btt As String
    Dim xx As Integer
    Dim yy As Variant
    
    valid = False
    outqe = False
    
    If IsNumeric(txtqty) = False Then
        MsgBox "Please enter a valid number!", vbCritical, "Error"
    ElseIf CInt(txtqty) <= 0 Then
        MsgBox "Please enter a valid number!", vbCritical, "Error"
    Else
        valid = True
    End If
    
    If valid = True Then
    
        For Each cell In Worksheets("Books").UsedRange.Columns("A").Cells
            id = cell.Value
            If id = bookid Then
                Btt = Cells(cell.row, 2)
                xx = CInt(txtqty)
                yy = xx * Cells(cell.row, 7)
                If Cells(cell.row, 9) < CInt(txtqty) Then
                    MsgBox "Please enter a number smaller than storage.", vbCritical, "Sorry"
                    outqe = True
                ElseIf Cells(cell.row, 9) = 0 Then
                    MsgBox "Out of stock.", vbCritical, "Sorry"
                    out = True
                    Unload formOrder_forstaff
                Else
                    Cells(cell.row, 9) = Cells(cell.row, 9) - CInt(txtqty)
                End If
            End If
        Next
        
        If outqe = False Then
               
            Worksheets("Orders").Activate
            
            lastrow = Cells(Rows.Count, 1).End(xlUp).row + 1
        
            If lastrow <= 10 Then
                Cells(lastrow, 1).Value = "O0000" & lastrow - 1
            ElseIf lastrow <= 100 Then
                Cells(lastrow, 1).Value = "O000" & lastrow - 1
            ElseIf lastrow <= 1000 Then
                Cells(lastrow, 1).Value = "O00" & lastrow - 1
            ElseIf lastrow <= 10000 Then
                Cells(lastrow, 1).Value = "O0" & lastrow - 1
            Else
                Cells(lastrow, 1).Value = "O" & lastrow - 1
            End If
            
            Cells(lastrow, 2).Value = bookid
                    
            Cells(lastrow, 3).Value = memberidd
                    
            Cells(lastrow, 4).Value = ""
                    
            Cells(lastrow, 5).Value = Date
                    
            Cells(lastrow, 6).Value = ""
                    
            Cells(lastrow, 7).Value = CInt(txtqty)
            

            
            MsgBox "Order created." & vbCrLf & "Thank you!" & vbCrLf & vbCrLf & _
            "Order Details:" & vbCrLf & "Order ID: " & Cells(lastrow, 1).Value & vbCrLf & _
            "Book Title: " & Btt & vbCrLf & _
            "Quantity: " & xx & vbCrLf & _
            "Price: $" & yy & vbCrLf & _
            "Order Date: " & Date, vbInformation, "Success"
            
            
            Unload Me
            
            Call Reloading
            
        End If
    End If
    
End Sub

Private Sub UserForm_Initialize()

    txtqty = ""
    
End Sub
