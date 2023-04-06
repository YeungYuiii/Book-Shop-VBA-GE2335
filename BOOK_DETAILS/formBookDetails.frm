VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formBookDetails 
   Caption         =   "Book Details"
   ClientHeight    =   8205.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14415
   OleObjectBlob   =   "formBookDetails.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formBookDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CB1_Click()
    Dim i As Integer
    i = get_i(bookid)
    If Cells(i, 9) < 1 Then
        MsgBox "You can put it into cart still, " & _
        "we will prepare your book!", vbOKOnly + _
        vbInformation, "Sorry, Inventory Shortage!"
    Else
        Call CallOrderform
    End If
End Sub

Private Sub UserForm_Initialize()

    Dim path As String
    Dim i As Integer
    Dim pubid As String
    
    On Error GoTo Errpic
    path = ThisWorkbook.path & "\BookCover\" & bookid & ".JPG"
    Image1.Picture = LoadPicture(path)
Regoing:
    i = get_i(bookid)
    
    Label6.Caption = Cells(i, 2)
    Label7.Caption = Cells(i, 6)
    Label8.Caption = Cells(i, 3)
    Label10.Caption = Cells(i, 5)
    Label11.Caption = "$ " & Cells(i, 7)
    Label18.Caption = Cells(i, 9)
    
    If Cells(i, 9) < 1 Then
        With CB1
        .Left = CB1.Left - 10
        .Width = CB1.Width + 20
        .ForeColor = &HFF&
        .BackColor = &H808080
        .Caption = "Inventory Shortage !"
        .MousePointer = 12
        End With
    End If
    
    Label9.Caption = 1
Exit Sub
Errpic:
    Image1.Picture = LoadPicture(ThisWorkbook.path & "\BookCover\B0.JPG")
    Resume Regoing
    
End Sub

Function get_i(bid As String) As Integer

    Dim i As Integer
    
    For i = 2 To finalRow
        If Cells(i, 1).Value = bid Then
            Exit For
        End If
    Next i
    
    get_i = i
    
End Function

