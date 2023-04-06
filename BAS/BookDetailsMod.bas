Attribute VB_Name = "BookDetailsMod"
Option Explicit

'Include in Declarations section of module.
Public bookid As String
Public bookname As String
Public title As String
Public finalRow As Integer
Public finalR As Integer
Public SearchBT(99) As String
Public SearchBid(99) As String
Public sh_k As Integer
Public Category(49) As String
Public Publisher(49) As String
Public memberidd As String

Public Sub BookDetailsMain()
    
    Dim wsBk As Worksheet
    Dim wsPh As Worksheet
    
    Worksheets("Books").Activate
    Set wsBk = ThisWorkbook.Worksheets("Books")
    Set wsPh = ThisWorkbook.Worksheets("Publishers")
    finalRow = wsBk.Cells(Rows.Count, "A").End(xlUp).row
    finalR = wsPh.Cells(Rows.Count, "A").End(xlUp).row
    
    Call Category_Initialize
    Call Publisher_Initialize
    Call CallBookList
    
End Sub

Sub CallBookList()

    formBookList.Show
    
End Sub

Sub CallBookDetails()
    
    Dim BDfrm As New formBookDetails
    BDfrm.Show
    Set BDfrm = Nothing
    
End Sub

Sub CallOrderform()
    
    Dim Ofrm As New formOrder_forcus
    Ofrm.Show
    Set Ofrm = Nothing
    
End Sub

Sub Publisher_Initialize()

    Dim i As Integer
    Dim j As Integer
    
'initialize Publisher array
    Dim finalR As Integer
    Worksheets("Publishers").Activate
    finalR = Cells(Rows.Count, "A").End(xlUp).row
    For i = 2 To finalR
        Publisher(j) = Cells(i, 2).Value
        j = j + 1
    Next i
End Sub

Sub Category_Initialize()

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim rep As Boolean
    
'initialize Category array
    rep = False
    For i = 2 To finalRow
        For j = 0 To 49
            If Cells(i, 6).Text = Category(j) Then
                rep = True
                Exit For
            End If
        Next j
        If rep = False Then
            Category(k) = Cells(i, 6).Value
            k = k + 1
        End If
        rep = False
    Next i
End Sub

Sub Reloading()

    Unload formBookDetails

End Sub
