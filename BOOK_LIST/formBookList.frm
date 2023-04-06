VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formBookList 
   Caption         =   "Book Searching"
   ClientHeight    =   13515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18000
   OleObjectBlob   =   "formBookList.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "formBookList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CB1_Click()

    sh_k = 0
    SBB1.Value = 0
    Dim i As Integer
    For i = 0 To 99
        SearchBT(i) = ""
        SearchBid(i) = ""
    Next i
    
    Call GeneralSearch
    Call TitleSearch
    Call CategorySearch
    Call PublisherSearch
    
    If OB1.Value = True Then
        Call StockWellSearch
    End If
    
    Call ShowSearch
    
End Sub

Private Sub CBLogin_Click()
    LoginForm.Show
End Sub

Private Sub CommandButton1_Click()

    sh_k = 0
    SBB1.Value = 0
    Dim i As Integer
    For i = 0 To 99
        SearchBT(i) = ""
        SearchBid(i) = ""
    Next i
    
    Worksheets("Books").Activate
    
    For i = 2 To finalRow
        SearchBT(j) = Cells(i, 2).Value
        SearchBid(j) = Cells(i, 1).Value
        j = j + 1
    Next i

    Call ShowSearch
End Sub

Private Sub Label25_Click()
    bookname = Label25.Caption
    Call AllLabel_Click
End Sub

Private Sub Label26_Click()
    bookname = Label26.Caption
    Call AllLabel_Click
End Sub

Private Sub Label27_Click()
    bookname = Label27.Caption
    Call AllLabel_Click
End Sub

Private Sub Label28_Click()
    bookname = Label28.Caption
    Call AllLabel_Click
End Sub

Private Sub Label29_Click()
    bookname = Label29.Caption
    Call AllLabel_Click
End Sub

Private Sub Label30_Click()
    bookname = Label30.Caption
    Call AllLabel_Click
End Sub



Private Sub Label7_Click()
    bookid = Cells(finalRow, 1).Value
    Call CallBookDetails
End Sub

Private Sub SBB1_Change()
    sh_k = SBB1.Value * 6
    Call ShowSearch
End Sub

Private Sub UserForm_Initialize()
    
    If LoginAc <> "" Then
    Label32.Caption = "Welcome! " & LoginAc
    Else
    Label32.Caption = "Welcome!"
    End If
    Label7.Caption = "Capital: A Critique of Political Economy, Vol 1"
    
    Worksheets("Books").Activate
    finalRow = Cells(Rows.Count, "A").End(xlUp).row
    
    Dim path As String
    Dim Inv As Variant
    Dim i As Integer
    Dim j As Integer
    
'initialize ComboBox1(Category) and ComboBox2(Publisher)
    For i = 0 To 49
        If Category(i) <> "" Then
            ComboBox1.AddItem Category(i)
        End If
    Next i
    
    For i = 0 To 49
        If Publisher(i) <> "" Then
            ComboBox2.AddItem Publisher(i)
        End If
    Next i

'initialize New arrival
    path = ThisWorkbook.path & "\BookCover\" & Cells(finalRow, 1) & ".JPG"
    Image12.Picture = LoadPicture(path)
    
    Label7.Caption = Cells(finalRow, 2)
    Label16.Caption = "$" & Cells(finalRow, 7)
    
    Inv = Cells(finalRow, 9)
    If Inv > 0 Then
        Label22.ForeColor = &HFF8080
        Label22.Caption = "enough"
    Else
        Label22.ForeColor = &HFF&
        Label22.Caption = "shortage"
        Image15.Picture = LoadPicture(ThisWorkbook.path & "\Pic\not-tick.JPG")
    End If
    
'initialize

    Worksheets("Books").Activate
    
    For i = 2 To finalRow
        SearchBT(j) = Cells(i, 2).Value
        SearchBid(j) = Cells(i, 1).Value
        j = j + 1
    Next i

    Call ShowSearch

    
End Sub

Sub AllLabel_Click()

    If bookname = "" Then
        Exit Sub
    End If
    
    Dim i As Integer
    
    i = 1
    For i = 1 To finalRow
        If Cells(i, 2).Value = bookname Then
            bookid = Cells(i, 1).Value
        End If
    Next i
    
    Call CallBookDetails
    
End Sub


Sub GeneralSearch()

    If SearchBox = "" Then
        Exit Sub
    End If

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    Worksheets("Books").Activate
    
    For i = 2 To finalRow
        For j = 1 To 3
            If j = 3 Then
                j = j + 3
            End If
            If InStr(1, Cells(i, j).Value, SearchBox, 1) <> 0 Then
                SearchBT(k) = Cells(i, 2).Value
                SearchBid(k) = Cells(i, 1).Value
                k = k + 1
            End If
        Next j
    Next i
    
    Call de_rep

End Sub
Sub TitleSearch()

    If TextBox1 = "" Then
        Exit Sub
    End If
    
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 99
        If SearchBT(i) = "" Then
            Exit For
        End If
    Next i
    
    For j = 2 To finalRow
        If InStr(1, Cells(j, 2).Value, TextBox1.Value, 1) <> 0 Then
            SearchBT(i) = Cells(j, 2).Value
            SearchBid(i) = Cells(j, 1).Value
            i = i + 1
        End If
    Next j
    
    Call de_rep
    
End Sub
Sub CategorySearch()

    If ComboBox1 = "" Then
        Exit Sub
    End If
    
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 98
        If SearchBT(i) = "" Then
            Exit For
        End If
    Next i
    
    For j = 2 To finalRow
        If Cells(j, 6).Value = ComboBox1.Value Then
            SearchBT(i) = Cells(j, 2).Value
            SearchBid(i) = Cells(j, 1).Value
            i = i + 1
        End If
    Next j
    
    Call de_rep

End Sub

Sub PublisherSearch()
    
    If ComboBox2 = "" Then
        Exit Sub
    End If
    
    Dim i As Integer
    Dim j As Integer
    Dim Pid As String
    
    Worksheets("Publishers").Activate
    finalR = Cells(Rows.Count, "A").End(xlUp).row
    
    For i = 0 To 98
        If SearchBT(i) = "" Then
            Exit For
        End If
    Next i
    
    For j = 2 To finalR
        If Cells(j, 2).Value = ComboBox2.Value Then
            Pid = Cells(j, 1).Value
        End If
    Next j
    
    Worksheets("Books").Activate
    For j = 2 To finalRow
        If Cells(j, 4).Value = Pid Then
            SearchBT(i) = Cells(j, 2).Value
            SearchBid(i) = Cells(j, 1).Value
            i = i + 1
        End If
    Next j
    
    Call de_rep
    
End Sub

Sub StockWellSearch()

    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 98
        For j = 2 To finalRow
            If SearchBT(i) = Cells(j, 2) Then
                If Cells(j, 9) = 0 Then
                    For k = i + 1 To 97
                        SearchBT(i) = SearchBT(k)
                        SearchBid(i) = SearchBid(k)
                    Next k
                    SearchBT(98) = ""
                    SearchBid(98) = ""
                End If
                Exit For
            End If
        Next j
    Next i
    
End Sub

Sub de_rep()

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    
    For i = 0 To 98
        For j = i + 1 To 98
            If SearchBT(i) = SearchBT(j) And i <> 99 And SearchBT(i) <> "" And SearchBT(j) <> "" Then
                For k = j To 97
                    SearchBT(k) = SearchBT(k + 1)
                    SearchBid(k) = SearchBid(k + 1)
                Next k
                SearchBT(98) = ""
            End If
        Next j
    Next i
    
End Sub

Sub ShowSearch()
    Dim i As Integer
    
    On Error GoTo PicErr
    
    Label25.Caption = SearchBT(sh_k)
    If SearchBid(sh_k) <> "" Then
        i = 1
        Image20.Picture = LoadPicture(ThisWorkbook.path & "\BookCover\" & SearchBid(sh_k) & ".JPG")
    Else
        Image20.Picture = LoadPicture(vbNullString)
    End If
    
    Label26.Caption = SearchBT(sh_k + 1)
Pic1:
    If SearchBid(sh_k + 1) <> "" Then
        i = 2
        Image21.Picture = LoadPicture(ThisWorkbook.path & "\BookCover\" & SearchBid(sh_k + 1) & ".JPG")
    Else
        Image21.Picture = LoadPicture(vbNullString)
    End If
    
    Label27.Caption = SearchBT(sh_k + 2)
Pic2:
    If SearchBid(sh_k + 2) <> "" Then
        i = 3
        Image22.Picture = LoadPicture(ThisWorkbook.path & "\BookCover\" & SearchBid(sh_k + 2) & ".JPG")
    Else
        Image22.Picture = LoadPicture(vbNullString)
    End If
    
    Label28.Caption = SearchBT(sh_k + 3)
Pic3:
    If SearchBid(sh_k + 3) <> "" Then
        i = 4
        Image23.Picture = LoadPicture(ThisWorkbook.path & "\BookCover\" & SearchBid(sh_k + 3) & ".JPG")
    Else
        Image23.Picture = LoadPicture(vbNullString)
    End If
    
    Label29.Caption = SearchBT(sh_k + 4)
Pic4:
    If SearchBid(sh_k + 4) <> "" Then
        i = 5
        Image24.Picture = LoadPicture(ThisWorkbook.path & "\BookCover\" & SearchBid(sh_k + 4) & ".JPG")
    Else
        Image24.Picture = LoadPicture(vbNullString)
    End If
    Label30.Caption = SearchBT(sh_k + 5)
    
Pic5:
    If SearchBid(sh_k + 5) <> "" Then
        i = 6
        Image25.Picture = LoadPicture(ThisWorkbook.path & "\BookCover\" & SearchBid(sh_k + 5) & ".JPG")
    Else
        Image25.Picture = LoadPicture(vbNullString)
    End If
Exit Sub

PicErr:
    Select Case i
        Case 1
            Image20.Picture = LoadPicture(ThisWorkbook.path & "\BookCover\B0.JPG")
            GoTo Pic1
        Case 2
            Image21.Picture = LoadPicture(ThisWorkbook.path & "\BookCover\B0.JPG")
            GoTo Pic2
        Case 3
            Image22.Picture = LoadPicture(ThisWorkbook.path & "\BookCover\B0.JPG")
            GoTo Pic3
        Case 4
            Image23.Picture = LoadPicture(ThisWorkbook.path & "\BookCover\B0.JPG")
            GoTo Pic4
        Case 5
            Image24.Picture = LoadPicture(ThisWorkbook.path & "\BookCover\B0.JPG")
            GoTo Pic5
        Case 6
            Image25.Picture = LoadPicture(ThisWorkbook.path & "\BookCover\B0.JPG")
    End Select

End Sub


