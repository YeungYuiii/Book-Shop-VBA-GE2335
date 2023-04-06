VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub Workbook_Open()

    LoginInstance = 0
    
    Application.Visible = False
    
    Call FirstLogin
    
End Sub

Sub FirstLogin()
    admincome = False
    LoginForm.Show
    Call BookDetailsMain
End Sub

