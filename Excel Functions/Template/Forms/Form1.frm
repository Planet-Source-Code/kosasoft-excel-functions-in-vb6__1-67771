VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Create"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cExcel As New ExcelFunc.cFunctions
Dim oExcel As New Excel.Application
Dim oWBook As Excel.Workbook

Private Sub Command1_Click()
'    Set oExcel = New Excel.Application
'    oExcel.Workbooks.Add
'    oExcel.Visible = True
'
'    With oExcel.Range("A1").Select
'        oExcel.ActiveCell.FormulaR1C1 = "asdasd"
'    End With
    Set oExcel = cExcel.NewExcel(oExcel)
    cExcel.CellText oExcel, 2, 2, "testing"
    cExcel.FormatFont oExcel, 2, 2, 20, True, True, 0, "Arial Narrow", vbGreen
    cExcel.FormatCell oExcel, 2, 2, True, vbRed
    cExcel.MergeCell oExcel, 1, 1, 5
    cExcel.CellText oExcel, 1, 1, "header"
    cExcel.FormatFont oExcel, 1, 1, 10, True, False, 1, , , 1
    cExcel.PutGridLines oExcel, 1, 5, 1, 5, 9
End Sub

Private Sub Command2_Click()
    Set oExcel = cExcel.OpenExcel(oExcel, "C:\", "testing.xls")
    cExcel.CellText oExcel, 2, 2, "testing"
    cExcel.FormatFont oExcel, 2, 2, 20, True, True, 0, "Arial Narrow"
End Sub
