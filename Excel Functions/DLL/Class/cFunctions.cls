VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'*******************************************************************
'*******************************************************************
'****   Programmed by:      Raymond Uy                          ****
'****   Email:              uy.raymond@gmail.com                ****
'****   Note:               Send me feedback regarding          ****
'****                       this application                    ****
'****                       Don't forget to vote.               ****
'****                       If you're planning to use           ****
'****                       this please put my name on credits. ****
'****                                                           ****
'****   MABUHAY MY FELLOW FILIPINOS! ITO'Y GAWANG PINOY!!!      ****
'****                                                           ****
'*******************************************************************
'*******************************************************************
'*******************************************************************

'Opens a New Excel Workbook
Public Function NewExcel(ByVal objExcel As Excel.Application) As Excel.Application
    objExcel.Workbooks.Add
    objExcel.Visible = True
    Set NewExcel = objExcel
End Function

'Opens an Existing Excel Workbook
Public Function OpenExcel(ByVal objExcel As Excel.Application, ByVal sPath As String, ByVal sfilename As String) As Excel.Application
    objExcel.Workbooks.Open sPath & sfilename
    objExcel.Visible = True
    Set OpenExcel = objExcel
End Function

'Put a specified Value to a selected Cell it can be a string or a formula
Public Sub CellText(ByVal objExcel As Excel.Application, _
                    ByVal iRow As Integer, _
                    ByVal iCol As Integer, _
                    ByVal sValue As String)
    objExcel.Cells(iRow, iCol).Select
    objExcel.ActiveCell.FormulaR1C1 = sValue
End Sub

'Format the Font
Public Sub FormatFont(ByVal objExcel As Excel.Application, _
                      ByVal iRow As Integer, _
                      ByVal iCol As Integer, _
                      Optional ByVal iSize As Integer = 10, _
                      Optional ByVal bBold As Boolean = False, _
                      Optional ByVal bItalic As Boolean = False, _
                      Optional ByVal iUnderline As Integer = 0, _
                      Optional ByVal sFontName As String = "Arial", _
                      Optional ByVal clrColor As ColorConstants = vbBlack, _
                      Optional ByVal iAlignment As Integer = 0)
    'iAlignment = 0 then Align Left
    'iAlignment = 1 then Align Center
    'iAlignment = 2 then Align Right
    
    objExcel.Cells(iRow, iCol).Select
    With objExcel.ActiveCell
        .Font.Size = iSize
        .Font.Bold = bBold
        .Font.Italic = bItalic
        .Font.Name = sFontName
        .Font.Color = clrColor
        Select Case iUnderline
            Case 0:
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                .Borders(xlEdgeBottom).LineStyle = xlNone
                .Borders(xlEdgeRight).LineStyle = xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
                .Font.Underline = xlUnderlineStyleNone
            Case 1:
                .Font.Underline = xlUnderlineStyleSingle
            Case 2:
                    .Borders(xlDiagonalDown).LineStyle = xlNone
                    .Borders(xlDiagonalUp).LineStyle = xlNone
                    .Borders(xlEdgeLeft).LineStyle = xlNone
                    .Borders(xlEdgeTop).LineStyle = xlNone
                    With .Borders(xlEdgeBottom)
                        .LineStyle = xlDouble
                        .Weight = xlThick
                        .ColorIndex = xlAutomatic
                    End With
                    .Borders(xlEdgeRight).LineStyle = xlNone
        End Select
        Select Case iAlignment
            Case 0: .HorizontalAlignment = xlLeft
            Case 1: .HorizontalAlignment = xlCenter
            Case 2: .HorizontalAlignment = xlRight
        End Select
    End With
End Sub

'Set the Cell Format
Public Sub FormatCell(ByVal objExcel As Excel.Application, _
                      ByVal iRow As Integer, _
                      ByVal iCol As Integer, _
                      Optional ByVal bAutoFit As Boolean = False, _
                      Optional ByVal clrColor As ColorConstants = vbWhite)
                      
    objExcel.Cells(iRow, iCol).Select
        
    With objExcel.ActiveCell
        .Interior.Color = clrColor
        .Interior.Pattern = xlSolid
    End With
    
    If bAutoFit = True Then objExcel.Columns.AutoFit
End Sub

'Merge Cells
Public Sub MergeCell(ByVal objExcel As Excel.Application, _
                     ByVal iCons As Integer, _
                     ByVal iStart As Integer, _
                     ByVal iEnd As Integer, _
                     Optional iDirection As Integer = 0)
    'iDirection = 0 then Merge Horizontal
    'iDirection = 1 then Merge vertical
    
    If iDirection = 0 Then
        objExcel.Range(objExcel.Cells(iCons, iStart), objExcel.Cells(iCons, iEnd)).Select
    Else
        objExcel.Range(objExcel.Cells(iStart, iCons), objExcel.Cells(iEnd, iCons)).Select
    End If
    objExcel.Selection.Merge
End Sub

'Set table Grids
Public Sub PutGridLines(ByVal objExcel As Excel.Application, _
                        ByVal xStart As Integer, _
                        ByVal xEnd As Integer, _
                        ByVal yStart As Integer, _
                        ByVal yEnd As Integer, _
                        Optional ByVal iGridFormat As Integer = 0)
    'iGrid is 0 to 11 referrencing the different Grid Formats
    objExcel.Range(objExcel.Cells(xStart, yStart), objExcel.Cells(xEnd, yEnd)).Select
    With objExcel.Selection
        Select Case iGridFormat
            
            Case 0:
            
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                .Borders(xlEdgeBottom).LineStyle = xlNone
                .Borders(xlEdgeRight).LineStyle = xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
            
            Case 1:
            
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                .Borders(xlEdgeRight).LineStyle = xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
            
            Case 2:
                
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                With .Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                .Borders(xlEdgeTop).LineStyle = xlNone
                .Borders(xlEdgeBottom).LineStyle = xlNone
                .Borders(xlEdgeRight).LineStyle = xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
                
            Case 3:
            
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                .Borders(xlEdgeBottom).LineStyle = xlNone
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
                
            Case 4:
                
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                With Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlDouble
                    .Weight = xlThick
                    .ColorIndex = xlAutomatic
                End With
                .Borders(xlEdgeRight).LineStyle = xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
                
            Case 5:
                
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = xlAutomatic
                End With
                .Borders(xlEdgeRight).LineStyle = xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
                
            Case 6:
            
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlNone
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                .Borders(xlEdgeRight).LineStyle = xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
            
            Case 7:
            
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlNone
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlDouble
                    .Weight = xlThick
                    .ColorIndex = xlAutomatic
                End With
                .Borders(xlEdgeRight).LineStyle = xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
                
            Case 8:
            
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .Borders(xlEdgeLeft).LineStyle = xlNone
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = xlAutomatic
                End With
                .Borders(xlEdgeRight).LineStyle = xlNone
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
                
            Case 9:
                
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                With .Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With .Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With .Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
            
            Case 10:
                
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                With .Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone
                
            Case 11:
                
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                With .Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = xlAutomatic
                End With
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = xlAutomatic
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = xlAutomatic
                End With
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                    .ColorIndex = xlAutomatic
                End With
                .Borders(xlInsideVertical).LineStyle = xlNone
                .Borders(xlInsideHorizontal).LineStyle = xlNone

        End Select
    End With
End Sub

