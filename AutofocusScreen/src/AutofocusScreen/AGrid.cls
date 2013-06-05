VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''
' Class for creating a grid
'''
Option Explicit

''''Position on grid
Public NameGrid As String
Const maxZeros = 4
Private Type GridPoint
    Name As String
    X As Double
    Y As Double
    Z As Double
    Valid As Boolean
End Type
Private GridArray() As GridPoint
Public numRow As Long
Public numCol As Long
Public numRowSub As Long
Public numColSub As Long

''''Current Row, Col, RowSub, ColSub
Public iRow As Long
Public iCol As Long
Public iRowSub As Long
Public iColSub As Long

'''
' Check if indeces are in the allowable range
'''
Private Function correctIndeces(Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Boolean
    If Row > 0 And Row <= numRow And Col > 0 And Col <= numCol And RowSub > 0 And RowSub <= numRowSub And ColSub > 0 And ColSub <= numColSub Then
        correctIndeces = True
    End If
End Function

Public Function thisSuffix() As String
    'convert numbers into a string
    Dim iWell As Long
    Dim iPosition As Long
    Dim Name As String
    Dim nrZero As Integer
    Name = ""
    iWell = (iRow - 1) * numCol + iCol
    iPosition = (iRowSub - 1) * numColSub + iColSub
    If iWell >= 0 Then
        nrZero = maxZeros - Len(CStr(iWell))
        Name = Name + "W" + ZeroString(nrZero) + CStr(iWell)
    End If
    If iPosition >= 0 Then
        nrZero = maxZeros - Len(CStr(iPosition))
        Name = Name + "_P" + ZeroString(nrZero) + CStr(iPosition)
    End If
    thisSuffix = Name
End Function


Public Function thisSuffixWell() As String
    'convert numbers into a string
    Dim iWell As Long
    Dim iPosition As Long
    Dim Name As String
    Dim nrZero As Integer
    Name = ""
    iWell = (iRow - 1) * numCol + iCol
    If iWell >= 0 Then
        nrZero = maxZeros - Len(CStr(iWell))
        Name = Name + "W" + ZeroString(nrZero) + CStr(iWell)
    End If
    thisSuffixWell = Name
End Function

Public Function suffix(Row As Long, Col As Long, RowSub As Long, ColSub As Long) As String
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        'convert numbers into a string
        Dim iWell As Long
        Dim iPosition As Long
        Dim Name As String
        Dim nrZero As Integer
        Name = ""
        iWell = (Row - 1) * numCol + Col
        iPosition = (RowSub - 1) * numColSub + ColSub
        If iWell >= 0 Then
            nrZero = maxZeros - Len(CStr(iWell))
            Name = Name + "W" + ZeroString(nrZero) + CStr(iWell)
        End If
        If iPosition >= 0 Then
            nrZero = maxZeros - Len(CStr(iPosition))
            Name = Name + "_P" + ZeroString(nrZero) + CStr(iPosition)
        End If
        suffix = Name
    End If
End Function



Public Function getThisName() As String
    getThisName = GridArray(iRow, iCol, iRowSub, iColSub).Name
End Function


Public Function getName(Row As Long, Col As Long, RowSub As Long, ColSub As Long) As String
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        getName = GridArray(Row, Col, RowSub, ColSub).Name
    End If
End Function


Public Sub setThisName(Value As String)
    GridArray(iRow, iCol, iRowSub, iColSub).Name = Value
End Sub


Public Sub setName(Value As String, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        GridArray(Row, Col, RowSub, ColSub).Name = Value
        If Value <> "" Then
            If Right(Value, 1) <> "_" Then
                GridArray(Row, Col, RowSub, ColSub).Name = Value & "_"
            End If
        End If
    End If
End Sub

'''
'   Initialize the dimensions of the array
'''
Public Sub initialize(Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    numRow = Row
    numCol = Col
    numRowSub = RowSub
    numColSub = ColSub
    iRow = 1
    iCol = 1
    iRowSub = 1
    iColSub = 1
    ReDim GridArray(1 To numRow, 1 To numCol, 1 To numRowSub, 1 To numColSub)
End Sub


'''
'   Initialize the dimensions of the array
'''
Public Sub redimPreserve(Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    numRow = Row
    numCol = Col
    numRowSub = RowSub
    numColSub = ColSub
    iRow = 1
    iCol = 1
    iRowSub = 1
    iColSub = 1
    ReDim Preserve GridArray(1 To numRow, 1 To numCol, 1 To numRowSub, 1 To numColSub)
End Sub


Public Function getIndeces(Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    Row = iRow
    Col = iCol
    RowSub = iRowSub
    ColSub = iColSub
End Function



Public Function setIndeces(Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        iRow = Row
        iCol = Col
        iRowSub = RowSub
        iColSub = ColSub
    End If
End Function

'''
'   get/set current grid position
'''
Public Sub getThisPt(X As Double, Y As Double, Z As Double, Valid As Boolean)
    X = GridArray(iRow, iCol, iRowSub, iColSub).X
    Y = GridArray(iRow, iCol, iRowSub, iColSub).Y
    Z = GridArray(iRow, iCol, iRowSub, iColSub).Z
    Valid = GridArray(iRow, iCol, iRowSub, iColSub).Valid
End Sub

Public Function getThisX() As Double
    getThisX = GridArray(iRow, iCol, iRowSub, iColSub).X
End Function


Public Function getThisY() As Double
    getThisY = GridArray(iRow, iCol, iRowSub, iColSub).Y
End Function


Public Function getThisZ() As Double
    getThisZ = GridArray(iRow, iCol, iRowSub, iColSub).Z
End Function


Public Function getThisValid() As Boolean
    getThisValid = GridArray(iRow, iCol, iRowSub, iColSub).Valid
End Function


Public Sub setThisPt(X As Double, Y As Double, Z As Double, Valid As Boolean)
    GridArray(iRow, iCol, iRowSub, iColSub).X = X
    GridArray(iRow, iCol, iRowSub, iColSub).Y = Y
    GridArray(iRow, iCol, iRowSub, iColSub).Z = Z
    GridArray(iRow, iCol, iRowSub, iColSub).Valid = Valid
End Sub


Public Sub setThisX(Value As Double)
    GridArray(iRow, iCol, iRowSub, iColSub).X = Value
End Sub

Public Sub setThisY(Value As Double)
    GridArray(iRow, iCol, iRowSub, iColSub).Y = Value
End Sub

Public Sub setThisZ(Value As Double)
    GridArray(iRow, iCol, iRowSub, iColSub).Z = Value
End Sub


Public Sub setThisValid(Value As Boolean)
    GridArray(iRow, iCol, iRowSub, iColSub).Valid = Value
End Sub




'''
'   get/set position defined by Row, Colm RowSub and ColSub
''''
Public Sub getPt(X As Double, Y As Double, Z As Double, Valid As Boolean, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        X = GridArray(Row, Col, RowSub, ColSub).X
        Y = GridArray(Row, Col, RowSub, ColSub).Y
        Z = GridArray(Row, Col, RowSub, ColSub).Z
        Valid = GridArray(Row, Col, RowSub, ColSub).Valid
    Else
        MsgBox "Grid getPt: was not able to set grid pt. Gridposition does not correspond to grid size"
    End If
End Sub


Public Function getX(Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Double
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        getX = GridArray(Row, Col, RowSub, ColSub).X
    Else
        MsgBox "Grid getPt: was not able to set grid pt. Gridposition does not correspond to grid size"
    End If
End Function

Public Function getY(Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Double
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        getY = GridArray(Row, Col, RowSub, ColSub).Y
    Else
        MsgBox "Grid getPt: was not able to set grid pt. Gridposition does not correspond to grid size"
    End If
End Function

Public Function getZ(Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Double
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        getZ = GridArray(Row, Col, RowSub, ColSub).Z
    Else
        MsgBox "Grid getPt: was not able to set grid pt. Gridposition does not correspond to grid size"
    End If
End Function

Public Function getValid(Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Boolean
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        getValid = GridArray(Row, Col, RowSub, ColSub).Valid
    Else
        MsgBox "Grid getPt: was not able to set grid pt. Gridposition does not correspond to grid size"
    End If
End Function

Public Sub setPt(X As Double, Y As Double, Z As Double, Valid As Boolean, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        GridArray(Row, Col, RowSub, ColSub).X = X
        GridArray(Row, Col, RowSub, ColSub).Y = Y
        GridArray(Row, Col, RowSub, ColSub).Z = Z
        GridArray(Row, Col, RowSub, ColSub).Valid = Valid
    Else
        MsgBox "Grid setPt: was not able to set grid pt. Gridposition does not correspond to grid size"
    End If
End Sub

Public Function setX(X As Double, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        GridArray(Row, Col, RowSub, ColSub).X = X
    Else
        MsgBox "Grid getPt: was not able to set grid pt. Gridposition does not correspond to grid size"
    End If
End Function

Public Function setY(Y As Double, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Double
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        GridArray(Row, Col, RowSub, ColSub).Y = Y
    Else
        MsgBox "Grid getPt: was not able to set grid pt. Gridposition does not correspond to grid size"
    End If
End Function

Public Function setZ(Z As Double, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Double
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        GridArray(Row, Col, RowSub, ColSub).Z = Z
    Else
        MsgBox "Grid getPt: was not able to set grid pt. Gridposition does not correspond to grid size"
    End If
End Function


Public Function setValid(Valid As Boolean, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Boolean
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        GridArray(Row, Col, RowSub, ColSub).Valid = Valid
    Else
        MsgBox "Grid getPt: was not able to set grid pt. Gridposition does not correspond to grid size"
    End If
End Function



'''
' Go to next gridPt according to meandering. True if there is a next grid pt. False if Grid has reached an end
'''
Public Function nextGridPt() As Boolean
    'update GridSub
    If iRowSub Mod 2 <> 0 Then
        iColSub = iColSub + 1
    Else
        iColSub = iColSub - 1
    End If
    
    If iColSub > numColSub Then
        iColSub = 1
        iRowSub = iRowSub + 1
    End If
    
    If iColSub < 1 Then
        iColSub = numColSub
        iRowSub = iRowSub + 1
    End If
        
    'update mainGrid
    If iRowSub > numRowSub Then
        iRowSub = 1
        If iRow Mod 2 <> 0 Then
            iCol = iCol + 1
        Else
            iCol = iCol - 1
        End If
        If iCol > numCol Then
            iCol = numCol
            iRow = iRow + 1
        End If
        If iCol < 1 Then
            iCol = 1
            iRow = iRow + 1
        End If
    End If
    
    ''end of grid has been reached
    If iRow > numRow Then
        nextGridPt = False
    Else
        nextGridPt = True
    End If
                    
End Function


'''''
'   makeGridFromOnePt( posGridX() As Double, posGridY() As Double, posGridXY_valid() )
'   Create a Grid and a Subgrid use using XStart, YStart, and ZStart as starting coordinates
'       dX, dY:         Gridspacing main Grid. The main grid is the center of the subgrid
'       dXsub, dYsub:   Gridspacing subGrid
'       refRow:         reference row where main grid starts
'       refCol:         reference column where main grid starts
'''''
Public Sub makeGridFromOnePt(XStart As Double, YStart As Double, ZStart As Double, dX As Double, dY As Double, _
dXsub As Double, dYsub As Double, Optional refCol As Integer = 1, Optional refRow As Integer = 1)
        ' A row correspond to Y movement and Column to X shift
        ' entries are posGridX(row, column)!! this what is
        'counters
        'Make main grid
        Dim iRow As Long
        Dim iCol As Long
        Dim iRowSub As Long
        Dim iColSub As Long
        'Make grid and subgrid
        For iRow = 1 To numRow
            For iCol = 1 To numCol
                For iRowSub = 1 To numRowSub
                    For iColSub = 1 To numColSub
                        GridArray(iRow, iCol, iRowSub, iColSub).X = Round(XStart + (1 - refCol) * dX + (iCol - 1) * dX + (iColSub - 1 - (numColSub - 1) / 2) * dXsub, PrecXY)
                        GridArray(iRow, iCol, iRowSub, iColSub).Y = Round(YStart + (1 - refRow) * dY + (iRow - 1) * dY + (iRowSub - 1 - (numRowSub - 1) / 2) * dYsub, PrecXY)
                        GridArray(iRow, iCol, iRowSub, iColSub).Z = Round(ZStart, PrecZ)
                    Next iColSub
                Next iRowSub
            Next iCol
        Next iRow
            
        ' All points are true as default
        For iRow = 1 To numRow
            For iCol = 1 To numCol
                For iRowSub = 1 To numRowSub
                    For iColSub = 1 To numColSub
                        GridArray(iRow, iCol, iRowSub, iColSub).Valid = True
                    Next iColSub
                Next iRowSub
            Next iCol
        Next iRow
End Sub


'''''
'   makeGridFromManyPts(X() As Double, Y() As Double, Z() As Double, dXsub As Double, dYsub As Double  )
'   Create a Grid and a Subgrid use using X, Y and Z as coordinates for the main grid. A row corresponds to
'   Y movement and Column to X shift\
'       X, Y, Z:        Array containing coordinates of main grid.
'       dXsub, dYsub:   Gridspacing subGrid
'''''
Public Sub makeGridFromManyPts(X() As Double, Y() As Double, Z() As Double, dXsub As Double, dYsub As Double)
        Dim iRow As Long
        Dim iRowSub As Long
        Dim iColSub As Long
        If UBound(X) <> UBound(Y) Or UBound(X) <> UBound(Z) Then
            MsgBox "makeGridFromManyPts: Number of points should be the same for every coordinate"
            Exit Sub
        End If
        'Make grid and subgri
        For iCol = 1 To numCol
            For iRowSub = 1 To numColSub
                For iColSub = 1 To numRowSub
                    GridArray(1, iCol, iRowSub, iColSub).X = Round(X(iCol - 1) + (iColSub - 1 - (numColSub - 1) / 2) * dXsub, PrecXY)
                    GridArray(1, iCol, iRowSub, iColSub).Y = Round(Y(iCol - 1) + (iRowSub - 1 - (numRowSub - 1) / 2) * dYsub, PrecXY)
                    GridArray(1, iCol, iRowSub, iColSub).Z = Round(Z(iCol - 1), PrecZ)
                Next iColSub
            Next iRowSub
        Next iCol
        
        ' All points are true as default
        For iCol = 1 To numCol
            For iRowSub = 1 To numRowSub
                For iColSub = 1 To numColSub
                    GridArray(1, iCol, iRowSub, iColSub).Valid = True
                Next iColSub
            Next iRowSub
        Next iCol

End Sub

''''
' isPositionGridFile
' Check that entries of first line correspond to 4 specified values
''''
Public Function isPositionGridFile(ByVal sFile As String, ByVal Row As Integer, ByVal Col As Integer, ByVal RowSub As Integer, ByVal ColSub As Integer) As Boolean
    If FileExist(sFile) Then
        Close
        On Error GoTo ErrorPosFile:
        Dim iFileNum As Integer
        Dim Fields As String
        Dim FieldEntries() As String
        iFileNum = FreeFile()
        Open sFile For Input As iFileNum
        Line Input #iFileNum, Fields
        While Left(Fields, 1) = "%"
            Line Input #iFileNum, Fields
        Wend
        FieldEntries = Split(Fields, " ")
        If FieldEntries(0) = Row And FieldEntries(1) = Col And FieldEntries(2) = RowSub And FieldEntries(3) = ColSub Then
            isPositionGridFile = True
        Else
           MsgBox ("CheckPosfile: Position file " & sFile & " needs to have same number of rows/col subrows/subcol as grid in form." & vbCrLf & _
           "File: Row= " & FieldEntries(0) & ", Col= " & FieldEntries(1) & ", RowSub= " & FieldEntries(2) _
           & ", ColSub= " & FieldEntries(3) & vbCrLf & _
            "Form: Row= " & Row & ", Col= " & Col & ", RowSub= " & RowSub _
           & ", ColSub= " & ColSub)
           isPositionGridFile = False
        End If
        Close #iFileNum
    Else
       MsgBox ("CheckPosfile: Was not able to load position file " & sFile & ". File does not exist.")
    End If
    Exit Function
ErrorPosFile:
    If Err.number = 70 Then
        MsgBox ("CheckPosfile: Was not able to load position file " & sFile & ". File is open by another program.")
    End If
End Function

'''''''
' loadPositionGridFile(ByVal sFile As String)
'       [sFile] In  - Output file name
'   Function loads a file and write entries in Double arrays. First entry of file is structure of grid
'   Format is numRow numCol numRowSub numColSub
'   X Y Z X Y Z X Y Z etc
''''''''
Public Function loadPositionGridFile(ByVal sFile As String) As Boolean
    Dim iRow As Integer
    Dim iCol As Integer
    Dim iRowSub As Integer
    Dim iColSub As Integer
    Close
    On Error GoTo ErrorPosFile:
    Dim iFileNum As Integer
    Dim Fields As String
    Dim FieldEntries() As String
    iFileNum = FreeFile()
    Open sFile For Input As iFileNum
    Line Input #iFileNum, Fields
    ''jump comments
    While Left(Fields, 1) = "%"
        Line Input #iFileNum, Fields
    Wend
    FieldEntries = Split(Fields, " ")
    initialize CLng(FieldEntries(0)), CLng(FieldEntries(1)), CLng(FieldEntries(2)), CLng(FieldEntries(3))
    
    For iRow = 1 To numRow
      For iCol = 1 To numCol
          Line Input #iFileNum, Fields
          While Left(Fields, 1) = "%"
            Line Input #iFileNum, Fields
          Wend
          FieldEntries = Split(Fields, " ")
          For iRowSub = 1 To numRowSub
              For iColSub = 1 To numColSub
                  GridArray(iRow, iCol, iRowSub, iColSub).X = CDbl(FieldEntries((iColSub - 1) * 3 + (iRowSub - 1) * UBound(posGridX, 3) * 3))
                  GridArray(iRow, iCol, iRowSub, iColSub).Y = CDbl(FieldEntries((iColSub - 1) * 3 + (iRowSub - 1) * UBound(posGridX, 3) * 3 + 1))
                  GridArray(iRow, iCol, iRowSub, iColSub).Z = CDbl(FieldEntries((iColSub - 1) * 3 + (iRowSub - 1) * UBound(posGridX, 3) * 3 + 2))
              Next iColSub
          Next iRowSub
      Next iCol
    Next iRow
    Close #iFileNum
    loadPositionGridFile = True
    Exit Function
ErrorPosFile:
    If Err.number = 70 Then
        MsgBox ("LoadPosFile: Was not able to load position file " & sFile & ". File is open by another program.")
    Else
        MsgBox ("LoadPosFile: Was not able to load position file " & sFile)
    End If
    MsgBox ("Was not able to load position file " & sFile)
End Function
    
 
'''''''
'   writePositionGridFile (ByVal sFile As String)
'       [sFile] In  - Output file name
'   Write out position of grid. The first uncommented line is the structure of the grid. Comments are preceeded by %
'   Format is
'   numRow numCol numRowSub numColSub
'   X Y Z X Y Z X Y Z etc
''''''''
Public Function writePositionGridFile(ByVal sFile As String) As Boolean
    Dim iRow As Integer
    Dim iCol As Integer
    Dim iRowSub As Integer
    Dim iColSub As Integer
    Dim Line As String
    Dim LineComm As String
    Close
    On Error GoTo ErrorPosFile:
    Dim iFileNum As Integer
    Dim Fields As String
    Dim FieldEntries() As String
    iFileNum = FreeFile()
    Open sFile For Output As iFileNum
    Print #iFileNum, "%nrRows nrColumns nrRowSubs nrColSubumns"
    Print #iFileNum, numRow & " " & numCol & " " & numRowSub & " " & numColSub
    For iRow = 1 To numRow
      For iCol = 1 To numCol
          LineComm = "%Row: " & iRow & ", Col: " & iCol & " "
          Print #iFileNum, LineComm
          Line = ""
          LineComm = "%Rowsub Colsub: "
          For iRowSub = 1 To numRowSub
              For iColSub = 1 To numColSub
                 LineComm = LineComm & iRowSub & " " & iColSub & ", "
                 Line = Line & GridArray(iRow, iCol, iRowSub, iColSub).X & " " & GridArray(iRow, iCol, iRowSub, iColSub).Y & " " _
                  & GridArray(iRow, iCol, iRowSub, iColSub).Z & " "
              Next iColSub
          Next iRowSub
        Print #iFileNum, LineComm
        Print #iFileNum, Line
      Next iCol
    Next iRow
    Close #iFileNum
    writePositionGridFile = True
    Exit Function
ErrorPosFile:
    If Err.number = 70 Then
        MsgBox ("WritePosFile: Was not able to load position file " & sFile & ". File is open by another program")
    Else
        MsgBox ("WritePosFile: Was not able to load position file " & sFile)
    End If
    Close #iFileNum
End Function

''''
'   Check if file has a validGrid default format
''''
Public Function isValidGridFile(ByVal sFile As String, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As String
    Dim CellBase As String
    Dim Default As String
    Dim last_entry  As String
    Dim Active As Boolean
    Dim GoodMatch As Boolean
    Dim RegEx As VBScript_RegExp_55.RegExp
    Set RegEx = CreateObject("vbscript.regexp")
    Dim Match As MatchCollection
    Dim format As String
    'Well--Row--Col--(Row,Col)--TypeofWell
    CellBase = "(\d+)--(\d+)--(\d+)--(\S+)--(\S+)"
    
    If FileExist(sFile) Then
        Close
        Dim iFileNum As Integer
        Dim Fields As String
        Dim FieldEntries() As String
        iFileNum = FreeFile()
        Open sFile For Input As iFileNum
        
        'read first line
        Line Input #iFileNum, Fields
        While Left(Fields, 1) = "%"
           Line Input #iFileNum, Fields
        Wend
        RegEx.Pattern = CellBase
        If RegEx.Test(Fields) Then
            format = "cellbase"
        Else
            'standard format
            FieldEntries = Split(Fields, " ")
            If UBound(FieldEntries) = 3 Then
                If FieldEntries(0) = Row And FieldEntries(1) = Col And FieldEntries(2) = RowSub And FieldEntries(3) = ColSub Then
                    isValidGridFile = "default"
                Else
                    MsgBox ("isValidGridFile: Valid Grid file " & sFile & " needs to have same number of rows/col subrows/subcol as grid in form." & vbCrLf & _
                    "File: Row= " & FieldEntries(0) & ", Col= " & FieldEntries(1) & ", RowSub= " & FieldEntries(2) _
                    & ", ColSub= " & FieldEntries(3) & vbCrLf & _
                     "Form: Row= " & Row & ", Col= " & Col & ", RowSub= " & RowSub _
                    & ", ColSub= " & ColSub)
                    isValidGridFile = "none"
                End If
            Else
                MsgBox sFile & " has not the correct format to set a validGrid and names! " & vbCrLf & "The format should be " & vbCrLf & "(In)active(0 or 1) Row(>=1)--Col(>=1) e.g." & vbCrLf & "0 1--1" & vbCrLf & "1 1--2" & vbCrLf _
                & "or CellBase format Well--Row--Col--(Row,Col)--Identifier Identifier = none=> InactiveWell e.g." & vbCrLf & "1--2--1--(1,1)--empty" & vbCrLf _
                & "1--1--2--(1,2)--MyCoolGene or empty" & vbCrLf & "or a list of 1 1 1 1 per subposition"
                isValidGridFile = "none"
            End If
        End If
        Close
    Else
        MsgBox "isValidGridFile: Could not load " & sFile
        isValidGridFile = "none"
    End If
End Function


'''''''
'
''''''''
Public Function loadValidGridFile(ByVal sFile As String, format As String) As Boolean
    
    Close
    Dim iRow As Long
    Dim iCol As Long
    Dim iRowSub As Long
    Dim iColSub As Long
    Dim iFileNum As Integer
    Dim Fields As String
    Dim FieldEntries() As String
    Dim GoodMatch As Boolean
    Dim CellBase As String
    Dim RegEx As VBScript_RegExp_55.RegExp
    Set RegEx = CreateObject("vbscript.regexp")
    Dim Match As MatchCollection
    CellBase = "(\d+)--(\d+)--(\d+)--(\S+)--(\S+)"
    Dim Active As Boolean
    If format <> "none" Or format <> "" Then
        iFileNum = FreeFile()
        Open sFile For Input As iFileNum
        If format = "default" Then
            Line Input #iFileNum, Fields
            While Left(Fields, 1) = "%"
                Line Input #iFileNum, Fields
            Wend
            FieldEntries = Split(Fields, " ")
            For iRow = 1 To numRow
                For iCol = 1 To numCol
                  Line Input #iFileNum, Fields
                  While Left(Fields, 1) = "%"
                    Line Input #iFileNum, Fields
                  Wend
                  FieldEntries = Split(Fields, " ")
                  For iRowSub = 1 To numRowSub
                      For iColSub = 1 To numColSub
                          GridArray(iRow, iCol, iRowSub, iColSub).Valid = CBool(FieldEntries((iColSub - 1) + (iRowSub - 1) * numRowSub))
                      Next iColSub
                  Next iRowSub
                Next iCol
            Next iRow
        ElseIf format = "cellbase" Then
            Do While Not EOF(iFileNum)
                Line Input #iFileNum, Fields
                GoodMatch = False
                RegEx.Pattern = CellBase
                If RegEx.Test(Fields) Then
                    Set Match = RegEx.Execute(Fields)
                    Active = (Match.Item(0).SubMatches.Item(4) <> "none")
                    GoodMatch = True
                End If
                'if this gridposition exists then update activity
                If GoodMatch And CInt(Match.Item(0).SubMatches.Item(1)) <= UBound(posGridXY_Valid, 1) And CInt(Match.Item(0).SubMatches.Item(1)) >= LBound(posGridXY_Valid, 1) _
                And CInt(Match.Item(0).SubMatches.Item(2)) <= UBound(posGridXY_Valid, 2) And CInt(Match.Item(0).SubMatches.Item(2)) >= LBound(posGridXY_Valid, 2) Then
                    For iRowSub = 1 To UBound(posGridXY_Valid, 3)
                        For iColSub = 1 To UBound(posGridXY_Valid, 4)
                            GridArray(CInt(Match.Item(0).SubMatches.Item(1)), CInt(Match.Item(0).SubMatches.Item(2)), iRowSub, iColSub).Valid = Active
                            setName Match.Item(0).SubMatches.Item(3) & "_" & Match.Item(0).SubMatches.Item(4), CInt(Match.Item(0).SubMatches.Item(1)), CInt(Match.Item(0).SubMatches.Item(2)), iRowSub, iColSub
                        Next iColSub
                    Next iRowSub
                End If
            Loop
        End If
    Else
        MsgBox ("loadValidGridFile: was not able to understand format " & format & " of " & sFile)
    End If
End Function

    
'''''''
'
'''''''
Public Function writeValidGridFile(ByVal sFile As String) As Boolean
    Dim iRow As Integer
    Dim iCol As Integer
    Dim iRowSub As Integer
    Dim iColSub As Integer
    Dim Line As String
    Dim LineComm As String
    Close
    On Error GoTo ErrorPosFile:
    Dim iFileNum As Integer
    Dim Fields As String
    Dim FieldEntries() As String
    iFileNum = FreeFile()
    Open sFile For Output As iFileNum
    Print #iFileNum, "%nrRows nrColumns nrRowSubs nrColSubumns"
    Print #iFileNum, numRow & " " & numCol & " " & numRowSub & " " & numColSub
    For iRow = 1 To numRow
      For iCol = 1 To numCol
          LineComm = "%Row: " & iRow & ", Col: " & iCol & " "
          Print #iFileNum, LineComm
          Line = ""
          LineComm = "%Rowsub Colsub: "
          For iRowSub = 1 To numRowSub
              For iColSub = 1 To numColSub
                 LineComm = LineComm & iRowSub & " " & iColSub & ", "
                 Line = Line & -GridArray(iRow, iCol, iRowSub, iColSub).Valid * 1 & " "
              Next iColSub
          Next iRowSub
        Print #iFileNum, LineComm
        Print #iFileNum, Line
      Next iCol
    Next iRow
    Close #iFileNum
    writeValidGridFile = True
    Exit Function
ErrorPosFile:
    If Err.number = 70 Then
        MsgBox ("WriteValidPosFile: Was not able to write valid position file " & sFile & ". File is open by another program")
    Else
        MsgBox ("WriteValidPosFile: Was not able to write valid position file " & sFile)
    End If
    Close #iFileNum
End Function



