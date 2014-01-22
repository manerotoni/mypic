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
' Class for creating and administrating a grid
' The positions and subpositions are numbered starting from 1
' A grid defines positions travelled with the stage for imaging
'''
Option Explicit

'''' Name if grid
Public NameGrid As String
'''' Max zeros padded in front of Wxxxx
Const maxZeros = 4

'''' A gridpoint is the unit of a grid. It will have a name specified by its position in the array
Private Type GridPoint
    Name As String    'this is used as prefix for the name
    X As Double
    Y As Double
    Z As Double
    Valid As Boolean
End Type

'''' An array of gridpoints
Private GridArray() As GridPoint
'''' Define size of Array
Public numRow As Long
Public numCol As Long
Public numRowSub As Long
Public numColSub As Long

''''Current Row, Col, RowSub, ColSub when 0 it is empty array. Proper array starts from 1!
Public iRow As Long
Public iCol As Long
Public iRowSub As Long
Public iColSub As Long

'''
'   Initialize the dimensions of the array
'''
Public Sub initialize(Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If Row = 0 Or Col = 0 Or RowSub = 0 Or ColSub = 0 Then
        initializeToZero
    Else
  On Error GoTo ErrorHandle:
        Me.numRow = Row
        Me.numCol = Col
        Me.numRowSub = RowSub
        Me.numColSub = ColSub
        Me.iRow = 1
        Me.iCol = 1
        Me.iRowSub = 1
        Me.iColSub = 1
        ReDim GridArray(1 To Me.numRow, 1 To Me.numCol, 1 To Me.numRowSub, 1 To Me.numColSub)
    End If
    Exit Sub
ErrorHandle:
    MsgBox "Error, " + Err.Source + " in AGrid.initialize " + Err.Description
End Sub


'''
'   Reset array to empty array
'''
Public Sub initializeToZero()
    Me.numRow = 0
    Me.numCol = 0
    Me.numRowSub = 0
    Me.numColSub = 0
    Me.iRow = 0
    Me.iCol = 0
    Me.iRowSub = 0
    Me.iColSub = 0
    ReDim GridArray(0, 0, 0, 0)
End Sub

'''
'   Redimension the array and preserve size
'''
Public Sub updateGridSizePreserve(Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    Dim iRow As Long
    Dim iCol As Long
    Dim iRowSub As Long
    Dim iColSub As Long
    Dim TmpGridArray() As GridPoint
    TmpGridArray = GridArray
    Me.iRow = 1
    Me.iCol = 1
    Me.iRowSub = 1
    Me.iColSub = 1
    
    If Me.numRow = Row And Me.numCol = Col And Me.numRowSub = RowSub Then
        ReDim Preserve GridArray(1 To Me.numRow, 1 To Me.numCol, 1 To Me.numRowSub, 1 To ColSub)
    Else
        'can only update last dimension
        ReDim GridArray(1 To Row, 1 To Col, 1 To RowSub, 1 To ColSub)
        
        For iRow = 1 To Me.numRow
            For iCol = 1 To Me.numCol
                For iRowSub = 1 To numRowSub
                    For iColSub = 1 To Me.numColSub
                        GridArray(iRow, iCol, iRowSub, iColSub) = TmpGridArray(iRow, iCol, iRowSub, iColSub)
                    Next iColSub
                Next iRowSub
            Next iCol
        Next iRow
    End If
    
    Me.numRow = Row
    Me.numCol = Col
    Me.numRowSub = RowSub
    Me.numColSub = ColSub

    ReDim Preserve GridArray(1 To Me.numRow, 1 To Me.numCol, 1 To Me.numRowSub, 1 To Me.numColSub)
End Sub

Public Function isGridEmpty() As Boolean
    If LBound(GridArray) = 0 And UBound(GridArray) = 0 Then
        isGridEmpty = True
    End If
End Function

'''
' Check if indeces are in the allowable range
'''
Private Function correctIndeces(Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Boolean
    If Row > 0 And Row <= Me.numRow And Col > 0 And Col <= Me.numCol And RowSub > 0 And RowSub <= Me.numRowSub And ColSub > 0 And ColSub <= Me.numColSub Then
        correctIndeces = True
    Else
        MsgBox "Grid getPt: was not able to access grid pt. Gridposition does not correspond to grid size"
    End If
End Function

'''
' suffix is WxxxxPxxxx created at current position of grid
'''
Public Function thisSuffix() As String
    'convert numbers into a string
    Dim iWell As Long
    Dim iPosition As Long
    Dim Name As String
    Dim nrZero As Integer
    Name = ""
    iWell = (Me.iRow - 1) * numCol + Me.iCol
    iPosition = (Me.iRowSub - 1) * numColSub + Me.iColSub
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

'''
'   Suffix without subposition indication
'''
Public Function thisSuffixWell() As String
    'convert numbers into a string
    Dim iWell As Long
    Dim iPosition As Long
    Dim Name As String
    Dim nrZero As Integer
    Name = ""
    iWell = (Me.iRow - 1) * numCol + Me.iCol
    If iWell >= 0 Then
        nrZero = maxZeros - Len(CStr(iWell))
        Name = Name + "W" + ZeroString(nrZero) + CStr(iWell)
    End If
    thisSuffixWell = Name
End Function

'''
' Suffix given a Row, Col etc.
'''
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
    getThisName = GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).Name
End Function


Public Function getName(Row As Long, Col As Long, RowSub As Long, ColSub As Long) As String
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        getName = GridArray(Row, Col, RowSub, ColSub).Name
    End If
End Function


Public Sub setThisName(Value As String)
    GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).Name = Value
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





Public Function getIndeces(Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    Row = Me.iRow
    Col = Me.iCol
    RowSub = Me.iRowSub
    ColSub = Me.iColSub
End Function



Public Function setIndeces(Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        Me.iRow = Row
        Me.iCol = Col
        Me.iRowSub = RowSub
        Me.iColSub = ColSub
    End If
End Function

'''
'   get/set current grid position
'''
Public Sub getThisPt(pos As Vector, Valid As Boolean)
    pos.X = GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).X
    pos.Y = GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).Y
    pos.Z = GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).Z
    Valid = GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).Valid
End Sub

Public Function getThisX() As Double
    getThisX = GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).X
End Function


Public Function getThisY() As Double
    getThisY = GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).Y
End Function


Public Function getThisZ() As Double
    getThisZ = GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).Z
End Function


Public Function getThisValid() As Boolean
    getThisValid = GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).Valid
End Function


Public Sub setThisPt(pos As Vector, Valid As Boolean)
    GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).X = pos.X
    GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).Y = pos.Y
    GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).Z = pos.Z
    GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).Valid = Valid
End Sub


Public Sub setThisX(Value As Double)
    GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).X = Value
End Sub

Public Sub setThisY(Value As Double)
    GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).Y = Value
End Sub

Public Sub setThisZ(Value As Double)
    GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).Z = Value
End Sub


Public Sub setThisValid(Value As Boolean)
    GridArray(Me.iRow, Me.iCol, Me.iRowSub, Me.iColSub).Valid = Value
End Sub




'''
'   get/set position defined by Row, Colm RowSub and ColSub
''''
Public Sub getPt(pos As Vector, Valid As Boolean, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        pos.X = GridArray(Row, Col, RowSub, ColSub).X
        pos.Y = GridArray(Row, Col, RowSub, ColSub).Y
        pos.Z = GridArray(Row, Col, RowSub, ColSub).Z
        Valid = GridArray(Row, Col, RowSub, ColSub).Valid
    End If
End Sub

'''
' get nr of points in grid
'''
Public Function getNrPts() As Long
    getNrPts = Me.numCol * Me.numRow * Me.numRowSub * Me.numColSub
End Function


'''
'   get nr of valid points in grid
'''
Public Function getNrValidPts() As Long
    Dim iRow As Long
    Dim iCol As Long
    Dim iRowSub As Long
    Dim iColSub As Long
    Dim nrValid As Long
    nrValid = 0
    ' All points are true as default
    For iRow = 1 To Me.numRow
        For iCol = 1 To Me.numCol
            For iRowSub = 1 To numRowSub
                For iColSub = 1 To Me.numColSub
                    If GridArray(iRow, iCol, iRowSub, iColSub).Valid Then
                        nrValid = nrValid + 1
                    End If
                Next iColSub
            Next iRowSub
        Next iCol
    Next iRow
    getNrValidPts = nrValid
End Function



Public Function getX(Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Double
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        getX = GridArray(Row, Col, RowSub, ColSub).X
    End If
End Function

Public Function getY(Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Double
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        getY = GridArray(Row, Col, RowSub, ColSub).Y
    End If
End Function

Public Function getZ(Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Double
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        getZ = GridArray(Row, Col, RowSub, ColSub).Z
    End If
End Function

Public Function getValid(Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Boolean
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        getValid = GridArray(Row, Col, RowSub, ColSub).Valid
    End If
End Function

Public Sub setPt(pos As Vector, Valid As Boolean, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        GridArray(Row, Col, RowSub, ColSub).X = pos.X
        GridArray(Row, Col, RowSub, ColSub).Y = pos.Y
        GridArray(Row, Col, RowSub, ColSub).Z = pos.Z
        GridArray(Row, Col, RowSub, ColSub).Valid = Valid
    End If
End Sub

Public Sub setX(X As Double, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        GridArray(Row, Col, RowSub, ColSub).X = X
    End If
End Sub

Public Sub setY(Y As Double, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        GridArray(Row, Col, RowSub, ColSub).Y = Y
    End If
End Sub

Public Sub setZ(Z As Double, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        GridArray(Row, Col, RowSub, ColSub).Z = Z
    End If
End Sub


Public Sub setValid(Valid As Boolean, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If correctIndeces(Row, Col, RowSub, ColSub) Then
        GridArray(Row, Col, RowSub, ColSub).Valid = Valid
    End If
End Sub

Public Sub setAllValid(Valid As Boolean)
    ' All points are true as default
    For iRow = 1 To Me.numRow
        For iCol = 1 To Me.numCol
            For iRowSub = 1 To numRowSub
                For iColSub = 1 To Me.numColSub
                    GridArray(iRow, iCol, iRowSub, iColSub).Valid = Valid
                Next iColSub
            Next iRowSub
        Next iCol
    Next iRow
End Sub



'''
' Go to next gridPt according to meandering. True if there is a next grid pt. False if Grid has reached an end
' WellFirst changes order of meandering. First Col and Row then Sub
'''
Public Function nextGridPt(Optional WellFirst As Boolean) As Boolean
    If WellFirst Then
        'update mainGrid
        If Me.iRow Mod 2 <> 0 Then
            Me.iCol = Me.iCol + 1
        Else
            Me.iCol = Me.iCol - 1
        End If
        If Me.iCol > Me.numCol Then
            Me.iCol = Me.numCol
            Me.iRow = Me.iRow + 1
        End If
    
        If Me.iCol < 1 Then
            Me.iCol = 1
            Me.iRow = Me.iRow + 1
        End If
        
        'update GridSub
        If Me.iRow > Me.numRow Then
            Me.iRow = 1
            If Me.iRowSub Mod 2 <> 0 Then
                Me.iColSub = Me.iColSub + 1
                Me.iCol = 1
            Else
                Me.iColSub = Me.iColSub - 1
                Me.iCol = 1
            End If
            If Me.iColSub > Me.numColSub Then
                Me.iColSub = Me.numColSub
                Me.iRowSub = Me.iRowSub + 1
                Me.iCol = 1
            End If
            If Me.iColSub < 1 Then
                Me.iColSub = 1
                Me.iRowSub = Me.iRowSub + 1
                Me.iCol = 1
            End If
        End If
    
        ''end of grid has been reached
        If Me.iRowSub > Me.numRowSub Then
            nextGridPt = False
        Else
            nextGridPt = True
        End If
    
    
    Else
        'update GridSub
        If Me.iRowSub Mod 2 <> 0 Then
            Me.iColSub = Me.iColSub + 1
        Else
            Me.iColSub = Me.iColSub - 1
        End If
        
        If Me.iColSub > numColSub Then
            Me.iColSub = Me.numColSub
            Me.iRowSub = Me.iRowSub + 1
        End If
        
        If Me.iColSub < 1 Then
            Me.iColSub = 1
            Me.iRowSub = Me.iRowSub + 1
        End If
            
        'update mainGrid
        If Me.iRowSub > Me.numRowSub Then
            Me.iRowSub = 1
            If Me.iRow Mod 2 <> 0 Then
                Me.iCol = Me.iCol + 1
                Me.iColSub = 1
            Else
                Me.iCol = Me.iCol - 1
                Me.iColSub = 1
            End If
            If Me.iCol > Me.numCol Then
                Me.iCol = numCol
                Me.iRow = Me.iRow + 1
                Me.iColSub = 1
            End If
            If Me.iCol < 1 Then
                Me.iCol = 1
                Me.iRow = Me.iRow + 1
                Me.iColSub = 1
            End If
        End If
        
        ''end of grid has been reached
        If Me.iRow > Me.numRow Then
            nextGridPt = False
        Else
            nextGridPt = True
        End If
    End If
End Function


'''''
'   makeGridFromOnePt( pos As Vector, posGridY() As Double, posGridXY_valid() )
'   Create a Grid and a Subgrid use using  pos as starting coordinates
'       dX, dY:         Gridspacing main Grid. The main grid is the center of the subgrid
'       dXsub, dYsub:   Gridspacing subGrid
'       refRow:         reference row where main grid starts
'       refCol:         reference column where main grid starts
'''''
Public Sub makeGridFromOnePt(pos As Vector, Row As Long, Col As Long, RowSub As Long, ColSub As Long, dRow As Double, dCol As Double, _
dRowSub As Double, dColSub As Double, Optional refRow As Integer = 1, Optional refCol As Integer = 1)
        
        initialize Row, Col, RowSub, ColSub
        ' A row correspond to Y movement and Column to X shift
        'Make main grid
        Dim iRow As Long
        Dim iCol As Long
        Dim iRowSub As Long
        Dim iColSub As Long
        'Make grid and subgrid
        For iRow = 1 To Me.numRow
            For iCol = 1 To Me.numCol
                For iRowSub = 1 To Me.numRowSub
                    For iColSub = 1 To Me.numColSub
                        GridArray(iRow, iCol, iRowSub, iColSub).X = Round(pos.X + (1 - refCol) * dCol + (iCol - 1) * dCol + (iColSub - 1 - (Me.numColSub - 1) / 2) * dColSub, PrecXY)
                        GridArray(iRow, iCol, iRowSub, iColSub).Y = Round(pos.Y + (1 - refRow) * dRow + (iRow - 1) * dRow + (iRowSub - 1 - (Me.numRowSub - 1) / 2) * dRowSub, PrecXY)
                        GridArray(iRow, iCol, iRowSub, iColSub).Z = Round(pos.Z, PrecZ)
                    Next iColSub
                Next iRowSub
            Next iCol
        Next iRow
        setAllValid True

End Sub


'''''
'   makeGridFromManyPts(X() As Double, Y() As Double, Z() As Double, dXsub As Double, dYsub As Double  )
'   Create a Grid and a Subgrid use using X, Y and Z as coordinates for the main grid. A row corresponds to
'   Y movement and Column to X shift\
'       X, Y, Z:        Array containing coordinates of main grid.
'       dXsub, dYsub:   Gridspacing subGrid
'''''
Public Sub makeGridFromManyPts(pos() As Vector, Row As Long, Col As Long, RowSub As Long, ColSub As Long, dColSub As Double, dRowSub As Double)
        Dim iCol As Long
        Dim iRow As Long
        Dim iRowSub As Long
        Dim iColSub As Long
        initialize Row, Col, RowSub, ColSub
        
        If UBound(pos) + 1 <> Me.numCol Then
            MsgBox "Number of positions does not correspond to number of gridPts"
            Exit Sub
        End If
        'Make grid and subgri
        For iCol = 1 To Me.numCol
            For iRowSub = 1 To Me.numRowSub
                For iColSub = 1 To Me.numColSub
                    GridArray(1, iCol, iRowSub, iColSub).X = Round(pos(iCol - 1).X + (iColSub - 1 - (Me.numColSub - 1) / 2) * dColSub, PrecXY)
                    GridArray(1, iCol, iRowSub, iColSub).Y = Round(pos(iCol - 1).Y + (iRowSub - 1 - (Me.numRowSub - 1) / 2) * dRowSub, PrecXY)
                    GridArray(1, iCol, iRowSub, iColSub).Z = Round(pos(iCol - 1).Z, PrecZ)
                Next iColSub
            Next iRowSub
        Next iCol
        
        ' All points are true as default
        For iCol = 1 To Me.numCol
            For iRowSub = 1 To Me.numRowSub
                For iColSub = 1 To Me.numColSub
                    GridArray(1, iCol, iRowSub, iColSub).Valid = True
                Next iColSub
            Next iRowSub
        Next iCol

End Sub


Public Function getGridDimFromFile(ByVal sFile As String) As Long()
    Dim iFileNum As Integer
    Dim Fields As String
    Dim FieldEntries() As String
    Dim outDim() As Long
On Error GoTo getGridDimFromFile_Error
    If Not FileExist(sFile) Then
        MsgBox ("getGridDimFromFile: position file " & sFile & " does not exist")
        Exit Function
    End If
    Close
    iFileNum = FreeFile()
    Open sFile For Input As iFileNum
    Line Input #iFileNum, Fields
    While Left(Fields, 1) = "%"
        Line Input #iFileNum, Fields
    Wend
    FieldEntries = Split(Fields, " ")
    If UBound(FieldEntries) = 3 Then
        ReDim outDim(0 To 3)
        outDim(0) = CLng(FieldEntries(0))
        outDim(1) = CLng(FieldEntries(1))
        outDim(2) = CLng(FieldEntries(2))
        outDim(3) = CLng(FieldEntries(3))
    Else
         ReDim outDim(0)
    End If
    getGridDimFromFile = outDim
   On Error GoTo 0
   Exit Function

getGridDimFromFile_Error:

    LogManager.UpdateErrorLog "Error " & Err.number & " (" & Err.Description & _
    ") in procedure getGridDimFromFile of Class Module AGrid at line " & Erl & " "
    ReDim outDim(0)
    getGridDimFromFile = outDim
End Function

''''
' isPositionGridFile
' Check that entries of first line correspond to 4 specified values
''''
Public Function isPositionGridFile(ByVal sFile As String, ByVal Row As Integer, ByVal Col As Integer, ByVal RowSub As Integer, ByVal ColSub As Integer) As Boolean
    If Not FileExist(sFile) Then
         MsgBox ("CheckPosfile: position file " & sFile & " does not exist")
        Exit Function
    End If
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
'   Format is Me.numRow numCol numRowSub numColSub
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
    
    For iRow = 1 To Me.numRow
      For iCol = 1 To Me.numCol
          Line Input #iFileNum, Fields
          While Left(Fields, 1) = "%"
            Line Input #iFileNum, Fields
          Wend
          FieldEntries = Split(Fields, " ")
          For iRowSub = 1 To Me.numRowSub
            'needs to check this
              For iColSub = 1 To Me.numColSub
                   GridArray(iRow, iCol, iRowSub, iColSub).X = CDbl(FieldEntries((iColSub - 1) * 3 + (iRowSub - 1) * Me.numColSub * 3))
                   GridArray(iRow, iCol, iRowSub, iColSub).Y = CDbl(FieldEntries((iColSub - 1) * 3 + (iRowSub - 1) * Me.numColSub * 3 + 1))
                   GridArray(iRow, iCol, iRowSub, iColSub).Z = CDbl(FieldEntries((iColSub - 1) * 3 + (iRowSub - 1) * Me.numColSub * 3 + 2))
              Next iColSub
          Next iRowSub
      Next iCol
    Next iRow
    setAllValid True
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
    Print #iFileNum, Me.numRow & " " & Me.numCol & " " & Me.numRowSub & " " & Me.numColSub
    For iRow = 1 To Me.numRow
      For iCol = 1 To Me.numCol
          LineComm = "%Row: " & iRow & ", Col: " & iCol & " "
          Print #iFileNum, LineComm
          Line = ""
          LineComm = "%Rowsub Colsub: "
          For iRowSub = 1 To Me.numRowSub
              For iColSub = 1 To Me.numColSub
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
        If RegEx.test(Fields) Then
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
'   need to be updated
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
            For iRow = 1 To Me.numRow
                For iCol = 1 To Me.numCol
                  Line Input #iFileNum, Fields
                  While Left(Fields, 1) = "%"
                    Line Input #iFileNum, Fields
                  Wend
                  FieldEntries = Split(Fields, " ")
                  For iRowSub = 1 To Me.numRowSub
                      For iColSub = 1 To Me.numColSub
                          GridArray(iRow, iCol, iRowSub, iColSub).Valid = CBool(FieldEntries((iColSub - 1) + (iRowSub - 1) * Me.numColSub))
                      Next iColSub
                  Next iRowSub
                Next iCol
            Next iRow
        ElseIf format = "cellbase" Then
            Do While Not EOF(iFileNum)
                Line Input #iFileNum, Fields
                GoodMatch = False
                RegEx.Pattern = CellBase
                If RegEx.test(Fields) Then
                    Set Match = RegEx.Execute(Fields)
                    Active = (Match.item(0).SubMatches.item(4) <> "none")
                    GoodMatch = True
                End If
                'check it
                'if this gridposition exists then update activity
'                If GoodMatch And CInt(Match.Item(0).SubMatches.Item(1)) <= UBound(posGridXY_Valid, 1) And CInt(Match.Item(0).SubMatches.Item(1)) >= LBound(posGridXY_Valid, 1) _
'                And CInt(Match.Item(0).SubMatches.Item(2)) <= UBound(posGridXY_Valid, 2) And CInt(Match.Item(0).SubMatches.Item(2)) >= LBound(posGridXY_Valid, 2) Then
'                    For iRowSub = 1 To UBound(posGridXY_Valid, 3)
'                        For iColSub = 1 To UBound(posGridXY_Valid, 4)
'                            GridArray(CInt(Match.Item(0).SubMatches.Item(1)), CInt(Match.Item(0).SubMatches.Item(2)), iRowSub, iColSub).Valid = Active
'                            setName Match.Item(0).SubMatches.Item(3) & "_" & Match.Item(0).SubMatches.Item(4), CInt(Match.Item(0).SubMatches.Item(1)), CInt(Match.Item(0).SubMatches.Item(2)), iRowSub, iColSub
'                        Next iColSub
'                    Next iRowSub
'                End If
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
    Print #iFileNum, Me.numRow & " " & Me.numCol & " " & Me.numRowSub & " " & Me.numColSub
    For iRow = 1 To Me.numRow
      For iCol = 1 To Me.numCol
          LineComm = "%Row: " & iRow & ", Col: " & iCol & " "
          Print #iFileNum, LineComm
          Line = ""
          LineComm = "%Rowsub Colsub: "
          For iRowSub = 1 To Me.numRowSub
              For iColSub = 1 To Me.numColSub
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



