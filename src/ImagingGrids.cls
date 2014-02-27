VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImagingGrids"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''
' A Class to manage several AGrid instances (commodity class in absence of inheritance in VBA)
'''

Option Explicit
Private Grids() As AGrid
Private GridNames As Dictionary

Private Sub MsgNoGrid(Name As String)
    MsgBox "ImagingGrids: Grid " & Name & " has not been defined"
End Sub

'''
' Create a set of grids with specified names. The default size is 0,0,0,0
'''
Public Sub initialize(Names() As String)
    Dim i As Integer
    ReDim Grids(0 To UBound(Names))
    Set GridNames = New Dictionary
    For i = 0 To UBound(Names)
        AddGrid (Names(i))
    Next i
End Sub



'''
'   Add a new grid to array
'''
Public Sub AddGrid(Name As String, Optional numRow As Long = 0, Optional numCol As Long = 0, _
Optional numRowSub As Long = 0, Optional numColSub As Long = 0)
    Dim iGrid As Integer
    If GridNames Is Nothing Then
        Set GridNames = New Dictionary
    End If
    ' only create once the grid with a specific name
    If GridNames.Exists(Name) Then
        Exit Sub
    End If
    
    If isArrayEmpty(Grids) Then
        ReDim Grids(0)
    Else
        ReDim Preserve Grids(0 To UBound(Grids) + 1)
    End If
    
    GridNames.Add Name, UBound(Grids)
    iGrid = GridNames.item(Name)
    Set Grids(iGrid) = New AGrid
    Grids(iGrid).NameGrid = Name
    Grids(iGrid).initialize numRow, numCol, numRowSub, numColSub
End Sub

Public Function checkGridName(Name As String) As Boolean
    If GridNames.Exists(Name) Then
        checkGridName = True
    Else
        MsgNoGrid Name
    End If
End Function


Public Function isGridEmpty(Name As String) As Boolean
    If checkGridName(Name) Then
        isGridEmpty = Grids(GridNames.item(Name)).isGridEmpty
    End If
End Function


Public Sub updateGridSize(Name As String, numRow As Long, numCol As Long, numRowSub As Long, numColSub As Long)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).initialize numRow, numCol, numRowSub, numColSub
    End If
End Sub


Public Sub updateGridSizePreserve(Name As String, numRow As Long, numCol As Long, numRowSub As Long, numColSub As Long)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).updateGridSizePreserve numRow, numCol, numRowSub, numColSub
    End If
End Sub

Public Function getGridDimFromFile(Name As String, ByVal sFile As String) As Long()
    If checkGridName(Name) Then
        getGridDimFromFile = Grids(GridNames.item(Name)).getGridDimFromFile(sFile)
    End If
End Function



Public Function isPositionGridFile(Name As String, ByVal sFile As String, ByVal Row As Integer, _
ByVal Col As Integer, ByVal RowSub As Integer, ByVal ColSub As Integer) As Boolean
    If checkGridName(Name) Then
        isPositionGridFile = Grids(GridNames.item(Name)).isPositionGridFile(sFile, Row, Col, RowSub, ColSub)
    End If
End Function

'''
'
'''
Public Function loadPositionGridFile(Name As String, sFile As String) As Boolean
    If checkGridName(Name) Then
        loadPositionGridFile = Grids(GridNames.item(Name)).loadPositionGridFile(sFile)
    End If
End Function


Public Function writePositionGridFile(Name As String, sFile As String) As Boolean
    If checkGridName(Name) Then
        writePositionGridFile = Grids(GridNames.item(Name)).writePositionGridFile(sFile)
    End If
End Function


Public Function isValidGridFile(Name As String, ByVal sFile As String, ByVal Row As Long, _
ByVal Col As Long, ByVal RowSub As Long, ByVal ColSub As Long) As String
    If checkGridName(Name) Then
        isValidGridFile = Grids(GridNames.item(Name)).isValidGridFile(sFile, Row, Col, RowSub, ColSub)
    End If
End Function


Public Function setIsRunning(Name As String, value As Boolean)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setIsRunning (value)
    End If
End Function

Public Function getIsRunning(Name As String) As Boolean
    If checkGridName(Name) Then
        getIsRunning = Grids(GridNames.item(Name)).getIsRunning
    End If
End Function

Public Function loadValidGridFile(Name As String, sFile As String, format As String) As Boolean
    If checkGridName(Name) Then
        loadValidGridFile = Grids(GridNames.item(Name)).loadValidGridFile(sFile, format)
    End If
End Function

Public Function setIndeces(Name As String, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Boolean
    If checkGridName(Name) Then
        setIndeces = Grids(GridNames.item(Name)).setIndeces(Row, Col, RowSub, ColSub)
    End If
End Function

Public Function nextGridPt(Name As String, Optional WellFirst As Boolean) As Boolean
    If checkGridName(Name) Then
        nextGridPt = Grids(GridNames.item(Name)).nextGridPt(WellFirst)
    End If
End Function

Public Function thisRow(Name As String) As Integer
    If checkGridName(Name) Then
        thisRow = Grids(GridNames.item(Name)).iRow
    End If
End Function


Public Function thisColumn(Name As String) As Integer
    If checkGridName(Name) Then
        thisColumn = Grids(GridNames.item(Name)).iCol
    End If
End Function

Public Function thisSubRow(Name As String) As Integer
    If checkGridName(Name) Then
        thisSubRow = Grids(GridNames.item(Name)).iRowSub
    End If
End Function


Public Function thisSubColumn(Name As String) As Integer
    If checkGridName(Name) Then
        thisSubColumn = Grids(GridNames.item(Name)).iColSub
    End If
End Function

Public Function getThisPt(Name As String, Pos As Vector, Valid As Boolean)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).getThisPt Pos, Valid
    End If
End Function


Public Function setThisPt(Name As String, Pos As Vector, Valid As Boolean)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setThisPt Pos, Valid
    End If
End Function

Public Function getNrPts(Name As String) As Long
    If checkGridName(Name) Then
        getNrPts = Grids(GridNames.item(Name)).getNrPts
    End If
End Function


Public Function getNrValidPts(Name As String) As Long
    If checkGridName(Name) Then
        getNrValidPts = Grids(GridNames.item(Name)).getNrValidPts
    End If
End Function

Public Function getPt(Name As String, Pos As Vector, Valid As Boolean, _
Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).getPt Pos, Valid, Row, Col, RowSub, ColSub
    End If
End Function

Public Function setPt(Name As String, Pos As Vector, Valid As Boolean, _
Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setPt Pos, Valid, Row, Col, RowSub, ColSub
    End If
End Function

Public Function getThisX(Name As String) As Double
    If checkGridName(Name) Then
        getThisX = Grids(GridNames.item(Name)).getThisX
    End If
End Function

Public Sub setThisX(Name As String, value As Double)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setThisX value
    End If
End Sub


Public Function getThisY(Name As String) As Double
    If checkGridName(Name) Then
        getThisY = Grids(GridNames.item(Name)).getThisY
    End If
End Function

Public Sub setThisY(Name As String, value As Double)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setThisY value
    End If
End Sub

Public Function getThisZ(Name As String) As Double
    If checkGridName(Name) Then
        getThisZ = Grids(GridNames.item(Name)).getThisZ
    End If
End Function

Public Sub setThisZ(Name As String, value As Double)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setThisZ value
    End If
End Sub

Public Function getThisValid(Name As String) As Boolean
    If checkGridName(Name) Then
        getThisValid = Grids(GridNames.item(Name)).getThisValid
    End If
End Function

Public Sub setThisValid(Name As String, value As Boolean)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setThisValid value
    End If
End Sub

Public Sub setAllValid(Name As String, value As Boolean)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setAllValid value
    End If
End Sub



Public Function getX(Name As String, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Double
    If checkGridName(Name) Then
        getX = Grids(GridNames.item(Name)).getX(Row, Col, RowSub, ColSub)
    End If
End Function


Public Sub setX(Name As String, value As Double, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setX value, Row, Col, RowSub, ColSub
    End If
End Sub

Public Function getY(Name As String, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Double
    If checkGridName(Name) Then
        getY = Grids(GridNames.item(Name)).getY(Row, Col, RowSub, ColSub)
    End If
End Function

Public Sub setY(Name As String, value As Double, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setY value, Row, Col, RowSub, ColSub
    End If
End Sub

Public Function getZ(Name As String, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Double
    If checkGridName(Name) Then
        getZ = Grids(GridNames.item(Name)).getZ(Row, Col, RowSub, ColSub)
    End If
End Function

Public Sub setZ(Name As String, value As Double, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setZ value, Row, Col, RowSub, ColSub
    End If
End Sub

Public Function getValid(Name As String, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Boolean
    If checkGridName(Name) Then
        getValid = Grids(GridNames.item(Name)).getValid(Row, Col, RowSub, ColSub)
    End If
End Function

Public Sub setValid(Name As String, value As Boolean, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setValid value, Row, Col, RowSub, ColSub
    End If
End Sub


Public Function writeValidGridFile(Name As String, sFile As String) As Boolean
    If checkGridName(Name) Then
        writeValidGridFile = Grids(GridNames.item(Name)).writeValidGridFile(sFile)
    End If
End Function


'''TODO: These two functions should be fused into one'''
Public Sub makeGridFromOnePt(Name As String, Pos As Vector, Row As Long, Col As Long, RowSub As Long, ColSub As Long, dRow As Double, dCol As Double, _
dRowSub As Double, dColSub As Double, Optional refRow As Integer = 1, Optional refCol As Integer = 1)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).makeGridFromOnePt Pos, Row, Col, RowSub, ColSub, dRow, dCol, dRowSub, dColSub, refRow, refCol
    End If
End Sub

Public Sub makeGridFromManyPts(Name As String, Pos() As Vector, Row As Long, Col As Long, RowSub As Long, ColSub As Long, _
dRowSub As Double, dColSub As Double)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).makeGridFromManyPts Pos, Row, Col, RowSub, ColSub, dRowSub, dColSub
    End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function thisSuffix(Name As String) As String
    If checkGridName(Name) Then
        thisSuffix = Grids(GridNames.item(Name)).thisSuffix
    End If
End Function

Public Function thisSuffixWell(Name As String) As String
    If checkGridName(Name) Then
        thisSuffixWell = Grids(GridNames.item(Name)).thisSuffixWell
    End If
End Function


Public Function getThisName(Name As String) As String
    If checkGridName(Name) Then
        getThisName = Grids(GridNames.item(Name)).getThisName
    End If
End Function


Public Function getThisParentPath(Name As String) As String
    If checkGridName(Name) Then
        getThisParentPath = Grids(GridNames.item(Name)).getThisParentPath
    End If
End Function


Public Sub setThisName(Name As String, value As String)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setThisName value
    End If
End Sub

Public Sub setThisParentPath(Name As String, value As String)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setThisParentPath value
    End If
End Sub


Public Sub setAllParentPath(Name As String, value As String)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setAllParentPath value
    End If
End Sub


Public Function getName(Name As String, numRow As Long, numCol As Long, numRowSub As Long, numColSub As Long) As String
    If checkGridName(Name) Then
        getName = Grids(GridNames.item(Name)).getName(numRow, numCol, numRowSub, numColSub)
    End If
End Function



Public Function getParentPath(Name As String, numRow As Long, numCol As Long, numRowSub As Long, numColSub As Long) As String
    If checkGridName(Name) Then
        getParentPath = Grids(GridNames.item(Name)).getParentPath(numRow, numCol, numRowSub, numColSub)
    End If
End Function

Public Sub setParentPath(Name As String, value As String, numRow As Long, numCol As Long, numRowSub As Long, numColSub As Long)
    If checkGridName(Name) Then
        Grids(GridNames.item(Name)).setParentPath value, numRow, numCol, numRowSub, numColSub
    End If
End Sub


Public Function suffix(Name As String, numRow As Long, numCol As Long, numRowSub As Long, numColSub As Long) As String
    If checkGridName(Name) Then
        suffix = Grids(GridNames.item(Name)).suffix(numRow, numCol, numRowSub, numColSub)
    End If
End Function

Public Function numCol(Name As String) As Integer
    If checkGridName(Name) Then
        numCol = Grids(GridNames.item(Name)).numCol
    End If
End Function


Public Function numRow(Name As String) As Integer
    If checkGridName(Name) Then
        numRow = Grids(GridNames.item(Name)).numCol
    End If
End Function


Public Function numColSub(Name As String) As Integer
    If checkGridName(Name) Then
        numColSub = Grids(GridNames.item(Name)).numColSub
    End If
End Function


Public Function numRowSub(Name As String) As Integer
    If checkGridName(Name) Then
        numRowSub = Grids(GridNames.item(Name)).numRowSub
    End If
End Function
