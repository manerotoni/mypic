VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImagingGrids"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Grids As New Collection


Private Sub MsgNoGrid(Name As String)
    MsgBox "ImagingGrids: Grid " & Name & " has not been defined"
End Sub


'''
' Add a new grid to collection
'''
Public Sub AddGrid(Name As String, Optional numRow As Long = 1, Optional numCol As Long = 1, _
Optional numRowSub As Long = 1, Optional numColSub As Long = 1)
    If InCollection(Grids, Name) Then
        Grids(Name).initialize numRow, numCol, numRowSub, numColSub
    Else
        Dim Grid As New AGrid
        Grid.NameGrid = Name
        Grid.initialize numRow, numCol, numRowSub, numColSub
        Grids.Add Grid, Name
    End If
End Sub

'''
'
'''
Public Sub updateGrid(Name As String, numRow As Long, numCol As Long, numRowSub As Long, numColSub As Long)
    If InCollection(Grids, Name) Then
        Grids(Name).initialize numRow, numCol, numRowSub, numColSub
    Else
        MsgNoGrid Name
    End If
End Sub


Public Sub redimPreserve(Name As String, numRow As Long, numCol As Long, numRowSub As Long, numColSub As Long)
    If InCollection(Grids, Name) Then
        Grids(Name).redimensionPreserve numRow, numCol, numRowSub, numColSub
    Else
        MsgNoGrid Name
    End If
End Sub

Public Function isPositionGridFile(Name As String, ByVal sFile As String, ByVal Row As Integer, _
ByVal Col As Integer, ByVal RowSub As Integer, ByVal ColSub As Integer) As Boolean
    If InCollection(Grids, Name) Then
        isPositionGridFile = Grids(Name).isPositionGridFile(sFile, Row, Col, RowSub, ColSub)
    Else
        MsgNoGrid Name
    End If
End Function

'''
'
'''
Public Function loadPositionGridFile(Name As String, sFile As String) As Boolean
    If InCollection(Grids, Name) Then
        loadPositionGridFile = Grids(Name).loadPositionGridFile(Name, sFile)
    Else
        MsgNoGrid Name
    End If
End Function


Public Function writePositionGridFile(Name As String, sFile As String) As Boolean
    If InCollection(Grids, Name) Then
        writePositionGridFile = Grids(Name).writePositionGridFile(sFile)
    Else
        MsgNoGrid Name
    End If
End Function


Public Function isValidGridFile(Name As String, ByVal sFile As String, ByVal Row As Integer, _
ByVal Col As Integer, ByVal RowSub As Integer, ByVal ColSub As Integer) As String
    If InCollection(Grids, Name) Then
        isValidGridFile = Grids(Name).isValidGridFile(sFile, Row, Col, RowSub, ColSub)
    Else
        MsgNoGrid Name
    End If
End Function

'''
'
'''
Public Function loadValidGridFile(Name As String, sFile As String, format As String) As Boolean
    If InCollection(Grids, Name) Then
        loadValidGridFile = Grids(Name).loadValidGridFile(Name, sFile, format)
    Else
        MsgNoGrid Name
    End If
End Function

Public Function setIndeces(Name As String, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Boolean
    If InCollection(Grids, Name) Then
        setIndeces = Grids(Name).setIndeces(Row, Col, RowSub, ColSub)
    Else
        MsgNoGrid Name
    End If
End Function

Public Function nextGridPt(Name As String) As Boolean
    If InCollection(Grids, Name) Then
        nextGridPt = Grids(Name).nextGridPt
    Else
        MsgNoGrid Name
    End If
End Function



Public Function getThisGridPoint(Name As String, X As Double, Y As Double, Z As Double, Valid As Boolean)
    If InCollection(Grids, Name) Then
        Grids(Name).getThisGridPoint X, Y, Z, Valid
    Else
        MsgNoGrid Name
    End If
End Function

Public Function getThisX(Name As String) As Double
    If InCollection(Grids, Name) Then
        getThisX = Grids(Name).getThisX
    Else
        MsgNoGrid Name
    End If
End Function

Public Sub setThisX(Name As String, Value As Double)
    If InCollection(Grids, Name) Then
        Grids(Name).setThisX = Value
    Else
        MsgNoGrid Name
    End If
End Sub


Public Function getThisY(Name As String) As Double
    If InCollection(Grids, Name) Then
        getThisY = Grids(Name).getThisY
    Else
        MsgNoGrid Name
    End If
End Function

Public Sub setThisY(Name As String, Value As Double)
    If InCollection(Grids, Name) Then
        Grids(Name).setThisY = Value
    Else
        MsgNoGrid Name
    End If
End Sub

Public Function getThisZ(Name As String) As Double
    If InCollection(Grids, Name) Then
        getThisZ = Grids(Name).getThisZ
    Else
        MsgNoGrid Name
    End If
End Function

Public Sub setThisZ(Name As String, Value As Double)
    If InCollection(Grids, Name) Then
        Grids(Name).setThisZ = Value
    Else
        MsgNoGrid Name
    End If
End Sub

Public Function getThisValid(Name As String) As Boolean
    If InCollection(Grids, Name) Then
        getThisValid = Grids(Name).getThisValid
    Else
        MsgNoGrid Name
    End If
End Function

Public Sub setThisValid(Name As String, Value As Boolean)
    If InCollection(Grids, Name) Then
        Grids(Name).setThisValid = Value
    Else
        MsgNoGrid Name
    End If
End Sub


Public Function getX(Name As String, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Double
    If InCollection(Grids, Name) Then
        getX = Grids(Name).getX(Row, Col, RowSub, ColSub)
    Else
        MsgNoGrid Name
    End If
End Function


Public Sub setX(Name As String, Value As Double, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If InCollection(Grids, Name) Then
        Grids(Name).setX Value, Row, Col, RowSub, ColSub
    Else
        MsgNoGrid Name
    End If
End Sub

Public Function getY(Name As String, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Double
    If InCollection(Grids, Name) Then
        getY = Grids(Name).getY(Row, Col, RowSub, ColSub)
    Else
        MsgNoGrid Name
    End If
End Function

Public Sub setY(Name As String, Value As Double, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If InCollection(Grids, Name) Then
        Grids(Name).setY Value, Row, Col, RowSub, ColSub
    Else
        MsgNoGrid Name
    End If
End Sub

Public Function getZ(Name As String, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Double
    If InCollection(Grids, Name) Then
        getZ = Grids(Name).getZ(Row, Col, RowSub, ColSub)
    Else
        MsgNoGrid Name
    End If
End Function

Public Sub setZ(Name As String, Value As Double, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If InCollection(Grids, Name) Then
        Grids(Name).setZ Value, Row, Col, RowSub, ColSub
    Else
        MsgNoGrid Name
    End If
End Sub

Public Function getValid(Name As String, Row As Long, Col As Long, RowSub As Long, ColSub As Long) As Boolean
    If InCollection(Grids, Name) Then
        getValid = Grids(Name).getValid(Row, Col, RowSub, ColSub)
    Else
        MsgNoGrid Name
    End If
End Function

Public Sub setValid(Name As String, Value As Boolean, Row As Long, Col As Long, RowSub As Long, ColSub As Long)
    If InCollection(Grids, Name) Then
        Grids(Name).setValid Value, Row, Col, RowSub, ColSub
    Else
        MsgNoGrid Name
    End If
End Sub


Public Function writeValidGridFile(Name As String, sFile As String) As Boolean
    If InCollection(Grids, Name) Then
        writeValidGridFile = Grids(Name).writeValidGridFile(sFile)
    Else
        MsgNoGrid Name
    End If
End Function



Public Sub makeGridFromOnePt(Name As String, XStart As Double, YStart As Double, ZStart As Double, dX As Double, dY As Double, _
dXsub As Double, dYsub As Double, Optional refCol As Integer = 1, Optional refRow As Integer = 1)
    If InCollection(Grids, Name) Then
        Grids(Name).makeGridFromOnePt XStart, YStart, ZStart, dX, dY, dXsub, dYsub, refCol, refRow
    Else
        MsgNoGrid Name
    End If
End Sub

Public Sub makeGridFromManyPts(Name As String, X() As Double, Y() As Double, Z() As Double, dXsub As Double, dYsub As Double)
    If InCollection(Grids, Name) Then
        Grids(Name).makeGridFromManyPts X, Y, Z, dXsub, dYsub
    Else
        MsgNoGrid Name
    End If
End Sub


Public Function thisSuffix(Name As String) As String
    If InCollection(Grids, Name) Then
        thisSuffix = Grids(Name).thisSuffix
    Else
        MsgNoGrid Name
    End If
End Function

Public Function thisSuffixWell(Name As String) As String
    If InCollection(Grids, Name) Then
        thisSuffixWell = Grids(Name).thisSuffixWell
    Else
        MsgNoGrid Name
    End If
End Function


Public Function getThisName(Name As String) As String
    If InCollection(Grids, Name) Then
        getThisName = Grids(Name).getThisName
    Else
        MsgNoGrid Name
    End If
End Function


Public Sub setThisName(Name As String, Value As String)
    If InCollection(Grids, Name) Then
        Grids(Name).setThisName
    Else
        MsgNoGrid Value
    End If
End Sub

Public Function getName(Name As String, numRow As Long, numCol As Long, numRowSub As Long, numColSub As Long) As String
    If InCollection(Grids, Name) Then
        getName = Grids(Name).getName(numRow, numCol, numRowSub, numColSub)
    Else
        MsgNoGrid Name
    End If
End Function

Public Sub setName(Name As String, Value As String, numRow As Long, numCol As Long, numRowSub As Long, numColSub As Long)
    If InCollection(Grids, Name) Then
        Grids(Name).setName Value, numRow, numCol, numRowSub, numColSub
    Else
        MsgNoGrid Name
    End If
End Sub


Public Function suffix(Name As String, numRow As Long, numCol As Long, numRowSub As Long, numColSub As Long) As String
    If InCollection(Grids, Name) Then
        suffix = Grids(Name).suffix(numRow, numCol, numRowSub, numColSub)
    Else
        MsgNoGrid Name
    End If
End Function

Public Function numCol(Name As String) As Integer
    numCol = Grids(Name).numCol
End Function


Public Function numRow(Name As String) As Integer
    numRow = Grids(Name).numCol
End Function


Public Function numColSub(Name As String) As Integer
    numColSub = Grids(Name).numColSub
End Function


Public Function numRowSub(Name As String) As Integer
    numRowSub = Grids(Name).numRowSub
End Function
