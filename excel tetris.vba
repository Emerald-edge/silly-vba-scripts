Option Explicit

' === TETRIS FOR EXCEL (arrow keys, rotate, drop, pause) ===

' Grid size and location
Private Const ROWS As Long = 20
Private Const COLS As Long = 10
Private Const START_ROW As Long = 2
Private Const START_COL As Long = 2

' Data
Private Colors As Variant          ' 1..7 color palette
Private Shapes As Variant          ' 1..7 jagged shapes (arrays of arrays)
Private Grid() As Integer          ' ROWS x COLS, stores color index (0=empty)

Private CurShape As Variant        ' current jagged shape
Private CurX As Long, CurY As Long ' top-left of shape in grid (1-based)
Private CurColor As Integer        ' 1..7

Private GameRunning As Boolean
Private GamePaused As Boolean
Private NextTick As Date

' ---------------------------
' Entry / Exit
' ---------------------------
Public Sub StartTetris()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    ws.Cells.Clear

    ReDim Grid(1 To ROWS, 1 To COLS)
    InitColors
    InitShapes
    FormatGrid ws

    Randomize
    GameRunning = True
    GamePaused = False

    BindKeys True
    SpawnPiece

    DrawGrid ws
    Application.ScreenUpdating = True

    ScheduleNextTick
End Sub

Public Sub StopTetris()
    GameRunning = False
    GamePaused = False
    BindKeys False
    ' cancel any scheduled tick
    On Error Resume Next
    If NextTick <> 0 Then Application.OnTime NextTick, "TetrisLoop", , False
    On Error GoTo 0
End Sub

' ---------------------------
' Keyboard bindings
' ---------------------------
Private Sub BindKeys(ByVal enable As Boolean)
    If enable Then
        Application.OnKey "{LEFT}", "KeyLeft"
        Application.OnKey "{RIGHT}", "KeyRight"
        Application.OnKey "{DOWN}", "KeyDown"
        Application.OnKey "{UP}", "KeyRotateCW"      ' Up = rotate CW
        Application.OnKey "z", "KeyRotateCCW"        ' Z = rotate CCW
        Application.OnKey "x", "KeyRotateCW"         ' X = rotate CW
        Application.OnKey " ", "KeyHardDrop"         ' Space = hard drop
        Application.OnKey "p", "KeyPause"            ' P = pause/resume
    Else
        Application.OnKey "{LEFT}"
        Application.OnKey "{RIGHT}"
        Application.OnKey "{DOWN}"
        Application.OnKey "{UP}"
        Application.OnKey "z"
        Application.OnKey "x"
        Application.OnKey " "
        Application.OnKey "p"
    End If
End Sub

' ---------------------------
' Init helpers
' ---------------------------
Private Sub InitColors()
    ' index 0 unused to align piece index = color index
    Colors = Array(0, _
                   RGB(255, 0, 0), _
                   RGB(0, 255, 0), _
                   RGB(0, 0, 255), _
                   RGB(255, 255, 0), _
                   RGB(0, 255, 255), _
                   RGB(255, 0, 255), _
                   RGB(255, 128, 0))
End Sub

Private Sub InitShapes()
    Dim s1 As Variant: s1 = Array( _
        Array(0, 0, 0, 0), _
        Array(1, 1, 1, 1), _
        Array(0, 0, 0, 0), _
        Array(0, 0, 0, 0))
    Dim s2 As Variant: s2 = Array( _
        Array(1, 1), _
        Array(1, 1))
    Dim s3 As Variant: s3 = Array( _
        Array(0, 1, 0), _
        Array(1, 1, 1), _
        Array(0, 0, 0))
    Dim s4 As Variant: s4 = Array( _
        Array(1, 0, 0), _
        Array(1, 1, 1), _
        Array(0, 0, 0))
    Dim s5 As Variant: s5 = Array( _
        Array(0, 0, 1), _
        Array(1, 1, 1), _
        Array(0, 0, 0))
    Dim s6 As Variant: s6 = Array( _
        Array(0, 1, 1), _
        Array(1, 1, 0), _
        Array(0, 0, 0))
    Dim s7 As Variant: s7 = Array( _
        Array(1, 1, 0), _
        Array(0, 1, 1), _
        Array(0, 0, 0))

    ReDim Shapes(1 To 7)
    Shapes(1) = s1
    Shapes(2) = s2
    Shapes(3) = s3
    Shapes(4) = s4
    Shapes(5) = s5
    Shapes(6) = s6
    Shapes(7) = s7
End Sub

Private Sub FormatGrid(ByVal ws As Worksheet)
    Dim r As Long, c As Long
    With ws
        .Cells.Font.Color = vbWhite
        .Cells.RowHeight = 15
        .Cells.ColumnWidth = 2

        For r = 1 To ROWS
            For c = 1 To COLS
                .Cells(START_ROW + r - 1, START_COL + c - 1).Interior.Color = vbBlack
            Next c
        Next r
    End With
End Sub

' ---------------------------
' Scheduler / loop
' ---------------------------
Private Sub ScheduleNextTick()
    NextTick = Now + TimeValue("00:00:01")
    Application.OnTime NextTick, "TetrisLoop"
End Sub

Public Sub TetrisLoop()
    If Not GameRunning Then Exit Sub

    If Not GamePaused Then
        If Not MovePiece(1, 0) Then
            LockPiece
            ClearLines
            SpawnPiece
            If Collision() Then
                GameRunning = False
                BindKeys False
                MsgBox "Game Over!"
                Exit Sub
            End If
        End If
    End If

    DrawGrid ActiveSheet
    ScheduleNextTick
End Sub

' ---------------------------
' Drawing
' ---------------------------
Private Sub DrawGrid(ByVal ws As Worksheet)
    Dim r As Long, c As Long

    Application.ScreenUpdating = False

    For r = 1 To ROWS
        For c = 1 To COLS
            ws.Cells(START_ROW + r - 1, START_COL + c - 1).Interior.Color = IIf(Grid(r, c) = 0, vbBlack, Colors(Grid(r, c)))
        Next c
    Next r

    ' Overlay current falling piece
    Dim h As Long, w As Long, i As Long, j As Long
    h = ShapeHeight(CurShape)
    w = ShapeWidth(CurShape)

    For i = 1 To h
        For j = 1 To w
            If ShapeCell(CurShape, i, j) = 1 Then
                Dim gy As Long, gx As Long
                gy = CurY + i - 1
                gx = CurX + j - 1
                If gy >= 1 And gy <= ROWS And gx >= 1 And gx <= COLS Then
                    ws.Cells(START_ROW + gy - 1, START_COL + gx - 1).Interior.Color = Colors(CurColor)
                End If
            End If
        Next j
    Next i

    Application.ScreenUpdating = True
End Sub

' ---------------------------
' Movement & rotation
' ---------------------------
Private Function MovePiece(ByVal dy As Long, ByVal dx As Long) As Boolean
    Dim oldY As Long, oldX As Long
    oldY = CurY: oldX = CurX

    CurY = CurY + dy
    CurX = CurX + dx

    If Collision Then
        CurY = oldY
        CurX = oldX
        MovePiece = False
    Else
        MovePiece = True
    End If
End Function

Private Function Collision() As Boolean
    Collision = CollisionAt(CurShape, CurY, CurX)
End Function

Private Function CollisionAt(ByVal shp As Variant, ByVal y As Long, ByVal x As Long) As Boolean
    Dim h As Long, w As Long, i As Long, j As Long
    h = ShapeHeight(shp)
    w = ShapeWidth(shp)

    For i = 1 To h
        For j = 1 To w
            If ShapeCell(shp, i, j) = 1 Then
                Dim ny As Long, nx As Long
                ny = y + i - 1
                nx = x + j - 1

                If nx < 1 Or nx > COLS Or ny > ROWS Then
                    CollisionAt = True
                    Exit Function
                End If

                If ny >= 1 Then
                    If Grid(ny, nx) > 0 Then
                        CollisionAt = True
                        Exit Function
                    End If
                End If
            End If
        Next j
    Next i
End Function

Private Sub LockPiece()
    Dim h As Long, w As Long, i As Long, j As Long
    h = ShapeHeight(CurShape)
    w = ShapeWidth(CurShape)

    For i = 1 To h
        For j = 1 To w
            If ShapeCell(CurShape, i, j) = 1 Then
                Dim ny As Long, nx As Long
                ny = CurY + i - 1
                nx = CurX + j - 1
                If ny >= 1 And ny <= ROWS And nx >= 1 And nx <= COLS Then
                    Grid(ny, nx) = CurColor
                End If
            End If
        Next j
    Next i
End Sub

Private Sub ClearLines()
    Dim r As Long, c As Long, rr As Long, full As Boolean

    For r = ROWS To 1 Step -1
        full = True
        For c = 1 To COLS
            If Grid(r, c) = 0 Then
                full = False
                Exit For
            End If
        Next c

        If full Then
            For rr = r To 2 Step -1
                For c = 1 To COLS
                    Grid(rr, c) = Grid(rr - 1, c)
                Next c
            Next rr
            For c = 1 To COLS
                Grid(1, c) = 0
            Next c
            r = r + 1
        End If
    Next r
End Sub

Private Sub SpawnPiece()
    Dim n As Integer
    n = Int(Rnd() * 7) + 1       ' 1..7
    CurShape = Shapes(n)
    CurColor = n

    CurY = 1
    CurX = (COLS \ 2) - (ShapeWidth(CurShape) \ 2) + 1
End Sub

' ---------------------------
' Shape utilities (jagged arrays)
' ---------------------------
Private Function ShapeHeight(ByVal shp As Variant) As Long
    ShapeHeight = UBound(shp) - LBound(shp) + 1
End Function

Private Function ShapeWidth(ByVal shp As Variant) As Long
    ShapeWidth = UBound(shp(LBound(shp))) - LBound(shp(LBound(shp))) + 1
End Function

Private Function ShapeCell(ByVal shp As Variant, ByVal i As Long, ByVal j As Long) As Integer
    ShapeCell = shp(i - 1)(j - 1) ' shapes are 0-based jagged arrays
End Function

Private Function RotateCW(ByVal shp As Variant) As Variant
    Dim h As Long, w As Long, i As Long, j As Long
    h = ShapeHeight(shp)
    w = ShapeWidth(shp)

    Dim newShp() As Variant
    ReDim newShp(0 To w - 1)

    For i = 0 To w - 1
        Dim row() As Variant
        ReDim row(0 To h - 1)
        For j = 0 To h - 1
            row(j) = shp(h - 1 - j)(i)
        Next j
        newShp(i) = row
    Next i

    RotateCW = newShp
End Function

Private Function RotateCCW(ByVal shp As Variant) As Variant
    Dim h As Long, w As Long, i As Long, j As Long
    h = ShapeHeight(shp)
    w = ShapeWidth(shp)

    Dim newShp() As Variant
    ReDim newShp(0 To w - 1)

    For i = 0 To w - 1
        Dim row() As Variant
        ReDim row(0 To h - 1)
        For j = 0 To h - 1
            row(j) = shp(j)(w - 1 - i)
        Next j
        newShp(i) = row
    Next i

    RotateCCW = newShp
End Function

Private Sub TryRotate(ByVal clockwise As Boolean)
    Dim rot As Variant
    If clockwise Then
        rot = RotateCW(CurShape)
    Else
        rot = RotateCCW(CurShape)
    End If

    ' basic wall-kick: try current, then nudge left/right
    Dim tryX As Long, k As Long
    For k = 0 To 2
        For tryX = CurX - k To CurX + k
            If Not CollisionAt(rot, CurY, tryX) Then
                CurShape = rot
                CurX = tryX
                DrawGrid ActiveSheet
                Exit Sub
            End If
        Next tryX
    Next k
    ' if all fail, no rotation
End Sub

' ---------------------------
' Key handlers
' ---------------------------
Public Sub KeyLeft()
    If Not GameRunning Or GamePaused Then Exit Sub
    If MovePiece(0, -1) Then DrawGrid ActiveSheet
End Sub

Public Sub KeyRight()
    If Not GameRunning Or GamePaused Then Exit Sub
    If MovePiece(0, 1) Then DrawGrid ActiveSheet
End Sub

Public Sub KeyDown()
    If Not GameRunning Or GamePaused Then Exit Sub
    If Not MovePiece(1, 0) Then
        LockPiece
        ClearLines
        SpawnPiece
        If Collision() Then
            GameRunning = False
            BindKeys False
            MsgBox "Game Over!"
            Exit Sub
        End If
    End If
    DrawGrid ActiveSheet
End Sub

Public Sub KeyRotateCW()
    If Not GameRunning Or GamePaused Then Exit Sub
    TryRotate True
End Sub

Public Sub KeyRotateCCW()
    If Not GameRunning Or GamePaused Then Exit Sub
    TryRotate False
End Sub

Public Sub KeyHardDrop()
    If Not GameRunning Or GamePaused Then Exit Sub
    Do While MovePiece(1, 0)
        ' keep falling
    Loop
    LockPiece
    ClearLines
    SpawnPiece
    If Collision() Then
        GameRunning = False
        BindKeys False
        MsgBox "Game Over!"
        Exit Sub
    End If
    DrawGrid ActiveSheet
End Sub

Public Sub KeyPause()
    If Not GameRunning Then Exit Sub
    GamePaused = Not GamePaused
End Sub
