Attribute VB_Name = "M02_PaletteEngine"
Option Explicit

'=========================================================
' Palette Engine
'
' Move palette up, down or back to the top.
'=========================================================

Public Enum PaletteMove

    pmDown = 1
    pmUp = 2
    pmTop = 3

End Enum

Public Sub MovePalette(ByVal MoveMode As PaletteMove)

    Dim highestRow As Long
    Dim lowestRow As Long
    Dim shiftRows As Long
    Dim i As Long, j As Long

    Dim ws As Worksheet
    Dim palette As Variant

    Set ws = ActiveSheet

    highestRow = GetPaletteTopRow()
    lowestRow = GetPaletteBottomRow()

    ' Prevent moving above row 1.
    If MoveMode = pmUp Then
        If highestRow - PALETTE_SHIFT_ROWS < 1 Then Exit Sub
    End If

    i = lowestRow - highestRow + 1
    j = PALETTE_RIGHT - PALETTE_LEFT + 1

    Select Case MoveMode

        Case pmDown
            shiftRows = PALETTE_SHIFT_ROWS

        Case pmUp
            shiftRows = -PALETTE_SHIFT_ROWS

        Case pmTop
            shiftRows = 1 - GetPaletteTopRow()

    End Select

    palette = ws.Range( _
                ws.Cells(highestRow, PALETTE_LEFT), _
                ws.Cells(lowestRow, PALETTE_RIGHT) _
              ).Value

    ws.Range( _
        ws.Cells(highestRow, PALETTE_LEFT), _
        ws.Cells(lowestRow, PALETTE_RIGHT) _
    ).ClearContents

    ws.Range( _
        ws.Cells(highestRow + shiftRows, PALETTE_LEFT), _
        ws.Cells(lowestRow + shiftRows, PALETTE_RIGHT) _
    ).Value = palette

    Application.EnableEvents = False
    ActiveCell.End(xlToLeft).Activate
    Application.EnableEvents = True

    PaletteText = vbNullString

End Sub

'=========================================================
' FollowPalette
'
' Move the palette near the current editing row.
' Supports optional smooth animation.
'=========================================================

Public Sub FollowPalette(ByVal TargetRow As Long)

    Dim highestRow As Long
    Dim lowestRow As Long
    Dim shiftRows As Long
    Dim newTopRow As Long

    Dim palette As Variant
    Dim ws As Worksheet

    Dim maxRow As Long
    Dim stepRow As Long
    Dim i As Long

    Set ws = ActiveSheet

    highestRow = GetPaletteTopRow()
    lowestRow = GetPaletteBottomRow()

    shiftRows = TargetRow - highestRow - PALETTE_HEADER_ROWS

    If shiftRows = 0 Then Exit Sub

    '----------------------------
    ' Clamp (top limit)
    '----------------------------
    newTopRow = highestRow + shiftRows

    If newTopRow < 1 Then
        shiftRows = 1 - highestRow
    End If

    '----------------------------
    ' Clamp (bottom limit)
    '----------------------------
    maxRow = ws.Rows.count

    If newTopRow + (lowestRow - highestRow) > maxRow Then
        shiftRows = maxRow - lowestRow
    End If

    '=====================================================
    ' SMOOTH MODE (optional)
    '=====================================================
    If PALETTE_SMOOTH_FOLLOW And shiftRows <> 0 Then

        stepRow = Sgn(shiftRows)

        For i = 1 To Abs(shiftRows)

            Call MovePaletteByStep(stepRow)

            Call PaletteDelay(PALETTE_FOLLOW_DELAY)

        Next i

        Exit Sub

    End If

    '=====================================================
    ' INSTANT MODE (original behavior)
    '=====================================================

    palette = ws.Range( _
                ws.Cells(highestRow, PALETTE_LEFT), _
                ws.Cells(lowestRow, PALETTE_RIGHT) _
              ).Value

    ws.Range( _
        ws.Cells(highestRow, PALETTE_LEFT), _
        ws.Cells(lowestRow, PALETTE_RIGHT) _
    ).ClearContents

    ws.Range( _
        ws.Cells(highestRow + shiftRows, PALETTE_LEFT), _
        ws.Cells(lowestRow + shiftRows, PALETTE_RIGHT) _
    ).Value = palette

End Sub

Private Sub MovePaletteByStep(ByVal stepRow As Long)

    Dim ws As Worksheet
    Dim topRow As Long, bottomRow As Long
    Dim palette As Variant

    Set ws = ActiveSheet

    topRow = GetPaletteTopRow()
    bottomRow = GetPaletteBottomRow()

    palette = ws.Range( _
                ws.Cells(topRow, PALETTE_LEFT), _
                ws.Cells(bottomRow, PALETTE_RIGHT) _
              ).Value

    ws.Range( _
        ws.Cells(topRow, PALETTE_LEFT), _
        ws.Cells(bottomRow, PALETTE_RIGHT) _
    ).ClearContents

    ws.Range( _
        ws.Cells(topRow + stepRow, PALETTE_LEFT), _
        ws.Cells(bottomRow + stepRow, PALETTE_RIGHT) _
    ).Value = palette

End Sub

Public Sub PaletteDelay(ByVal seconds As Double)

    Dim t As Double
    t = Timer

    Do While Timer < t + seconds
        DoEvents
    Loop

End Sub

'=========================================================
' GetPaletteTopRow
'
' Returns the first row of the palette.
' The palette position is identified by the "ª" marker.
'=========================================================

Private Function GetPaletteTopRow() As Long

    Dim f As Range

    Set f = ActiveSheet.Columns(PALETTE_MARK_COLUMN).Find( _
                What:=PALETTE_MARK, _
                LookIn:=xlValues, _
                LookAt:=xlWhole, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlNext)
    
    If f Is Nothing Then
        GetPaletteTopRow = 1
    Else
        GetPaletteTopRow = f.Row
    End If

End Function

'=========================================================
' GetPaletteBottomRow
'
' Returns the last used row of the palette.
'=========================================================

Private Function GetPaletteBottomRow() As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentLastRow As Long
    Dim col As Long

    Set ws = ActiveSheet

    lastRow = 0

    For col = PALETTE_LEFT To PALETTE_RIGHT

        currentLastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row

        If currentLastRow > lastRow Then
            lastRow = currentLastRow
        End If

    Next col

    GetPaletteBottomRow = lastRow

End Function
