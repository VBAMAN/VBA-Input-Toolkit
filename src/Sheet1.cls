VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'=========================================================
' VBA Input Toolkit
' Worksheet Events
'
' Handles user interaction and delegates processing
' to each engine.
'=========================================================

Private Enum PaletteMode

    pmDisabled = 0
    pmEnabled = 1

End Enum

Private CurrentPaletteMode As PaletteMode

Private Const INPUT_COLUMN As Long = 8
Private Const SERIAL_COLUMN As Long = 7
Private Const OVERWRITE_CONFIRM As Boolean = True

Private Sub Worksheet_Change(ByVal Target As Range)

    If Target.Address <> "$A$1" Then Exit Sub

    If UCase$(Trim$(Target.Value)) = "TOP" Then

        Application.EnableEvents = False

        MovePalette pmTop

        Target.ClearContents

        Application.EnableEvents = True

        MsgBox "Palette moved to the top."

    End If

End Sub

'=========================================================
' Worksheet_SelectionChange
'
' Entry point for all worksheet interactions.
'=========================================================

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.CountLarge > 1 Then Exit Sub

    If HandlePalette(Target) Then Exit Sub

    If HandleInput(Target) Then Exit Sub

    If HandleSerialNumber(Target) Then Exit Sub

End Sub

'=========================================================
' HandlePalette
'
' Handles all palette operations.
'
' Returns True when the selection was handled.
'=========================================================

Private Function HandlePalette(ByVal Target As Range) As Boolean

    ' Outside the palette area.
    If Target.Column < PALETTE_LEFT _
    Or Target.Column > PALETTE_RIGHT Then
        Exit Function
    End If

    Select Case UCase$(CStr(Target.Value))

        Case "ON"

            CurrentPaletteMode = pmEnabled
            EditStarted = True

        Case "OFF"

            CurrentPaletteMode = pmDisabled
            PaletteText = vbNullString
            EditStarted = True

        Case "Å™"

            MovePalette pmUp
            EditStarted = True

        Case "Å´"

            MovePalette pmDown
            EditStarted = True

        Case Else

            If CurrentPaletteMode = pmEnabled Then

                SelectPaletteItem Target
                EditStarted = True

            Else
                Exit Function
            End If

    End Select

    HandlePalette = True

End Function

'=========================================================
' HandleInput
'
' Handles worksheet input using the selected palette item.
'
' Returns True when the selection was handled.
'=========================================================

Private Function HandleInput(ByVal Target As Range) As Boolean

    '========================
    ' EARLY EXIT (OFF MODE)
    '========================
    If CurrentPaletteMode = pmDisabled Then
        Exit Function
    End If

    ' Input column only.
    If Target.Column <> INPUT_COLUMN Then Exit Function

    If Not OVERWRITE_CONFIRM Then

        PastePaletteItem Target
        HandleInput = True
        Exit Function

    End If

    ' Confirm overwrite.
    If Not IsEmpty(Target.Value) Then

        If MsgBox( _
            "This cell already contains data." & vbCrLf & _
            "Do you want to overwrite it?", _
            vbYesNo + vbExclamation, _
            "Confirm") = vbNo Then

            HandleInput = True
            Exit Function

        End If

    End If

    PastePaletteItem Target

    HandleInput = True

End Function

'=========================================================
' HandleSerialNumber
'
' Auto-increments serial number column.
' Triggers palette follow when needed.
'
' Returns True when handled.
'=========================================================

Private Function HandleSerialNumber(ByVal Target As Range) As Boolean

    Dim r As Long
    Dim v As Variant
    Dim nextNo As Long

    ' Only serial column
    If Target.Column <> SERIAL_COLUMN Then Exit Function

    Application.EnableEvents = False

    ' Search upward for last numeric value
    For r = Target.Row - 1 To 1 Step -1

        v = Target.Worksheet.Cells(r, SERIAL_COLUMN).Value

        If v <> "" Then

            If IsNumeric(v) Then
                nextNo = CLng(v) + 1
            Else
                nextNo = 1
            End If

            Exit For

        End If

    Next r

    ' If nothing found above
    If nextNo = 0 Then nextNo = 1

    Target.Value = nextNo

    ' Palette follow control
    If EditStarted Then

        EditLoopCount = EditLoopCount + 1

        If EditLoopCount >= FOLLOW_INTERVAL Then

            EditLoopCount = 0

            FollowPalette Target.Row

        End If

    End If

    Application.EnableEvents = True

    HandleSerialNumber = True

End Function


