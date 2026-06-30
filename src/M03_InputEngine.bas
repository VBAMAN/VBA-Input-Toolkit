Attribute VB_Name = "M03_InputEngine"
Option Explicit

'=========================================================
' VBA Input Toolkit
' M03_InputEngine
'
' Palette selection and worksheet input.
'=========================================================

'=========================================================
' SelectPaletteItem
'
' Stores the selected palette text.
'=========================================================

Public Sub SelectPaletteItem(ByVal Target As Range)

    PaletteText = Target.Value

End Sub

'=========================================================
' PastePaletteItem
'
' Writes the selected palette text into the target cell.
'=========================================================

Public Sub PastePaletteItem(ByVal Target As Range)

    If Len(PaletteText) = 0 Then Exit Sub

    Target.Value = PaletteText

    ApplyDefaultQuantity Target

    Target.Offset(0, 2).Activate

End Sub

'=========================================================
' ApplyDefaultQuantity
'
' Automatically fills the quantity column.
'=========================================================

Private Sub ApplyDefaultQuantity(ByVal Target As Range)

    If Not AUTO_QUANTITY Then Exit Sub

    If Target.Offset(0, 1).Value = "" Then
        Target.Offset(0, 1).Value = DEFAULT_QUANTITY
    End If

End Sub

