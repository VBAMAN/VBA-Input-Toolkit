Attribute VB_Name = "M04_AutoNumberEngine"
Option Explicit

'=========================================================
' VBA Input Toolkit
' M04_AutoNumberEngine
'
' Automatic serial number assignment.
'=========================================================

'=========================================================
' FillSerialNumber
'
' Automatically assigns the next serial number.
'=========================================================

Public Sub FillSerialNumber(ByVal Target As Range)

    Dim Number As Long

    Number = GetNextSerialNumber(Target)

    Application.EnableEvents = False

    Target.Value = Number

    If EditStarted Then

        EditLoopCount = EditLoopCount + 1

        If EditLoopCount >= FOLLOW_INTERVAL Then

            EditLoopCount = 0

            FollowPalette Target.Row

        End If

    End If

    Application.EnableEvents = True

End Sub

'=========================================================
' GetNextSerialNumber
'
' Returns the next available serial number.
'=========================================================

Private Function GetNextSerialNumber(ByVal Target As Range) As Long

    Dim r As Long

    For r = Target.Row - 1 To 1 Step -1

        If Target.Worksheet.Cells(r, 7).Value <> "" Then

            If IsNumeric(Target.Worksheet.Cells(r, 7).Value) Then

                GetNextSerialNumber = _
                    CLng(Target.Worksheet.Cells(r, 7).Value) + 1

            Else

                GetNextSerialNumber = 1

            End If

            Exit Function

        End If

    Next r

    GetNextSerialNumber = 1

End Function
