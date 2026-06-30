Attribute VB_Name = "M01_Core"
Option Explicit

'=========================================================
' VBA Input Toolkit
' M01_Core
'
' Global configuration and runtime state.
'=========================================================


'=========================================================
' Palette Configuration
'=========================================================

Public Const PALETTE_LEFT As Long = 1
Public Const PALETTE_RIGHT As Long = 4
Public Const PALETTE_SHIFT_ROWS As Long = 4

Public Const PALETTE_MARK As String = "Å™"
Public Const PALETTE_MARK_COLUMN As Long = 4

Public Const PALETTE_HEADER_ROWS As Long = 5

' Smooth follow enabled/disabled
Public Const PALETTE_SMOOTH_FOLLOW As Boolean = True

' Delay per step (seconds)
Public Const PALETTE_FOLLOW_DELAY As Double = 0.01

'=========================================================
' Configuration
'=========================================================

' Automatically fill the quantity column.
Public Const AUTO_QUANTITY As Boolean = True

' Default value for the quantity column.
Public Const DEFAULT_QUANTITY As String = "1"

' Number of edits before the palette follows the cursor.
Public Const FOLLOW_INTERVAL As Long = 4


'=========================================================
' Runtime State
'=========================================================

' Current text selected from the palette.
Public PaletteText As String

' True while continuous editing is active.
Public EditStarted As Boolean

' Counter used for automatic palette following.
Public EditLoopCount As Long


'=========================================================
' Initialization
'=========================================================

Public Sub InitializeToolkit()

    PaletteText = vbNullString

    EditStarted = False
    EditLoopCount = 0

    Application.EnableEvents = True

End Sub
