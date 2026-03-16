Attribute VB_Name = "Module2"
' ==============================================================================
' Project: SKU-Level Location Manager (Temp & De-Assign Engines)
' Author: Pavel Iliev
' Functionality: Handles batch processing of SKU assignments and removals by
'                synchronizing Excel data with "Project Pick Location Manager."
' ==============================================================================

Option Explicit

' --- API Fail-Safe Declaration ---
#If VBA7 Then
    Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As LongPtr)
    Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    ' Timing and Keyboard Fail-Safe
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
    (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, ByVal hwndCallback As LongPtr) As Long
#End If

Const VK_ESCAPE = &H1B
' Constants for API Operations
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4

' --- CORE HELPER: Mouse Interaction ---
Private Sub ClickAtCoordinates(x As Long, y As Long)
    SetCursorPos x, y
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub


' ==============================================================================
' 1. SUBROUTINE: Temp_By_SKU
' Purpose: Changing Permanant to Temporary Fixed Pick Locations.
' ==============================================================================
Sub Temp_By_SKU()
    ' --- Variable Declarations ---
    Dim xCoords(1 To 8) As Long, yCoords(1 To 8) As Long
    Dim sMods(1 To 8) As Double
    Dim linesToComplete As Long, globalSleep As Double
    Dim i As Long, x As Long, innerLoopLimit As Long
    Dim sku As String, progName As String: progName = "Project Pick Location Manager"
    
    ' --- 1. Load Configuration from Excel (Control Panel) ---
    ' This block maps UI coordinates and timing offsets defined in the spreadsheet
    For i = 1 To 8
        xCoords(i) = Range("B" & i).Value
        yCoords(i) = Range("C" & i).Value
        sMods(i) = IIf(Range("D" & i).Value = 0, 1, Range("D" & i).Value)
    Next i
    
    linesToComplete = Range("B11").Value
    globalSleep = IIf(Range("D11").Value = 0, 1, Range("D11").Value)

    ' --- 2. Main Processing Loop (Rows) ---
    For i = 1 To linesToComplete
        ' Fail-Safe: Emergency stop
        If GetAsyncKeyState(VK_ESCAPE) <> 0 Then GoTo ExitPoint
        
        ' Load SKU specific data
        innerLoopLimit = Range("N" & i + 1).Value ' Number of iterations per SKU
        sku = Range("M" & i + 1).Value
        
        Sleep 200 * globalSleep
        
        ' Initial UI Focus: Double-click to select search field
        ClickAtCoordinates xCoords(1), yCoords(1): Sleep 40
        ClickAtCoordinates xCoords(1), yCoords(1): Sleep 450 * sMods(1)
        
        ' Input SKU and navigate
        SendKeys sku, True: Sleep 650 * sMods(1)
        ClickAtCoordinates xCoords(2), yCoords(2): Sleep 40
        ClickAtCoordinates xCoords(2), yCoords(2): Sleep 800 * sMods(2)

        ' --- 3. Nested Loop: Multi-Location Processing ---
        For x = 1 To innerLoopLimit
            DoEvents
            On Error Resume Next
            AppActivate progName
            If Err.Number <> 0 Then
                MsgBox "Application '" & progName & "' not found.", vbCritical: Exit Sub
            End If
            On Error GoTo 0
            
            Sleep 900 ' Wait for UI to stabilize
            
            ' Navigation Sequence Line selection and Repl Button
            ClickAtCoordinates xCoords(3), yCoords(3): Sleep 400 * sMods(3)
            ClickAtCoordinates xCoords(4), yCoords(4): Sleep 500 * sMods(4)
            
            ' Open Assignment Dialog / Repl field
            ClickAtCoordinates xCoords(5), yCoords(5): Sleep 30
            ClickAtCoordinates xCoords(5), yCoords(5): Sleep 500 * sMods(5)
            
            ' Field Entry Sequence: Clears and resets Replenishment/Required/Max quantities
            ' Uses Backspace logic to ensure fields are empty before typing "0"
            SendKeys "{BKSP 5}0", True: Sleep 300
            SendKeys "{TAB}{BKSP 10}0", True: Sleep 350
            SendKeys "{TAB}{BKSP 10}0", True: Sleep 350
            
            ' Confirmation Sequence and "Lag" handler
            ClickAtCoordinates xCoords(6), yCoords(6): Sleep 350 * sMods(6)
            ClickAtCoordinates xCoords(7), yCoords(7): Sleep 200 * sMods(7)
            ClickAtCoordinates xCoords(8), yCoords(8): Sleep 200 * sMods(8)
        Next x
        
        ' Audit Trail: Mark as processed
        Range("M" & i + 1).Interior.Color = vbGreen
    Next i
    Exit Sub

ExitPoint:
    MsgBox "Automation Halted: Escape key detected.", vbCritical
End Sub

' ==============================================================================
' 2. SUBROUTINE: De_Assigning_By_SKU
' Purpose: Rapid removal of Fixed Pick Locations.
' ==============================================================================
Sub De_Assigning_By_SKU()
    Dim xCoords(1 To 8) As Long, yCoords(1 To 8) As Long
    Dim sMods(1 To 8) As Double
    Dim linesToComplete As Long, globalSleep As Double
    Dim i As Long, x As Long, innerLoopLimit As Long
    Dim sku As String, progName As String: progName = "Project Pick Location Manager"
    
    ' Load UI Mappings
    For i = 1 To 8
        xCoords(i) = Range("B" & i).Value
        yCoords(i) = Range("C" & i).Value
        sMods(i) = IIf(Range("D" & i).Value = 0, 1, Range("D" & i).Value)
    Next i
    
    linesToComplete = Range("B11").Value
    globalSleep = Range("D11").Value

    ' Process SKU Rows
    For i = 1 To linesToComplete
        If GetAsyncKeyState(VK_ESCAPE) <> 0 Then GoTo ExitPoint
        
        innerLoopLimit = Range("N" & i + 1).Value
        sku = Range("M" & i + 1).Value
        
        Sleep 200 * globalSleep
        
        ' Select SKU Field
        ClickAtCoordinates xCoords(1), yCoords(1): Sleep 40
        ClickAtCoordinates xCoords(1), yCoords(1): Sleep 450 * sMods(1)
        
        ' Search/Confermation Button
        SendKeys sku, True: Sleep 650 * sMods(1)
        ClickAtCoordinates xCoords(2), yCoords(2): Sleep 40
        ClickAtCoordinates xCoords(2), yCoords(2): Sleep 800 * sMods(2)

        ' Removal Loop
        For x = 1 To innerLoopLimit
            DoEvents
            On Error Resume Next
            AppActivate progName
            On Error GoTo 0
            
            Sleep 900
            
            ' Execution of removal clicks
            ClickAtCoordinates xCoords(3), yCoords(3): Sleep 400 * sMods(3)
            ClickAtCoordinates xCoords(4), yCoords(4): Sleep 500 * sMods(4)
            ClickAtCoordinates xCoords(6), yCoords(6): Sleep 350 * sMods(6)
            ClickAtCoordinates xCoords(7), yCoords(7): Sleep 200 * sMods(7)
        Next x
        
        Range("M" & i + 1).Interior.Color = vbGreen
    Next i
    Exit Sub

ExitPoint:
    MsgBox "Automation Halted: Escape key detected.", vbCritical
End Sub
