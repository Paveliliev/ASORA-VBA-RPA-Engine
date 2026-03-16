Attribute VB_Name = "Module1"
' ==============================================================================
' Project: Automated Location Management Engine (The "Location Bot")
' Author: Pavel Iliev
' Purpose: Automates manual data entry and UI interaction between Excel and Reflex.
' Result: 70%+ Increase in process efficiency; 20+ hours saved weekly.
' ==============================================================================

Option Explicit

' --- WINDOWS API DECLARATIONS (64-Bit Compatible) ---
' These allow VBA to interact directly with the Windows OS for mouse/keyboard control.

#If VBA7 Then
    ' Mouse and Cursor Control
    Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As LongPtr)
    Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    ' Timing and Keyboard Fail-Safe
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
    Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

' Constants for API Operations
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const VK_ESCAPE = &H1B

' ==============================================================================
' CORE HELPER TOOLS
' ==============================================================================

' Helper: Handles Mouse Movement and Clicks in one call
Private Sub ClickAtCoordinates(x As Long, y As Long)
    SetCursorPos x, y
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

' ==============================================================================
' MAIN AUTOMATION: Location De-Assignment
' ==============================================================================
Sub Location_De_Assigne_By_Order()
    Dim repeatCount As Long, i As Long
    Dim sleepMod As Single
    Dim muteFirst As Integer, muteSecond As Integer
    
    ' --- CONFIGURATION LOAD ---
    ' Pulls user-defined parameters from the "Control Panel" worksheet
    On Error Resume Next
    repeatCount = Range("B5").Value
    sleepMod = IIf(Range("B7").Value = 0, 1, Range("B7").Value) ' Default to 1 if empty
    muteFirst = Range("B9").Value
    muteSecond = Range("B10").Value
    On Error GoTo 0

    If repeatCount <= 0 Then
        MsgBox "Please enter a valid cycle count in cell B5.", vbExclamation
        Exit Sub
    End If

    ' Safety Buffer: Gives user 2 seconds to switch to the target application
    Application.Wait (Now + TimeValue("0:00:02"))
    
    ' --- EXECUTION LOOP ---
    For i = 1 To repeatCount
        ' Fail-Safe: Immediate halt if ESCAPE key is pressed
        If GetAsyncKeyState(VK_ESCAPE) <> 0 Then
            MsgBox "Automation Aborted: Escape key detected.", vbCritical
            Exit Sub
        End If
        
        DoEvents ' Keeps Excel responsive during the loop

        ' UI Interaction Sequence
        ' Note: Coordinates are pulled dynamically from the worksheet for flexibility
        ClickAtCoordinates Range("B2").Value, Range("C2").Value
        Sleep 500 * sleepMod
        
        ClickAtCoordinates Range("B3").Value, Range("C3").Value
        Sleep 1000 * sleepMod
        
        ClickAtCoordinates Range("B4").Value, Range("C4").Value
        Sleep 2000 * sleepMod
        
        ' Audio feedback for loop completion (User Experience)
        If muteFirst = 1 Then Call vine_boom
    Next i
    
    ' Final Audio Cue and Status Report
    If muteSecond = 1 Then Call Jobs_Finished
    MsgBox "Task Completed: " & repeatCount & " cycles processed successfully.", vbInformation
End Sub

' ==============================================================================
' EXTERNAL INTEGRATION: Paste to Third-Party Program (Reflex/Project Manager)
' ==============================================================================
Sub Pick_Face_Assign_By_SKU()
    Dim progName As String: progName = "Project Pick Location Manager"
    Dim repeatCount As Long, i As Long
    Dim sku As String, repl As String, maxQ As String

    repeatCount = Range("B7").Value

    For i = 1 To repeatCount
        ' Emergency Stop Check
        If GetAsyncKeyState(VK_ESCAPE) <> 0 Then Exit Sub

        ' Data Extraction from dynamic range
        sku = Range("M" & i + 1).Value
        repl = Range("Q" & i + 1).Value
        maxQ = Range("R" & i + 1).Value

        ' Switch focus to target External Program
        On Error Resume Next
        AppActivate progName
        If Err.Number <> 0 Then
            MsgBox "Error: Target program '" & progName & "' not found.", vbCritical
            Exit Sub
        End If
        On Error GoTo 0
        
        Sleep 500
        
        ' --- COORDINATED DATA ENTRY ---
        ' Uses the ClickAtCoordinates helper to navigate UI menus
        ClickAtCoordinates Range("B1").Value, Range("C1").Value ' Row Selection
        Sleep 500
        ClickAtCoordinates Range("B2").Value, Range("C2").Value ' Assign Button
        Sleep 500
        
        ' SKU Entry
        SendKeys sku, True
        SendKeys "{TAB}", True: Sleep 1000
        
        ' Replenishment and Max Quantity Logic
        ClickAtCoordinates Range("B4").Value, Range("C4").Value
        SendKeys "{BKSP}" & repl & "{TAB}", True: Sleep 100
        SendKeys "{BKSP}" & maxQ & "{TAB}", True: Sleep 100
        SendKeys "{BKSP}" & maxQ & "{ENTER}", True: Sleep 550
        
        ' Confirmation Button
        ClickAtCoordinates Range("B5").Value, Range("C5").Value
        Sleep 250
        ClickAtCoordinates Range("B6").Value, Range("C6").Value
        
        ' Visual Feedback: Mark row green in Excel once processed
        Range("M" & i + 1).Interior.Color = vbGreen
        DoEvents
    Next i
End Sub



' Tool: Captures current mouse X/Y and copies to clipboard for configuration
Sub WhereIsMyMouse()
    Dim pos As POINTAPI
    Dim clipboard As Object
    
    ' Wait 2 seconds for user to position mouse over target UI element
    Application.Wait (Now + TimeValue("0:00:02"))
    GetCursorPos pos
    
    Dim coordString As String: coordString = pos.x & ", " & pos.y
    
    ' Copy to Clipboard for easy pasting into Excel config cells
    On Error Resume Next
    Set clipboard = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    clipboard.SetText coordString
    clipboard.PutInClipboard
    On Error GoTo 0
    
    MsgBox "Coordinates Saved!" & vbCrLf & "X: " & pos.x & " | Y: " & pos.y, vbInformation
End Sub
