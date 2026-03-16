Attribute VB_Name = "Module3"
' ==============================================================================
' Project: LARA (Logistics Automated Relocation Assistant)
' Module: Bin-to-Bin Transfer & Transaction Logging
' Author: Pavel Iliev
' Business Value: Automates slow system stock movements in the ERP.
' Key Feature: Real-time transaction logging for audit and inventory compliance.
' ==============================================================================

Option Explicit

' --- WINDOWS API DECLARATIONS (64-Bit Compatible) ---
#If VBA7 Then
    Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As LongPtr)
    Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
    Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If

' Constants
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const VK_ESCAPE = &H1B

' ==============================================================================
' SUBROUTINE: Bin_To_Bin_Transfer
' Purpose: Executes bulk bin transfers and records movements in a digital log.
' ==============================================================================
Sub Bin_To_Bin_Transfer()
    ' --- Configuration Arrays ---
    ' Using arrays to store coordinates (X, Y) and Sleep Multipliers (S)
    Dim coordsX(2 To 7) As Long, coordsY(2 To 7) As Long, sMods(2 To 7) As Double
    
    ' --- Data & Logic Variables ---
    Dim linesToComplete As Long, fromBin As String, toBin As String, repeatLimit As Long
    Dim i As Long, t As Long, lastLogRow As Long
    Dim progName As String: progName = "Project Pick Location Manager"
    
    ' --- 1. Load System Configuration ---
    ' Dynamically pulls UI map from the setup sheet (Rows 2 to 7)
    For i = 2 To 7
        coordsX(i) = Range("B" & i).Value
        coordsY(i) = Range("C" & i).Value
        sMods(i) = IIf(Range("D" & i).Value = 0, 1, Range("D" & i).Value)
    Next i
    
    linesToComplete = Range("B11").Value
    
    ' --- 2. Main Row Processing Loop ---
    For i = 1 To linesToComplete
        fromBin = Range("H" & i + 1).Value
        toBin = Range("I" & i + 1).Value
        repeatLimit = Range("J" & i + 1).Value
        
        ' Skip logic for empty data rows
        If fromBin = "" Then GoTo NextLine
        
        ' --- 3. Repetition Loop (Transaction Execution) ---
        For t = 1 To repeatLimit
            ' FAIL-SAFE: Allow the OS to process events and check for ESCAPE key
            DoEvents
            If GetAsyncKeyState(VK_ESCAPE) <> 0 Then
                MsgBox "LARA Halted: User-initiated abort via Escape key.", vbCritical
                Exit Sub
            End If

            ' Window Management
            On Error Resume Next
            AppActivate progName
            On Error GoTo 0
            
            ' STEP A: Source Bin Selection
            ClickAtCoordinates coordsX(2), coordsY(2)
            Sleep 100 * sMods(2)
            SendKeys "^a{BACKSPACE}" & fromBin & "{ENTER}", True
            Sleep 500 * sMods(2)
            
            ' STEP B: Target Bin Selection
            ClickAtCoordinates coordsX(3), coordsY(3)
            Sleep 100 * sMods(3)
            SendKeys "^a{BACKSPACE}" & toBin & "{ENTER}", True
            Sleep 500 * sMods(3)
            
            ' STEP C: UI Execution Sequence (Move Action)
            ' Double-click source to activate move
            ClickAtCoordinates coordsX(4), coordsY(4): Sleep 50
            ClickAtCoordinates coordsX(4), coordsY(4): Sleep 400 * sMods(4)
            
            ' Confirm Target and Finalize Move
            ClickAtCoordinates coordsX(5), coordsY(5): Sleep 400 * sMods(5)
            ClickAtCoordinates coordsX(6), coordsY(6): Sleep 400 * sMods(6)
            
            ' Final Arrow/Submit Click
            ClickAtCoordinates coordsX(7), coordsY(7)
            Sleep 600 * sMods(7)
            
            ' --- 4. AUDIT LOGGING ---
            ' Identifies the next available log entry point (starting from row 20)
            lastLogRow = Cells(Rows.Count, 1).End(xlUp).Row
            If lastLogRow < 19 Then lastLogRow = 19
            
            With Cells(lastLogRow + 1, 1)
                .Offset(0, 0).Value = fromBin       ' Col A: Source
                .Offset(0, 1).Value = toBin         ' Col B: Destination
                .Offset(0, 2).Value = Now           ' Col C: Timestamp
                .Offset(0, 2).NumberFormat = "dd/mm/yyyy hh:mm:ss"
            End With
            
        Next t
        
        ' Visual Confirmation: Row processed successfully
        Range("H" & i + 1).Interior.Color = vbGreen
        
NextLine:
    Next i
    
    MsgBox "LARA Status: Batch processing and audit logging completed.", vbInformation
End Sub

' --- CORE HELPER: Mouse Interaction ---
Private Sub ClickAtCoordinates(x As Long, y As Long)
    SetCursorPos x, y
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub
