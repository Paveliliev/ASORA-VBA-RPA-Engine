Attribute VB_Name = "Module5"
' ==============================================================================
' Project: Audio Notification & Speech Engine
' Author: Pavel Iliev
' Purpose: Provides real-time audio feedback for background automation tasks.
' Logic: Uses Windows Multimedia API (MCI) for MP3/WAV and SAPI for Text-to-Speech.
' Professional Spin: Enhances accessibility and allows for asynchronous
'                   monitoring of long-running data processes.
' ==============================================================================

Option Explicit

' --- WINDOWS MULTIMEDIA API ---
' This allows Excel to play external audio files without opening a media player.
#If VBA7 Then
    Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
        (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
         ByVal uReturnLength As Long, ByVal hwndCallback As LongPtr) As Long
#Else
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
        (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
         ByVal uReturnLength As Long, ByVal uReturnLength As Long) As Long
#End If

' ==============================================================================
' 1. CORE HELPER: PlaySoundEffect
' Instead of writing the same code for every sound, we use this one shared engine.
' ==============================================================================
Private Sub PlaySoundEffect(ByVal filePath As String)
    ' Close any existing audio stream named "bot_audio" to prevent overlaps
    mciSendString "close bot_audio", 0, 0, 0
    
    ' Attempt to open and play the specified file
    ' We use an alias ("bot_audio") so we can control it later
    On Error Resume Next
    mciSendString "open """ & filePath & """ type mpegvideo alias bot_audio", 0, 0, 0
    mciSendString "play bot_audio", 0, 0, 0
    On Error GoTo 0
End Sub

' ==============================================================================
' 2. SPECIFIC NOTIFICATION TRIGGERS
' Call these subs at the end of your data loops to notify the user.
' ==============================================================================

Sub Notification_Success()
    ' Professional "Success" sound
    Call PlaySoundEffect("C:\Users\pavel.iliev\Desktop\Sound Effects\Job's Finished.mp3")
End Sub

Sub Notification_Alert()
    ' Used for unexpected data patterns or errors
    Call PlaySoundEffect("C:\Users\pavel.iliev\Desktop\Sound Effects\vine-boom.mp3")
End Sub

Sub Notification_Warning()
    ' For data discrepancies
    Call PlaySoundEffect("C:\Users\pavel.iliev\Desktop\Sound Effects\brother-ewwwwwww.mp3")
End Sub

Sub Stop_All_Audio()
    ' Immediate kill-switch for all bot audio
    mciSendString "stop bot_audio", 0, 0, 0
    mciSendString "close bot_audio", 0, 0, 0
End Sub

' ==============================================================================
' 3. TEXT-TO-SPEECH (TTS) ENGINE
' Uses the Speech API (SAPI) to provide verbal status updates.
' ==============================================================================

Sub Bot_Status_Update(Optional ByVal customMessage As String)
    Dim vls As Object
    Set vls = CreateObject("SAPI.SpVoice")
    
    ' Set preferred voice (0 is usually the system default)
    Set vls.voice = vls.GetVoices.Item(0)
    vls.Rate = 0     ' Normal speaking speed
    vls.Volume = 100 ' Max volume
    
    ' If no message is passed, read from a specific Excel cell (A15)
    If customMessage = "" Then customMessage = Range("A15").Text
    
    vls.Speak customMessage
End Sub

' Example: Specifically for greeting/starting a process
Sub Bot_Greet_User()
    Call Bot_Status_Update("Automation engine online. Awaiting data input.")
End Sub

