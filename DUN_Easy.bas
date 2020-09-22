Attribute VB_Name = "DUN_Easy"
Option Explicit
  
Public Sub Connect_DUN(Optional DUN_Name As String)
  'call the Shell method and run the Dial Up Networking
  'DLL to open the Dial Up Connection. Note: I think that
  'this call will open the default DUN setup for windows,
  'so if you want to open a specific DUN use the following
  'syntax to open it:
  '
  'Shell "rundll32.exe rnaui.dll,RnaDial " & DUN_Name,vbHide
  '
  'Where DUN_Name is a String Variable containing the Name
  'of the Dial Up Connection that you want to open like
  ' DUN_Name = "My Connection"
  
  Shell "rundll32.exe rnaui.dll,RnaDial", vbHide
End Sub

Public Sub Disconnect_DUN(Optional DUN_Name As String)
  'Prompt the user to confirm the disconnect if you wish to
  'make the Shell call again, after making the Shell call this
  'time, use SendKeys method to send an "Alt+C" to the active
  'window, which is the Connection Manager at this point. Alt+C
  'will effectively click the "Disconnect" button for you.
  'Then use the simple timer routine to wait for 5 seconds to
  'be sure that the connection has closed before returning
  'focus to the application that this code is in.
  
  If MsgBox("Do you want to disconnect now?", vbQuestion + vbYesNo, "You are connected to the Internet") = vbYes Then
    Shell "rundll32.exe rnaui.dll,RnaDial", vbHide
    SendKeys "{%}c"
    Wait 5
  End If
End Sub

Private Sub Wait(sngSeconds As Single)
  'simple timer
  Dim sngEndTime As Single
  sngEndTime = Timer + sngSeconds
  While Timer < sngEndTime
    DoEvents
    Wend
End Sub
