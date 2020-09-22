Attribute VB_Name = "ModMain"
Option Explicit

Public gStrServerName As String

Sub Main()
   'Set this app to run everytime the computer starts
   SetKeyValue "Software\Microsoft\Windows\Currentversion\Run", _
              "VolControl", App.Path & "\" & App.EXEName & ".exe", REG_SZ
   'Get server name
   gStrServerName = QueryValue("Software\VolumeController\Client", "ServerName")

   If Len(gStrServerName) Then
      Load frmReceiver
   Else
      frmConfiguration.Show
   End If
   
End Sub

