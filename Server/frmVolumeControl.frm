VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmVolumeControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Master Volume Control"
   ClientHeight    =   4035
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3390
   Icon            =   "frmVolumeControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   1650
      Top             =   30
   End
   Begin MSWinsockLib.Winsock tcpVolCont 
      Index           =   0
      Left            =   2130
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "About"
      Height          =   2835
      Left            =   90
      TabIndex        =   2
      Top             =   1140
      Width           =   3225
      Begin VB.Label Label6 
         Caption         =   "Your vote will be very much be appreciated."
         Height          =   435
         Left            =   135
         TabIndex        =   8
         Top             =   2340
         Width           =   2355
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Douglas Cavan  -  d_cavan@lycos.com"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   2100
         Width           =   2820
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Developed by:"
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   1830
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmVolumeControl.frx":058A
         Height          =   1365
         Left            =   135
         TabIndex        =   3
         Top             =   300
         Width           =   2820
         WordWrap        =   -1  'True
      End
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   675
      Left            =   45
      TabIndex        =   1
      Top             =   420
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1191
      _Version        =   327682
      LargeChange     =   1000
      SmallChange     =   100
      Max             =   65535
      SelectRange     =   -1  'True
      TickStyle       =   2
      TickFrequency   =   4250
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   570
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Max"
      Height          =   195
      Left            =   2970
      TabIndex        =   5
      Top             =   210
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Min"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   210
      Width           =   255
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show Master Volume"
      End
   End
End
Attribute VB_Name = "frmVolumeControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Dim mySysTray As NOTIFYICONDATA

Dim nCon As Integer


Private Sub Form_Load()
   nCon = 0
   tcpVolCont(0).LocalPort = 3210
   tcpVolCont(0).Listen
   Slider1.Value = 33457
   Text1 = Slider1.Value
   Call Slider1_Scroll
   With mySysTray
      .cbSize = Len(mySysTray)
      .hWnd = Me.hWnd
      .uId = 1&
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .ucallbackMessage = WM_LBUTTONDOWN
      .hIcon = Me.Icon
      .szTip = "Master Volume Controller by greedy" & Chr$(0)
      Shell_NotifyIcon NIM_ADD, mySysTray
   End With

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim Msg As Long
   Msg = X / Screen.TwipsPerPixelX
   If Msg = WM_LBUTTONDBLCLK Then
       Me.Show
   ElseIf Msg = WM_RBUTTONUP Then
       Me.PopupMenu mnuPopup
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = 0 Then
      Cancel = 1
      Me.Hide
   Else
      Call mnuExit_Click
   End If
End Sub

Private Sub mnuExit_Click()
Dim i As Integer
On Error Resume Next
   With mySysTray
      .cbSize = Len(mySysTray)
      .hWnd = Me.hWnd
      .uId = 1&
      Shell_NotifyIcon NIM_DELETE, mySysTray
   End With
   For i = 0 To nCon
      tcpVolCont(i).Close
   Next
   End
End Sub

Private Sub mnuShow_Click()
   Me.Show
End Sub

Private Sub Slider1_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
      Text1 = Slider1.Value
   End If
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Text1.Text = Slider1.Value
End Sub

Private Sub Slider1_Scroll()
   Slider1.SelStart = 0
   Slider1.SelLength = Val(Slider1.Value)
End Sub

Private Sub tcpVolCont_ConnectionRequest(Index As Integer, ByVal requestID As Long)
   If Index = 0 Then
      nCon = nCon + 1
      Load tcpVolCont(nCon)
      tcpVolCont(nCon).LocalPort = 0
      tcpVolCont(nCon).Accept requestID
      Timer1.Enabled = True
   End If
End Sub

Private Sub Text1_Change()
   On Error Resume Next
   Dim i As Integer
   For i = 1 To nCon
      tcpVolCont(i).SendData CStr(Text1)
   Next
End Sub

'Timer is use to make sure that value in the textbox is sent to all clients
Private Sub Timer1_Timer()
   On Error Resume Next
   Dim i As Integer
   
   For i = 1 To nCon
      tcpVolCont(i).SendData CStr(Text1)
   Next
   Timer1.Enabled = False
End Sub
