VERSION 5.00
Begin VB.Form frmConfiguration 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Volume Control"
   ClientHeight    =   2280
   ClientLeft      =   4440
   ClientTop       =   3285
   ClientWidth     =   4425
   Icon            =   "frmConfiguration.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3210
      TabIndex        =   3
      Top             =   1800
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2010
      TabIndex        =   2
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   45
      TabIndex        =   0
      Top             =   690
      Width           =   4335
      Begin VB.TextBox txtServerName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   4005
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Please enter the name of the computer where the Master Volume Control is installed."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   120
      TabIndex        =   4
      Top             =   90
      Width           =   4095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
   'Create new key
   CreateNewKey "Software\VolumeController\Client", _
                HKEY_LOCAL_MACHINE
   'Set new key value
   SetKeyValue "Software\VolumeController\Client", _
              "ServerName", UCase(txtServerName.Text), REG_SZ
   Unload Me
   
   gStrServerName = GetStringValue("HKEY_LOCAL_MACHINE\Software\VolumeController\Client" _
                                 , "ServerName")
   Load frmReceiver
End Sub

Private Sub Command2_Click()
   End
End Sub

