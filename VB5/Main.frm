VERSION 5.00
Object = "{616B1177-4E85-11D3-AF35-A916D26ACA3B}#1.0#0"; "CDNOTIFY.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CDNotification Demo"
   ClientHeight    =   3324
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   4608
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3324
   ScaleWidth      =   4608
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton btnHomepage 
      Caption         =   "&Homepage"
      Height          =   372
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   1092
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   372
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   1092
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Play sound"
      Height          =   1572
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4332
      Begin VB.CommandButton btnRemoved 
         Caption         =   ">>"
         Height          =   252
         Left            =   3720
         TabIndex        =   10
         Top             =   1080
         Width           =   372
      End
      Begin VB.CommandButton btnArrived 
         Caption         =   ">>"
         Height          =   252
         Left            =   3720
         TabIndex        =   8
         Top             =   480
         Width           =   372
      End
      Begin VB.TextBox txtRemoved 
         Height          =   288
         Left            =   480
         TabIndex        =   9
         Top             =   1080
         Width           =   3132
      End
      Begin VB.TextBox txtArrived 
         Height          =   288
         Left            =   480
         TabIndex        =   7
         Top             =   480
         Width           =   3132
      End
      Begin VB.Label Label4 
         Caption         =   "When removed"
         Height          =   252
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   2172
      End
      Begin VB.Label Label3 
         Caption         =   "When arrived"
         Height          =   252
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   2292
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   0
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
      DefaultExt      =   ".wav"
      Filter          =   "Wave file|*.wav|All files|*.*"
      Flags           =   4096
   End
   Begin CDNotification.CDNotify CDNotify1 
      Left            =   3000
      Top             =   0
      _ExtentX        =   677
      _ExtentY        =   677
      Enabled         =   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   "Copyright (c) 1999 Hai Li, Zeal SoftStudio."
      Height          =   372
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   4332
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please open or close your CDRom tray"
      ForeColor       =   &H8000000D&
      Height          =   492
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   3012
   End
   Begin VB.Label Label1 
      Caption         =   "Event:"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1092
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub btnArrived_Click()
    CommonDialog1.FileName = txtArrived.Text
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        txtArrived.Text = CommonDialog1.FileName
        sndPlaySound txtArrived.Text, 1
    End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnHomepage_Click()
    ShellExecute 0, "open", "http://members.tripod.com/~zealsoft", "", "", 0
End Sub

Private Sub btnRemoved_Click()
    CommonDialog1.FileName = txtRemoved.Text
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        txtRemoved.Text = CommonDialog1.FileName
        sndPlaySound txtRemoved.Text, 1
    End If
End Sub

Private Sub CDNotify1_Arrival(ByVal Drive As String)
    Label2.Caption = "Drive " + Drive + ": arrived."
    If txtArrived.Text <> "" Then
        sndPlaySound txtArrived.Text, 1
    End If
End Sub

Private Sub CDNotify1_RemoveComplete(ByVal Drive As String)
    Label2.Caption = "Drive " + Drive + ": was removed."
    If txtRemoved.Text <> "" Then
        sndPlaySound txtRemoved.Text, 1
    End If
End Sub
