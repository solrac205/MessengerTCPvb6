VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmSplash2 
   BackColor       =   &H00A75021&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5130
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E49569&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   960
      Picture         =   "frmSplash2.frx":000C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   360
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   4800
      Top             =   4440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TCPMessenger V"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1613
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E49569&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   533
      Shape           =   2  'Oval
      Top             =   120
      Width           =   4815
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   3960
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   5970
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   10530
      _cy             =   6985
   End
End
Attribute VB_Name = "frmSplash2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
On Error Resume Next

  WindowsMediaPlayer1.URL = App.Path & "\LogoEngage.avi"
  Label1.Caption = Label1.Caption & App.Major & "." & App.Minor & Chr(13) & App.CompanyName
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Dim i As Integer

i = 0
WindowsMediaPlayer1.Close
For i = 1 To 25
  DoEvents
  Sleep 75
  DoEvents
  
    Call Aplicar_Transparencia(Me.hWnd, 250 - (i * 5))
  
Next i

Me.Visible = False
Sleep 75
frmChatForm.Show


End Sub

Private Sub Timer1_Timer()

Unload Me


End Sub
