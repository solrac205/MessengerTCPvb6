VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00A75021&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1470
   ClientLeft      =   15480
   ClientTop       =   13695
   ClientWidth     =   3720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E49569&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   215
      Picture         =   "frmSplash.frx":000C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   215
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E49569&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim OpenWindowsFlag As Integer

Private Sub Form_Load()
On Error Resume Next

Dim ret As Long
Dim T_rect As RECT

OpenWindowsFlag = 0
ret = SystemParametersInfo(SPI_GETWORKAREA, 0, T_rect, 0)

   Me.Top = Me.ScaleY((T_rect.Bottom - T_rect.Top), vbPixels, vbTwips) - Me.Height
   Me.Left = Me.ScaleY((T_rect.Right - T_rect.Left), vbPixels, vbTwips) - Me.Width
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim i As Integer

For i = 1 To 25
  DoEvents
  Sleep 50
  DoEvents
  Call Aplicar_Transparencia(Me.hWnd, 250 - (i * 5))
Next i

End Sub

Private Sub Timer1_Timer()
On Error GoTo Err_Timer1_Timer

 OpenWindowsFlag = OpenWindowsFlag + 1
 
  If OpenWindowsFlag = 5 Then
    DoEvents
    OpenWindowsFlag = 0
    Unload Me
  End If
  
Exit_Err_Timer1_Timer:
Exit Sub
Err_Timer1_Timer:
OpenWindowsFlag = 0
Unload Me
Resume Exit_Err_Timer1_Timer
End Sub
