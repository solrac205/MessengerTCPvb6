VERSION 5.00
Begin VB.Form frmAbaut 
   BackColor       =   &H00A75021&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5205
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbaut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3120
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1313
      Picture         =   "frmAbaut.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E49569&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   480
      Picture         =   "frmAbaut.frx":08D6
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00EBFDD2&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2775
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TCPMessenger"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00E49569&
      FillStyle       =   0  'Solid
      Height          =   4695
      Left            =   173
      Shape           =   4  'Rounded Rectangle
      Top             =   255
      Width           =   3135
   End
End
Attribute VB_Name = "frmAbaut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
On Error Resume Next

Command1.Enabled = False
Unload Me

End Sub

Private Sub Form_Load()
On Error Resume Next
Label2.Caption = "Author: " & Chr(13) & _
                 "  Carlos F. Ramírez Abdalla" & Chr(13) & Chr(13)

Label2.Caption = Label2.Caption & "Version: " & Chr(13) & _
                 "  " & App.Major & "." & App.Minor & Chr(13) & Chr(13)
                 
Label2.Caption = Label2.Caption & "Fecha de Creación: " & Chr(13) & _
                 DateToFile(App.Path & "\" & App.EXEName & ".exe", Text1, FechaCreacion) & Chr(13) & Chr(13)
                 
Label2.Caption = Label2.Caption & "Ultima Modificación: " & Chr(13) & _
                 DateToFile(App.Path & "\" & App.EXEName & ".exe", Text1, FechaModificacion) & Chr(13) & Chr(13)
                 
Command1.Enabled = True

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
On Error Resume Next
  DoEvents
End Sub
