VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSplash1 
   BackColor       =   &H00A75021&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2925
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SendFile"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2093
      Picture         =   "frmSplash1.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3773
      Picture         =   "frmSplash1.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   855
   End
   Begin MSWinsockLib.Winsock TransferFile 
      Left            =   6360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Transfiriendo:"
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
      Height          =   1455
      Left            =   300
      TabIndex        =   0
      Top             =   360
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E49569&
      BackStyle       =   1  'Opaque
      Height          =   2655
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public IpTransferFile As String
Public PortTransferFile As String
Public UserSendFile As String
Public F_Name As String
Public FileTransfer As String
Public FlagProsigue As Boolean
Public FlagOpenForm As Boolean
Public FlagConError As Boolean

Private Sub Command1_Click()
On Error Resume Next
If TransferFile.State <> 0 Then
   TransferFile.Close
End If

Unload Me
End Sub




Private Sub Command2_Click()
On Error GoTo Err_Command2_Click

Dim FileLogic As Integer
Dim Size As Long
Dim Archivo() As Byte
Dim StringSend As String
Dim TimeInicio As String
Dim TimeFin As String

TimeInicio = Format(Date, "dd-mmm-yyyy") & " - " & Format(Time, "hh:mm:ss AMPM")

Command2.Enabled = False
Command1.Enabled = False
Me.SetFocus
If TransferFile.State <> 7 Then
   MsgBox "Desconexión de Usuario", vbCritical, "Transferencia inconclusa"
   Exit Sub
Else

With frmChatForm
StringSend = ""
StringSend = .Server(0).LocalIP & "|" & Str(.Server(0).LocalPort) & "|" & LocalName & "@TRF|" & F_Name
TransferFile.SendData StringSend & vbCrLf

End With

Label1.Caption = "Solicitando Recepción....."
Me.SetFocus
Do While FlagProsigue = False
  DoEvents
Loop
FlagProsigue = False

If FlagConError = True Then
    TransferFile.Close
    FlagConError = False
    Label1.Caption = "Acción Cancelada, Usuario Rechazo envío o conexión Ocupada"
    Me.Show
    DoEvents
    Sleep 2000
    Unload Me
    Exit Sub
    
End If
Label1.Caption = "Recepción Aceptada....."
Me.SetFocus
FileLogic = FreeFile
Open FileTransfer For Binary Access Read As #FileLogic

Size = LOF(FileLogic)

ReDim Archivo(Size - 1)
Get #FileLogic, , Archivo

Close #FileLogic

With frmChatForm

Label1.Caption = "Transfiriendo: " & Format(Size, "#,###,###,###,##0") & " bytes"
Me.SetFocus
StringSend = ""
StringSend = .Server(0).LocalIP & "|" & Str(.Server(0).LocalPort) & "|" & LocalName & "@TRT|" & Size
TransferFile.SendData StringSend & vbCrLf

Do While FlagProsigue = False
  DoEvents
Loop
FlagProsigue = False

If FlagConError = True Then
    TransferFile.Close
    FlagConError = False
    Label1.Caption = "Acción Cancelada, Error en parametros de envío"
    Me.Show
    DoEvents
    Sleep 2000
    Unload Me
    Exit Sub
    
End If


TransferFile.SendData Archivo

Do While FlagProsigue = False
  DoEvents
Loop
FlagProsigue = False
Me.SetFocus
If FlagConError = True Then

    TransferFile.Close
    FlagConError = False
    Label1.Caption = "Error al Transferir Archivo, confirme recepción"
    Me.Show
    DoEvents
    Sleep 2000
    Unload Me
    Exit Sub
    
End If

StringSend = ""
StringSend = .Server(0).LocalIP & "|" & Str(.Server(0).LocalPort) & "|" & LocalName & "@TRE|Archivo Recibido y Confirmado"
TransferFile.SendData StringSend & vbCrLf

Do While FlagProsigue = False
  DoEvents
Loop
FlagProsigue = False

Me.SetFocus
If FlagConError = True Then

    TransferFile.Close
    FlagConError = False
    Label1.Caption = "Acción Cancelada, Confirmación de envío no recibida"
    Me.Show
    DoEvents
    Sleep 2000
    Unload Me
    Exit Sub
    
End If


TransferFile.Close

End With

End If
TimeFin = Format(Date, "dd-mmm-yyyy") & " - " & Format(Time, "hh:mm:ss AMPM")

Label1.Caption = "Archivo Transferido......" & Chr(13) & "Início: " & TimeInicio & _
                 Chr(13) & "Finalizo: " & TimeFin & Chr(13) & "Total Transferido: " & Format(Size, "#,###,###,###,##0")
Me.Show
DoEvents
Sleep 2000
Unload Me

Exit_Err_Command2_Click:
Exit Sub

Err_Command2_Click:
    If TransferFile.State = 7 Then
       
       With frmChatForm
       
         StringSend = ""
         StringSend = .Server(0).LocalIP & "|" & Str(.Server(0).LocalPort) & "|" & LocalName & "@TRE|Archivo Abortado " & Chr(13) & Err.Description
         TransferFile.SendData StringSend & vbCrLf
       End With
       
       Do While FlagProsigue = False
          DoEvents
       Loop
       FlagProsigue = False
 
    End If
    FlagProsigue = False
    TransferFile.Close
    FlagConError = False
    Label1.Caption = "Acción Cancelada, Se produjo error: " & Chr(13) & Err.Description
    Me.Show
    DoEvents
    Sleep 2000
    Unload Me
Resume Exit_Err_Command2_Click

End Sub

Private Sub Form_Activate()
On Error GoTo Err_Form_Activate
If TransferFile.State <> 7 And FlagOpenForm = True Then
     MsgBox "Conexión no se ha logrado Establecer, Intente mas tarde", vbInformation, "Error en transferencia"
     Unload Me
End If

FlagOpenForm = False

Exit_Err_Form_Activate:
Exit Sub

Err_Form_Activate:
Unload Me
Resume Exit_Err_Form_Activate
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim i As Integer

  Label1.Caption = Label1.Caption & Chr(13) & FileTransfer
  
  
  If TransferFile.State <> 0 Then
      TransferFile.Close
  End If
  
  FlagConError = False
  FlagProsigue = False
  TransferFile.RemoteHost = IpTransferFile
  TransferFile.RemotePort = PortTransferFile
  TransferFile.Connect
  
  For i = 1 To 50
   DoEvents
   DoEvents
   DoEvents
  Next i
  
  FlagOpenForm = True
  
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim i As Integer

If TransferFile.State <> 0 Then
    TransferFile.Close
    DoEvents
    DoEvents
End If

FlagProsigue = False
FlagConError = False

For i = 1 To 25
  DoEvents
  Sleep 50
  DoEvents
  Call Aplicar_Transparencia(Me.hWnd, 250 - (i * 5))
Next i
End Sub

Private Sub TransferFile_Close()
FlagProsigue = True
FlagConError = True
End Sub

Private Sub TransferFile_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
 Dim DataMessage As String
 Dim FlagColor As String
 Dim RichTemp As String
 Dim Encabezado As String
 
 TransferFile.GetData DataMessage
 
 If DataMessage = "OK" & vbCrLf Then
   FlagProsigue = True
   FlagConError = False
   
 Else
 
 '*****************************************************************************************************
 With frmChatForm
        If InStr(1, .RichTextBox1.TextRTF, "\cf1") <> 0 And _
           InStr(1, .RichTextBox1.TextRTF, "\cf2") <> 0 Then
             Encabezado = "{\colortbl ;\red0\green128\blue128;\red0\green0\blue255;}"
              FlagColor = "\cf0"
        Else
         If InStr(1, .RichTextBox1.TextRTF, "\cf1") <> 0 And _
           InStr(1, .RichTextBox1.TextRTF, "\cf2") = 0 Then
             Encabezado = "{\colortbl ;\red0\green128\blue128;}"
             FlagColor = "\cf0"
         Else
            If InStr(1, .RichTextBox1.TextRTF, "\cf1") = 0 And _
               InStr(1, .RichTextBox1.TextRTF, "\cf2") <> 0 Then
                 Encabezado = "{\colortbl ;\red0\green0\blue255;}"
                 FlagColor = "\cf0"
            Else
               If InStr(1, .RichTextBox1.TextRTF, "\cf1") = 0 And _
                  InStr(1, .RichTextBox1.TextRTF, "\cf2") = 0 Then
                    Encabezado = ""
                    FlagColor = "\cf0"
               End If
            End If
         End If
        End If
        
        RichTemp = ""
        If .RichTextBox1.Text <> "" Then
           RichTemp = Mid(.RichTextBox1.TextRTF, Len("{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}} "), (Len(.RichTextBox1.TextRTF) - Len("{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}} ")) - 2)
        End If
        
        RichTemp = RichTemp & "\viewkind4\uc1\pard\lang1033\f0\fs20" & FlagColor & "\b " & DataMessage & "\b0\par"
        RichTemp = Replace(RichTemp, Mid(RichTemp, InStr(1, RichTemp, "{\colortbl"), InStr(InStr(1, RichTemp, "{\colortbl"), RichTemp, "}")), "")
        
        
        .RichTextBox1.TextRTF = "{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}}" & Chr(13) & Encabezado & Chr(13) & RichTemp & Chr(13) & "\Par}"
        .RichTextBox1.SelStart = Len(.RichTextBox1.Text)
 End With
'******************************************************************************************************
 

 
 
   'frmChatForm.RichTextBox1.Text = frmChatForm.RichTextBox1.Text & DataMessage
   
   FlagProsigue = True
   FlagConError = True
 End If
 
End Sub

Private Sub TransferFile_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Dim Encabezado As String
Dim FlagColor As String
Dim RichTemp As String

'*****************************************************************************************************
 With frmChatForm
        If InStr(1, .RichTextBox1.TextRTF, "\cf1") <> 0 And _
           InStr(1, .RichTextBox1.TextRTF, "\cf2") <> 0 Then
             Encabezado = "{\colortbl ;\red0\green128\blue128;\red0\green0\blue255;}"
              FlagColor = "\cf0"
        Else
         If InStr(1, .RichTextBox1.TextRTF, "\cf1") <> 0 And _
           InStr(1, .RichTextBox1.TextRTF, "\cf2") = 0 Then
             Encabezado = "{\colortbl ;\red0\green128\blue128;}"
             FlagColor = "\cf0"
         Else
            If InStr(1, .RichTextBox1.TextRTF, "\cf1") = 0 And _
               InStr(1, .RichTextBox1.TextRTF, "\cf2") <> 0 Then
                 Encabezado = "{\colortbl ;\red0\green0\blue255;}"
                 FlagColor = "\cf0"
            Else
               If InStr(1, .RichTextBox1.TextRTF, "\cf1") = 0 And _
                  InStr(1, .RichTextBox1.TextRTF, "\cf2") = 0 Then
                    Encabezado = ""
                    FlagColor = "\cf0"
               End If
            End If
         End If
        End If
        
        RichTemp = ""
        If .RichTextBox1.Text <> "" Then
           RichTemp = Mid(.RichTextBox1.TextRTF, Len("{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}} "), (Len(.RichTextBox1.TextRTF) - Len("{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}} ")) - 2)
        End If
        
        RichTemp = RichTemp & "\viewkind4\uc1\pard\lang1033\f0\fs20" & FlagColor & "\b " & Description & "\b0\par"
        RichTemp = Replace(RichTemp, Mid(RichTemp, InStr(1, RichTemp, "{\colortbl"), InStr(InStr(1, RichTemp, "{\colortbl"), RichTemp, "}")), "")
        
        
        .RichTextBox1.TextRTF = "{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}}" & Chr(13) & Encabezado & Chr(13) & RichTemp & Chr(13) & "\Par}"
        .RichTextBox1.SelStart = Len(.RichTextBox1.Text)
 End With
'******************************************************************************************************
 
'frmChatForm.RichTextBox1.Text = frmChatForm.RichTextBox1.Text & Description

FlagProsigue = True
FlagConError = True

End Sub
