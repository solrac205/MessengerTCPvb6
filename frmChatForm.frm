VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmChatForm 
   BackColor       =   &H00A75021&
   Caption         =   "Messenger TCP"
   ClientHeight    =   8745
   ClientLeft      =   5280
   ClientTop       =   3255
   ClientWidth     =   8535
   Icon            =   "frmChatForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   8535
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   5400
      TabIndex        =   11
      Top             =   4440
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Emot Icons"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   95
         Width           =   1815
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   13
         Left            =   1920
         Picture         =   "frmChatForm.frx":0CCA
         ToolTipText     =   "Burro"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   345
         Index           =   12
         Left            =   1320
         Picture         =   "frmChatForm.frx":109B
         ToolTipText     =   "Dormir"
         Top             =   1815
         Width           =   360
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   11
         Left            =   720
         Picture         =   "frmChatForm.frx":13B0
         ToolTipText     =   "Silencio"
         Top             =   1785
         Width           =   360
      End
      Begin VB.Image Image2 
         Height          =   360
         Index           =   10
         Left            =   120
         Picture         =   "frmChatForm.frx":1707
         ToolTipText     =   "Angel"
         Top             =   1800
         Width           =   360
      End
      Begin VB.Image Image2 
         Height          =   345
         Index           =   9
         Left            =   2520
         Picture         =   "frmChatForm.frx":1A9C
         ToolTipText     =   "Enojado"
         Top             =   1080
         Width           =   360
      End
      Begin VB.Image Image2 
         Height          =   375
         Index           =   8
         Left            =   1920
         Picture         =   "frmChatForm.frx":1DF5
         ToolTipText     =   "Diablo"
         Top             =   1050
         Width           =   360
      End
      Begin VB.Image Image2 
         Height          =   360
         Index           =   7
         Left            =   1320
         Picture         =   "frmChatForm.frx":2163
         ToolTipText     =   "Lengua"
         Top             =   1065
         Width           =   360
      End
      Begin VB.Image Image2 
         Height          =   360
         Index           =   6
         Left            =   720
         Picture         =   "frmChatForm.frx":249A
         ToolTipText     =   "Llorar"
         Top             =   1065
         Width           =   360
      End
      Begin VB.Image Image2 
         Height          =   360
         Index           =   5
         Left            =   120
         Picture         =   "frmChatForm.frx":2806
         ToolTipText     =   "Flor"
         Top             =   1065
         Width           =   360
      End
      Begin VB.Image Image2 
         Height          =   360
         Index           =   4
         Left            =   2520
         Picture         =   "frmChatForm.frx":2AC9
         ToolTipText     =   "Muy Feliz"
         Top             =   360
         Width           =   360
      End
      Begin VB.Image Image2 
         Height          =   270
         Index           =   3
         Left            =   1920
         Picture         =   "frmChatForm.frx":2E15
         ToolTipText     =   "Beso"
         Top             =   450
         Width           =   270
      End
      Begin VB.Image Image2 
         Height          =   270
         Index           =   2
         Left            =   1320
         Picture         =   "frmChatForm.frx":3346
         ToolTipText     =   "Triste"
         Top             =   450
         Width           =   270
      End
      Begin VB.Image Image2 
         Height          =   270
         Index           =   1
         Left            =   600
         Picture         =   "frmChatForm.frx":373B
         ToolTipText     =   "Preocupado"
         Top             =   450
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   270
         Index           =   0
         Left            =   120
         Picture         =   "frmChatForm.frx":5074
         ToolTipText     =   "Feliz"
         Top             =   450
         Width           =   270
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   2675
         Picture         =   "frmChatForm.frx":552D
         ToolTipText     =   "Cerrar Cuador"
         Top             =   80
         Width           =   240
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00C0FFC0&
         FillColor       =   &H00A75021&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3000
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   3960
      Top             =   120
   End
   Begin MSComDlg.CommonDialog CmDialog1 
      Left            =   6480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Client_Trafic 
      Left            =   8160
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   3480
      Top             =   120
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4455
      Left            =   4080
      TabIndex        =   8
      Top             =   1200
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   7858
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmChatForm.frx":5F2F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2520
      Picture         =   "frmChatForm.frx":5FB3
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Eliminar Contacto"
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1560
      Picture         =   "frmChatForm.frx":687D
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Editar Contacto"
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      Picture         =   "frmChatForm.frx":7147
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Agregar Contacto"
      Top             =   5640
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChatForm.frx":7A11
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChatForm.frx":82EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChatForm.frx":8BC5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4695
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   8281
      _Version        =   393217
      LabelEdit       =   1
      Style           =   3
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock Client 
      Left            =   7680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Server 
      Index           =   0
      Left            =   7200
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7200
      Picture         =   "frmChatForm.frx":949F
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Enviar Mensaje"
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   7200
      Width           =   6735
   End
   Begin VB.PictureBox Picture2 
      Height          =   495
      Left            =   0
      Picture         =   "frmChatForm.frx":9D69
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   495
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   480
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   0
      Picture         =   "frmChatForm.frx":ABAB
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   270
      Left            =   6720
      Picture         =   "frmChatForm.frx":B875
      Top             =   6900
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Hablando Con:"
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
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   6960
      Width           =   6135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00800000&
      FillColor       =   &H00A75021&
      FillStyle       =   0  'Solid
      Height          =   365
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   6735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Messenger TCP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   4560
      TabIndex        =   10
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ultimos Mensajes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   840
      Width           =   3855
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00E49569&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   1695
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   6720
      Width           =   8055
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00E49569&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   6015
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   4575
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E49569&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C00000&
      Height          =   6015
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   3375
   End
   Begin VB.Menu Main1 
      Caption         =   "OptionsIcon"
      Visible         =   0   'False
      Begin VB.Menu OpenMessenger 
         Caption         =   "Abrir Messenger"
      End
      Begin VB.Menu ExitMessenger 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Main2 
      Caption         =   "ChatTextOption"
      Visible         =   0   'False
      Begin VB.Menu BorrarChat 
         Caption         =   "Borrar Texto"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuraciones"
      Visible         =   0   'False
      Begin VB.Menu mnuNombreTerminal 
         Caption         =   "Nombre de Terminal"
      End
      Begin VB.Menu MSGPersonal 
         Caption         =   "Mensaje de Inactividad"
      End
      Begin VB.Menu mnuAbaut 
         Caption         =   "Acerca de TCPMessenger"
      End
   End
End
Attribute VB_Name = "frmChatForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim DebeCerrar As Boolean
Dim EncriptarObject As New CRAEncryptClass
Dim DesencriptarObject As New CRADEEncryptClass
Dim ConectionTransfer As Integer
Dim EnEsperadeArchivo As Boolean
Dim ArchivoEsperado As String
Dim sizeFileTransmit As Long
Dim sizeFileR As Long
Dim FileReceive As Integer
Dim ContinuarWriteRichText As Boolean
Dim ActividadMessenger As Long
Dim RefreshContacts As Integer

Private Sub BorrarChat_Click()

   ActividadMessenger = 0
   RichTextBox1.Text = ""
   
   If Me.Visible = True Then
     Text1.SetFocus
   End If
End Sub

Private Sub Client_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next

Client.Close

End Sub

Private Sub Client_Trafic_Close()

ContinuarWriteRichText = True

End Sub

Private Sub Client_Trafic_DataArrival(ByVal bytesTotal As Long)
Dim DataReceived As String
Dim FlagColor As String
Dim Encabezado As String
Dim RichTemp As String

Client_Trafic.GetData DataReceived, vbString, bytesTotal

ActividadMessenger = 0

If Mid(DataReceived, 1, 8) = "CMD_MSG:" Then

'*****************************************************************************************************
        If InStr(1, RichTextBox1.TextRTF, "\cf1") <> 0 And _
           InStr(1, RichTextBox1.TextRTF, "\cf2") <> 0 Then
             Encabezado = "{\colortbl ;\red0\green128\blue128;\red0\green0\blue255;}"
              FlagColor = "\cf0"
        Else
         If InStr(1, RichTextBox1.TextRTF, "\cf1") <> 0 And _
           InStr(1, RichTextBox1.TextRTF, "\cf2") = 0 Then
             Encabezado = "{\colortbl ;\red0\green128\blue128;}"
             FlagColor = "\cf0"
         Else
            If InStr(1, RichTextBox1.TextRTF, "\cf1") = 0 And _
               InStr(1, RichTextBox1.TextRTF, "\cf2") <> 0 Then
                 Encabezado = "{\colortbl ;\red0\green0\blue255;}"
                 FlagColor = "\cf0"
            Else
               If InStr(1, RichTextBox1.TextRTF, "\cf1") = 0 And _
                  InStr(1, RichTextBox1.TextRTF, "\cf2") = 0 Then
                    Encabezado = ""
                    FlagColor = "\cf0"
               End If
            End If
         End If
        End If
        
        RichTemp = ""
        If RichTextBox1.Text <> "" Then
           RichTemp = Mid(RichTextBox1.TextRTF, Len("{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}} "), (Len(RichTextBox1.TextRTF) - Len("{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}} ")) - 2)
        End If
        
        RichTemp = RichTemp & "\viewkind4\uc1\pard\lang1033\f0\fs20" & FlagColor & "\b " & Mid(DataReceived, 9, Len(DataReceived) - 8) & "\b0\par"
        RichTemp = Replace(RichTemp, Mid(RichTemp, InStr(1, RichTemp, "{\colortbl"), InStr(InStr(1, RichTemp, "{\colortbl"), RichTemp, "}")), "")
        
        
        RichTextBox1.TextRTF = "{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}}" & Chr(13) & Encabezado & Chr(13) & RichTemp & Chr(13) & "\Par}"
        RichTextBox1.SelStart = Len(RichTextBox1.Text)
'******************************************************************************************************
 



   
Else
   
   DataReceived = ""
   
End If

End Sub



Private Sub Client_Trafic_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    MsgBox "Usuario se ha desconectado...", vbInformation, "Contacto Desconectado"
    Command1.Enabled = False
    ContinuarWriteRichText = True
    Client_Trafic.Close
    
    If Me.Visible = True Then
      Text1.SetFocus
    End If
ActividadMessenger = 0

End Sub

Private Sub Client_Trafic_SendComplete()
  
  ContinuarWriteRichText = True
ActividadMessenger = 0

End Sub

Private Sub Command1_Click()
On Error Resume Next

Dim MessageSendCliente As String
Dim i As Integer
Dim Bloque As String
Dim Segmentos As Integer
Dim FlagColor As String
Dim Encabezado As String
Dim RichTemp As String

ActividadMessenger = 0
Frame1.Visible = False
If Nulidad(Text1.Text) = False Then
   If Client_Trafic.State <> 7 Then
     
         MsgBox "Usuario se ha desconectado...", vbInformation, "Contacto Desconectado"
         Command1.Enabled = False
     
        If Me.Visible = True Then
           Text1.SetFocus
        End If
   Else
    Segmentos = 0
    For i = 1 To Len(Text1.Text)
    
        If Len(Bloque) = 200 Or _
           i = Len(Text1.Text) Then
           
              If i = Len(Text1.Text) Then
                Bloque = Bloque & Mid(Text1.Text, i, 1)
              End If
              Segmentos = Segmentos + 1
              
              If Not (Segmentos = 1 And i = Len(Text1.Text)) Then
                Bloque = "[msg:" & Segmentos & "] " & Bloque
              End If
              
              ContinuarWriteRichText = False
              
              MessageSendCliente = Server(0).LocalIP & "|" & Str(Server(0).LocalPort) & "|" & LocalName & "@" & EncriptarObject.EncryptStringToPlainCompress(Replace(Bloque, Chr(13), " \par ") & Chr(13) & Chr(13))
              Client_Trafic.SendData MessageSendCliente
              
              DoEvents
              
              Do While ContinuarWriteRichText = False
                DoEvents
              Loop
              
              ContinuarWriteRichText = True
              '****************************************************************************
                
                
                If InStr(1, RichTextBox1.TextRTF, "\cf1") <> 0 And _
                   InStr(1, RichTextBox1.TextRTF, "\cf2") <> 0 Then
                   Encabezado = "{\colortbl ;\red0\green128\blue128;\red0\green0\blue255;}"
                   FlagColor = "\cf1"
                Else
                 If InStr(1, RichTextBox1.TextRTF, "\cf1") <> 0 And _
                   InStr(1, RichTextBox1.TextRTF, "\cf2") = 0 Then
                    Encabezado = "{\colortbl ;\red0\green128\blue128;}"
                    FlagColor = "\cf1"
                 Else
                    If InStr(1, RichTextBox1.TextRTF, "\cf1") = 0 And _
                       InStr(1, RichTextBox1.TextRTF, "\cf2") <> 0 Then
                         Encabezado = "{\colortbl ;\red0\green0\blue255;}"
                         FlagColor = "\cf1"
                    Else
                       If InStr(1, RichTextBox1.TextRTF, "\cf1") = 0 And _
                          InStr(1, RichTextBox1.TextRTF, "\cf2") = 0 Then
                            Encabezado = "{\colortbl ;\red0\green128\blue128;}"
                            FlagColor = "\cf1"
                       End If
                    End If
                 End If
                End If
              
              RichTemp = ""
              If RichTextBox1.Text <> "" Then
                RichTemp = Mid(RichTextBox1.TextRTF, Len("{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}} "), (Len(RichTextBox1.TextRTF) - Len("{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}} ")) - 2)
              End If
              
              '************* Esto es lo nuevo *********************
              
              If LineaEmotIcon(UCase(Bloque)) = True Then

                   
                   RichTemp = RichTemp & Chr(13) & "\viewkind4\uc1\pard\lang1033\f0\fs20" & FlagColor & "\b [Yo a: " & UserSendMessage & "]\b0" & FlagColor & " : " & "" & Chr(13) & Chr(13) & "\par"
                   RichTemp = Replace(RichTemp, Mid(RichTemp, InStr(1, RichTemp, "{\colortbl"), InStr(InStr(1, RichTemp, "{\colortbl"), RichTemp, "}")), "")
                   RichTextBox1.TextRTF = "{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}}" & Chr(13) & Encabezado & Chr(13) & RichTemp & Chr(13) & "\par}"
                   
                   PintaEmotIcons RichTextBox1, Bloque
                   Bloque = ""
                   
              Else
                Bloque = LimpiaEmotIcon(Bloque)
                RichTemp = RichTemp & Chr(13) & "\viewkind4\uc1\pard\lang1033\f0\fs20" & FlagColor & "\b [Yo a: " & UserSendMessage & "]\b0" & FlagColor & " : " & Replace(Bloque, Chr(13), " \par ") & Chr(13) & Chr(13) & "\par"
                RichTemp = Replace(RichTemp, Mid(RichTemp, InStr(1, RichTemp, "{\colortbl"), InStr(InStr(1, RichTemp, "{\colortbl"), RichTemp, "}")), "")
                
                RichTextBox1.TextRTF = "{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}}" & Chr(13) & Encabezado & Chr(13) & RichTemp & Chr(13) & "\par}"
              
              End If
              
              '****************************************************
              
              RichTextBox1.SelStart = Len(RichTextBox1.Text)
              '****************************************************************************
 
              
              Bloque = ""
           
        Else
              Bloque = Bloque & Mid(Text1.Text, i, 1)
              DoEvents
        End If
    
    Next i
    
    Text1.Text = ""
    
    If Me.Visible = True Then
      Text1.SetFocus
    End If
   End If
Else
  MsgBox "Debe escribir un mensaje para que este sea enviado", vbInformation, "Mensaje en Blanco"
  If Me.Visible = True Then
    Text1.SetFocus
  End If
  
End If
ActividadMessenger = 0

End Sub


Private Sub Command2_Click()
   Frame1.Visible = False
   ActividadMessenger = 0
   EditOrNewContact NewContact, Command2
End Sub

Private Sub Command3_Click()
  ActividadMessenger = 0
  Frame1.Visible = False
  EditOrNewContact EditContact, Command3
End Sub

Private Sub Command4_Click()
Frame1.Visible = False
ActividadMessenger = 0
EliminaContactoINI Mid(frmChatForm.TreeView1.SelectedItem.Key, 1, InStr(1, frmChatForm.TreeView1.SelectedItem.Key, "-") - 1)

ActividadMessenger = 0

End Sub


Private Function NuevoSocket() As Integer
On Error GoTo Err_NuevoSocket

    Dim numElementos As Integer 'numero de sockets
    Dim i As Integer 'contador
    
    'obtiene la cantidad de Winsocks que tenemos
    numElementos = Server.UBound
    
    On Error Resume Next
    'recorre el arreglo de sockets
    For i = 0 To numElementos
        'si algun socket ya creado esta inactivo
        'utiliza este mismo para la nueva conexion
        If Server(i).State = sckClosed Then
            If Err.Number = 340 Then
               Load Server(i)
               NuevoSocket = i
               On Error GoTo 0
               Exit Function
            End If
        
            NuevoSocket = i 'retorna el indice
            Exit Function 'abandona la funcion
        End If
        
    Next
    
    'si no encuentra sockets inactivos
    'crea uno nuevo y devuelve su identidad
    Load Server(numElementos + 1) 'carga un nuevo socket al arreglo
    
    'devuelve el nuevo indice
    NuevoSocket = Server.UBound
    
Exit_Err_NuevoSocket:
Exit Function

Err_NuevoSocket:
Load Server(numElementos + 1)
NuevoSocket = numElementos + 1
Resume Exit_Err_NuevoSocket
End Function

Private Sub ExitMessenger_Click()

  DebeCerrar = True
  Unload Me
  
End Sub

Public Sub StartListenService(Index As Integer)
On Error Resume Next

    If Server(Index).State <> sckListening And _
       Server(Index).State <> sckConnected Then
       
        Server(Index).Close
        Server(Index).LocalPort = TCPListenPort
        Server(Index).Protocol = sckTCPProtocol
        Server(Index).Listen
        
    End If
    
End Sub

Private Sub Form_Activate()
On Error Resume Next

ActividadMessenger = 0
If TreeView1.Enabled = False Then
  Command2.SetFocus
Else
  TreeView1.SetFocus
End If

End Sub

Private Sub Form_Load()
On Error Resume Next


'********************   0        1         2      3       4        5      6       7       8       9      10      11       12     13
IconsPosibles = Array(FELIZ, PREOCUPADO, TRISTE, BESO, CARCAJADA, FLOR, LLORAR, LENGUA, DIABLO, ENOJO, ANGEL, SILENCIO, DORMIR, MULA)

    ActividadMessenger = 0
    InitConfValues
    StartListenService 0 'NuevoSocket
    Label2.Caption = Label2.Caption & Chr(13) & "v. " & App.Major & "." & App.Minor
    Me.Caption = Me.Caption & " - " & Server(0).LocalIP
    EnEsperadeArchivo = False
    ArchivoEsperado = ""
    
    RichTextBox1.Text = ""
    
    LoadContacts TreeView1, ImageList1
    Command1.Enabled = False
    ConectionTransfer = -1
    
    If TreeView1.Enabled = False Then
       Command3.Enabled = False
       Command4.Enabled = False
    Else
       Command3.Enabled = True
       Command4.Enabled = True
    End If
    
    DebeCerrar = False
    StartIconAPP 1, Picture1
     
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

ActividadMessenger = 0
If (x >= Image3.Left And x <= (Image3.Left + Image3.Width)) And _
   (Y >= Image3.Top And Y <= (Image3.Top + Image3.Height)) Then
   Image3.Visible = True
Else
   Image3.Visible = False
End If

If x >= 0 And x <= Me.Width And Y >= 0 And Y <= 100 Then
   mnuConfig.Visible = True
   Frame1.Visible = False
   Main1.Visible = False
   Main2.Visible = False
Else
   mnuConfig.Visible = False
   Main1.Visible = False
   Main2.Visible = False

End If

End Sub

Private Sub Form_Resize()
On Error Resume Next

ActividadMessenger = 0
If Me.Height < 9255 Then
  Me.Height = 9255
End If

If Me.Width < 8655 Then
  Me.Width = 8655
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

 If DebeCerrar = True Then
     CloseIconAPP
     End
  Else
     Cancel = -1
     Me.Hide
 End If
 
End Sub



Private Sub Image1_Click(Index As Integer)

     Frame1.Visible = False
     Text1.SetFocus

End Sub

Private Sub Image2_Click(Index As Integer)
    Text1.Text = Text1.Text & IconsPosibles(Index)
End Sub

Private Sub Image3_Click()

   Frame1.Visible = True
   
End Sub



Private Sub mnuAbaut_Click()
ActividadMessenger = 0
   frmAbaut.Show
   
End Sub

Private Sub mnuNombreTerminal_Click()
On Error Resume Next
Dim NewUser As String

ActividadMessenger = 0
   NewUser = InputBox("El nombre de terminal es el que aparecerá " & Chr(13) & _
                      "a otro usuario cuando usted intente contactarlo." & Chr(13) & _
                      "Ingrese Nuevo Nombre de su Terminal: ", "Cambio de Nombre de Terminal", LocalName)
   
   If Nulidad(NewUser) = True Or NewUser = LocalName Then
     MsgBox "Cambio de Nombre de terminal Cancelado", vbInformation, "Acción Cancelada"
   Else
     WriteStringINI App.Path & "\" & App.EXEName & ".ini", "Service", "LocalName", NewUser
     LocalName = NewUser
   End If
   
mnuConfig.Visible = False
ActividadMessenger = 0

End Sub

Private Sub MSGPersonal_Click()
On Error Resume Next
Dim NewMessage As String
Dim Confirmacion As VbMsgBoxResult

ActividadMessenger = 0
   
   NewMessage = MSGMensajeInactividad
   
   NewMessage = InputBox("Puede ingresar un mensaje personalizado el cual" & Chr(13) & _
                      "al momento que usted este en inactividad sera mostrado." & Chr(13) & _
                      "Ingrese Mensaje Personalizado: ", "Mensaje Personalizado", MSGMensajeInactividad)
   
   If Trim(NewMessage) = "" Then
    Confirmacion = MsgBox("El valor del nuevo mensaje esta vacio, quiere dejar el mensaje anterior?", vbQuestion + vbOKCancel, "Mensaje Personalizado")
    
    If Confirmacion = vbOK Then
       NewMessage = MSGMensajeInactividad
    Else
       NewMessage = ""
    End If
   End If
   
      WriteStringINI App.Path & "\" & App.EXEName & ".ini", "Service", "MsgInactividad", "" & NewMessage & ""
      MSGMensajeInactividad = NewMessage

   
mnuConfig.Visible = False
ActividadMessenger = 0
End Sub

Private Sub OpenMessenger_Click()
ActividadMessenger = 0
If Me.Visible = False Then
  Me.Show
  If NotificaPantalla.hIcon = Picture2.Picture Then
    
    CloseIconAPP
    StartIconAPP 1, Picture1
 
  End If
  Text1.SetFocus
  
End If

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

  ActividadMessenger = 0
  PictureMoveMouse Me, Button, Shift, x, Y
  
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

  ActividadMessenger = 0
  PictureMoveMouse Me, Button, Shift, x, Y
  
End Sub



Private Sub RichTextBox1_DblClick()
ActividadMessenger = 0
Me.PopupMenu Main2

End Sub

Private Sub RichTextBox1_GotFocus()
Frame1.Visible = False

ActividadMessenger = 0
If NotificaPantalla.hIcon = Picture2.Picture Then
 CloseIconAPP
 StartIconAPP 1, Picture1
End If

End Sub

Private Sub RichTextBox1_LostFocus()

ActividadMessenger = 0
If NotificaPantalla.hIcon = Picture2.Picture Then
 CloseIconAPP
 StartIconAPP 1, Picture1
End If

End Sub

Private Sub RichTextBox1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

ActividadMessenger = 0

End Sub

Private Sub Server_Close(Index As Integer)
On Error Resume Next

   If Index <> 0 Then
     Unload Server(Index)
   End If

End Sub

Private Sub Server_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error GoTo err_Server_ConnectionRequest
Dim NuevoIndice As Integer

NuevoIndice = NuevoSocket

Server(NuevoIndice).Accept requestID




exit_Server_ConnectionRequest:
Exit Sub

err_Server_ConnectionRequest:
MsgBox "Se Detecto Error:" & Chr(13) & Err.Description, vbInformation, "Información..."
Resume exit_Server_ConnectionRequest

End Sub

Private Sub Server_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim DataReceived As String
Dim EntryIP As String
Dim EntryPort As String
Dim ContactName As String
Dim ContactSend As String
Dim CadTransferFile As String
Dim Archivo() As Byte
Dim RichTemp As String
Dim Encabezado As String
Dim FlagColor As String
Dim CuerpoMensaje As String

DataReceived = ""

If EnEsperadeArchivo = True And _
   Index = ConectionTransfer Then
  
  sizeFileR = sizeFileR + bytesTotal
  
  Server(Index).GetData Archivo
  
  Put #FileReceive, , Archivo

  If sizeFileR >= sizeFileTransmit Then
    Close #FileReceive
    FileReceive = 0
    EnEsperadeArchivo = False
    Server(Index).SendData "OK" & vbCrLf
  End If
  
  Exit Sub

End If

Server(Index).GetData DataReceived, vbString, bytesTotal


EntryIP = Mid(DataReceived, 1, InStr(1, DataReceived, "|") - 1)
EntryPort = Val(Mid(DataReceived, InStr(1, DataReceived, "|") + 1, InStr(InStr(1, DataReceived, "|") + 1, DataReceived, "|") - InStr(1, DataReceived, "|") - 1))
ContactSend = Mid(DataReceived, InStr(InStr(1, DataReceived, "|") + 1, DataReceived, "|") + 1, InStr(1, DataReceived, "@") - InStr(InStr(1, DataReceived, "|") + 1, DataReceived, "|") - 1)
'Validar si el usuario es valido para que se comunique.....
If ValidaExisteContactoSilen(EntryIP, EntryPort, ContactName) = True Then

'Buscar nombre del usuario luego de validarlo como usuario valido.
 If Mid(Mid(DataReceived, InStr(1, DataReceived, "@") + 1, Len(DataReceived) - InStr(1, DataReceived, "@")), 1, 3) = "TRF" Or _
    Mid(Mid(DataReceived, InStr(1, DataReceived, "@") + 1, Len(DataReceived) - InStr(1, DataReceived, "@")), 1, 3) = "TRT" Or _
    Mid(Mid(DataReceived, InStr(1, DataReceived, "@") + 1, Len(DataReceived) - InStr(1, DataReceived, "@")), 1, 3) = "TRE" Then
    
    If Me.Visible = False Then
      Me.Show
    End If
    
   CadTransferFile = ""
   CadTransferFile = Mid(DataReceived, InStr(1, DataReceived, "@") + 1, Len(DataReceived) - (InStr(1, DataReceived, "@") + 2))
   
   If Mid(CadTransferFile, 1, 3) = "TRF" Then
   
         If EnEsperadeArchivo = True And _
            ArchivoEsperado <> "" And _
            sizeFileTransmit <> 0 Then
            
            Server(Index).SendData "Usuario esta recibiendo otro archivo... intente luego" & vbCrLf
            Exit Sub
         
         End If
   
         If ConectionTransfer > -1 Then
             If Server(ConectionTransfer).State <> 0 Then
                Server(ConectionTransfer).Close
             End If
         End If
         
         On Error Resume Next
         ConectionTransfer = Index
         CmDialog1.CancelError = True
         CmDialog1.DialogTitle = "Guardar Archivo Transferido en: "
         CmDialog1.FileName = Mid(CadTransferFile, 5, Len(CadTransferFile) - 4)
         CmDialog1.Filter = "Receibe File (*.*)|*.*"
         CmDialog1.Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly
         CmDialog1.ShowSave
         
         
         
         If Err.Number = cdlCancel Then
            
            If Server(ConectionTransfer).State = 7 Then
              Server(ConectionTransfer).SendData "Acción Rechazada..." & vbCrLf
            End If
            ConectionTransfer = -1
            Exit Sub
         End If
         
         Server(ConectionTransfer).SendData "OK" & vbCrLf
         
         On Error GoTo 0
         
   Else
      If Mid(CadTransferFile, 1, 3) = "TRT" Then
         
         
         EnEsperadeArchivo = True
         ArchivoEsperado = CmDialog1.FileName
         sizeFileTransmit = Val(Mid(CadTransferFile, 5, Len(CadTransferFile) - 4))
         
         If FileReceive <> 0 Then
           Close #FileReceive
         End If
         
         FileReceive = FreeFile
  
         Open ArchivoEsperado For Binary Access Write As #FileReceive
         Server(ConectionTransfer).SendData "OK" & vbCrLf
         
      Else
         If Mid(CadTransferFile, 1, 3) = "TRE" Then
           Server(ConectionTransfer).SendData "OK" & vbCrLf
           
           frmSplash.Label1 = Mid(CadTransferFile, 5, Len(CadTransferFile) - 4) & Chr(13) & _
                              "Bytes Transmitidos: " & Format(sizeFileTransmit, "#,###,###,###,##0")

           ArchivoEsperado = ""
           sizeFileTransmit = 0
           sizeFileR = 0
           frmSplash.Show
           If Server(0).State <> sckListening Then
              If Server(0).State <> sckClosed Then
                Server(0).Close
              End If
              StartListenService 0
           End If
           ConectionTransfer = -1
         End If
      End If
   End If
 
 Else
        
'*****************************************************************************************************
        
        
        If InStr(1, RichTextBox1.TextRTF, "\cf1") <> 0 And _
           InStr(1, RichTextBox1.TextRTF, "\cf2") <> 0 Then
           Encabezado = "{\colortbl ;\red0\green128\blue128;\red0\green0\blue255;}"
           FlagColor = "\cf2"
        Else
         If InStr(1, RichTextBox1.TextRTF, "\cf1") <> 0 And _
           InStr(1, RichTextBox1.TextRTF, "\cf2") = 0 Then
            Encabezado = "{\colortbl ;\red0\green128\blue128;\red0\green0\blue255;}"
              FlagColor = "\cf2"
         Else
            If InStr(1, RichTextBox1.TextRTF, "\cf1") = 0 And _
               InStr(1, RichTextBox1.TextRTF, "\cf2") <> 0 Then
                 Encabezado = "{\colortbl ;\red0\green0\blue255;}"
                 FlagColor = "\cf1"
            Else
               If InStr(1, RichTextBox1.TextRTF, "\cf1") = 0 And _
                  InStr(1, RichTextBox1.TextRTF, "\cf2") = 0 Then
                    Encabezado = "{\colortbl ;\red0\green0\blue255;}"
                    FlagColor = "\cf1"
               End If
            End If
         End If
        End If

        
        RichTemp = ""
        If RichTextBox1.Text <> "" Then
           RichTemp = Mid(RichTextBox1.TextRTF, Len("{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}} "), (Len(RichTextBox1.TextRTF) - Len("{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}} ")) - 2)
        End If
        
        CuerpoMensaje = ""
        CuerpoMensaje = Replace(Trim(DesencriptarObject.DEncryptStringToPlainCompress(Mid(DataReceived, InStr(1, DataReceived, "@") + 1, Len(DataReceived) - InStr(1, DataReceived, "@")))), Chr(13), "")
        

        
        If LineaEmotIcon(UCase(CuerpoMensaje)) = True Then
        
                   RichTemp = RichTemp & "\viewkind4\uc1\pard\lang1033\f0\fs20" & FlagColor & "\b [" & ContactName & "]\b0" & FlagColor & " :  " & "" & "\par"
                   RichTemp = Replace(RichTemp, Mid(RichTemp, InStr(1, RichTemp, "{\colortbl"), InStr(InStr(1, RichTemp, "{\colortbl"), RichTemp, "}")), "")
                   RichTextBox1.TextRTF = "{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}}" & Chr(13) & Encabezado & Chr(13) & RichTemp & Chr(13) & "\Par}"
                   PintaEmotIcons RichTextBox1, CuerpoMensaje
                   CuerpoMensaje = ""
                   

        Else
        
            CuerpoMensaje = LimpiaEmotIcon(CuerpoMensaje)
            RichTemp = RichTemp & "\viewkind4\uc1\pard\lang1033\f0\fs20" & FlagColor & "\b [" & ContactName & "]\b0" & FlagColor & " :  " & CuerpoMensaje & "\par"
            RichTemp = Replace(RichTemp, Mid(RichTemp, InStr(1, RichTemp, "{\colortbl"), InStr(InStr(1, RichTemp, "{\colortbl"), RichTemp, "}")), "")
            RichTextBox1.TextRTF = "{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}}" & Chr(13) & Encabezado & Chr(13) & RichTemp & Chr(13) & "\Par}"
        End If
        
        RichTextBox1.SelStart = Len(RichTextBox1.Text)
'******************************************************************************************************
        If ActividadMessenger >= 300 Then
          If MSGMensajeInactividad = "" Then
           Server(Index).SendData "CMD_MSG:El usuario tiene mas de 30 Segundos de Inactividad" & Chr(13)
          Else
           Server(Index).SendData "CMD_MSG:" & MSGMensajeInactividad & Chr(13)
          End If
        End If
        
        If ActividadMessenger >= 300 And _
           frmChatForm.Visible = True Then
           
           CloseIconAPP
           StartIconAPP EnterMessage, Picture2
           frmSplash.Label1 = "Usted Tiene nuevos Mensajes de: " & Chr(13) & ContactName
           frmSplash.Show
           
        End If
        
        If frmChatForm.Visible = False Then
           
           CloseIconAPP
           StartIconAPP EnterMessage, Picture2
           frmSplash.Label1 = "Usted Tiene nuevos Mensajes de: " & Chr(13) & ContactName
           frmSplash.Show
        Else
        
          frmChatForm.SetFocus
          Text1.SetFocus
          
        End If
 End If
Else
    
    
    Me.Show
    
    
    If MsgBox("La Terminal: <" & EntryIP & "> de:  " & ContactSend & Chr(13) & _
              "Desea agregar a su lista de contactos? ", vbOKCancel + vbQuestion, "Solicitud de Acceso") = vbOK Then
              
      DlgContacts.Text1.Text = ContactSend
      DlgContacts.Text2.Text = EntryIP
      DlgContacts.Text3.Text = EntryPort
    
      EditOrNewContact NewContact, Command2
         Server(Index).SendData "CMD_MSG:El usuario agregará su usuario, Espere mientras se comunica..." & Chr(13)
    Else
    
         Server(Index).SendData "CMD_MSG:Acción Rechazada..." & Chr(13)
    End If


End If



End Sub

Private Sub Server_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next

   Server(Index).Close
   
End Sub

Private Sub Text1_Click()
   
   Frame1.Visible = False
   
End Sub

Private Sub Text1_GotFocus()

Frame1.Visible = False

ActividadMessenger = 0
 If NotificaPantalla.hIcon = Picture2.Picture Then
    CloseIconAPP
    StartIconAPP 1, Picture1
 End If
 
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Frame1.Visible = False
ActividadMessenger = 0

End Sub

Private Sub Timer1_Timer()
DoEvents
RefreshContacts = RefreshContacts + 1
If RefreshContacts = 12 Then
    RefreshContacts = 0
    ValidateContactConect TreeView1, Client
End If
DoEvents
End Sub

Private Sub Timer2_Timer()

DoEvents
If Server(0).State <> sckListening Then
    If Server(0).State <> sckClosed Then
       Server(0).Close
    End If
   StartListenService 0
End If
DoEvents

End Sub

Private Sub Timer3_Timer()

ActividadMessenger = ActividadMessenger + 1
DoEvents
If ActividadMessenger >= 86400 Then
  DoEvents
  ActividadMessenger = 300
  
End If

End Sub

Private Sub TreeView1_DblClick()
On Error Resume Next
DoEvents
ActividadMessenger = 0
  If TreeView1.SelectedItem.Image = 2 And _
     TreeView1.SelectedItem.SelectedImage = 2 Then
  
  Else
   With frmSplash1
   .UserSendFile = TreeView1.SelectedItem.Text
   .IpTransferFile = Mid(TreeView1.SelectedItem.Key, InStr(1, TreeView1.SelectedItem.Key, "-") + 1, _
                    InStr(1, TreeView1.SelectedItem.Key, "|") - InStr(1, TreeView1.SelectedItem.Key, "-") - 1)
   .PortTransferFile = Mid(TreeView1.SelectedItem.Key, InStr(1, TreeView1.SelectedItem.Key, "|") + 1, Len(TreeView1.SelectedItem.Key) - InStr(1, TreeView1.SelectedItem.Key, "|"))
   
   
    CmDialog1.CancelError = True
    CmDialog1.DialogTitle = "Archivo a Transferir"
    CmDialog1.FileName = ""
    CmDialog1.Filter = "Send File (*.*)|*.*"
    CmDialog1.FilterIndex = 1
    CmDialog1.Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly
    CmDialog1.ShowOpen
    
    If Err = cdlCancel Then
       .FileTransfer = vbNullString
       .UserSendFile = ""
       .IpTransferFile = ""
       .PortTransferFile = ""
       On Error GoTo 0
       Exit Sub
    Else
      .FileTransfer = CmDialog1.FileName
      .F_Name = CmDialog1.FileTitle
      On Error GoTo 0
    End If
    DoEvents
   .Show
   End With
  End If
  
End Sub

Private Sub TreeView1_GotFocus()
Frame1.Visible = False
End Sub

Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

ActividadMessenger = 0

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo Err_TreeView1_NodeClick
Dim j As Integer
Dim i As Integer

ActividadMessenger = 0
   
   UserSendMessage = TreeView1.SelectedItem.Text
   IPuserSelected = Mid(TreeView1.SelectedItem.Key, InStr(1, TreeView1.SelectedItem.Key, "-") + 1, _
                    InStr(1, TreeView1.SelectedItem.Key, "|") - InStr(1, TreeView1.SelectedItem.Key, "-") - 1)
   PortUserSelected = Mid(TreeView1.SelectedItem.Key, InStr(1, TreeView1.SelectedItem.Key, "|") + 1, Len(TreeView1.SelectedItem.Key) - InStr(1, TreeView1.SelectedItem.Key, "|"))
   
   If Client_Trafic.State <> 0 Then
      Client_Trafic.Close
      For i = 1 To 50
        DoEvents
         If Client_Trafic.State = 0 Then
            Exit For
         End If
        DoEvents
        DoEvents
      Next i
   End If
   
   Client_Trafic.RemoteHost = IPuserSelected
   Client_Trafic.RemotePort = PortUserSelected
   Client_Trafic.Connect
   
   For j = 1 To 50
    DoEvents
    DoEvents
    DoEvents
   Next j
   
   If Client_Trafic.State = 7 Then
      
      Label4.Caption = "Hablando Con: " & TreeView1.SelectedItem.Text
      TreeView1.SelectedItem.Image = 1
      TreeView1.SelectedItem.SelectedImage = 1
      Command1.Enabled = True
   
   Else
   
        TreeView1.SelectedItem.Image = 2
        TreeView1.SelectedItem.SelectedImage = 2
        Label4.Caption = "Hablando Con: "
        
        If Client_Trafic.State <> 0 Then
           Client_Trafic.Close
           DoEvents
           DoEvents
        End If
        
        UserSendMessage = ""
        IPuserSelected = ""
        PortUserSelected = ""
        Command1.Enabled = False
   
   End If
Exit_Err_TreeView1_NodeClick:
Exit Sub

Err_TreeView1_NodeClick:
MsgBox Err.Number & ": " & Chr(13) & Err.Description, vbInformation, "Error al dar click en contacto"
Resume Exit_Err_TreeView1_NodeClick
End Sub
