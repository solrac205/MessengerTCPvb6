VERSION 5.00
Begin VB.Form DlgContacts 
   BackColor       =   &H00A75021&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   4380
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2910
      Picture         =   "DlgContacts.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1470
      Picture         =   "DlgContacts.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text3 
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
      Left            =   1440
      TabIndex        =   5
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text2 
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
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text1 
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
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Puerto:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "IP:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contacto:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E49569&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      Height          =   3855
      Left            =   270
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "DlgContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public ActionForm As TypeActionContact
Public IniItem As String

Public ContactNameIn As String
Public IpContactIn As String
Public PortContactIn As String


Public Function ValidaDatos() As Boolean
On Error Resume Next
ValidaDatos = False

If Nulidad(Text1.Text) = False And _
   Nulidad(Text2.Text) = False And _
   Nulidad(Text3.Text) = False Then
   
      If CompruebaIp(Text2.Text) = False Then
         MsgBox "IP ingresada es incorrecta", vbInformation, "Error validando datos"
         Text2.SetFocus
      Else
         Text2.Text = Val(Mid(Text2.Text, 1, InStr(1, Text2.Text, ".") - 1)) & "." & _
                       Val(Mid(Text2.Text, InStr(1, Text2.Text, ".") + 1, (InStr(InStr(1, Text2.Text, ".") + 1, Text2.Text, ".") - InStr(1, Text2.Text, ".")) - 1)) & "." & _
                       Val(Mid(Text2.Text, InStr(InStr(1, Text2.Text, ".") + 1, Text2.Text, ".") + 1, (InStr(InStr(InStr(1, Text2.Text, ".") + 1, Text2.Text, ".") + 1, Text2.Text, ".") - InStr(InStr(1, Text2.Text, ".") + 1, Text2.Text, ".")) - 1)) & "." & _
                       Val(Mid(Text2.Text, InStr(InStr(InStr(1, Text2.Text, ".") + 1, Text2.Text, ".") + 1, Text2.Text, ".") + 1, Len(Text2.Text) - InStr(InStr(InStr(1, Text2.Text, ".") + 1, Text2.Text, ".") + 1, Text2.Text, ".")))
       
         If IsNumeric(Text3.Text) = True And _
            (Val(Text3.Text) - Int(Val(Text3.Text))) = 0 Then
               ValidaDatos = True
         Else
           MsgBox "Puerto invalido, favor revisar debe ser un numero entero", vbInformation, "Error validando datos"
           Text3.SetFocus
         End If
      End If
   
Else
  MsgBox "Debe agregar valores en todos los campos, valores en blanco o nulos no permitidos", vbInformation, "Error validando datos"
  Text1.SetFocus
End If

End Function


Private Sub Command1_Click()
On Error Resume Next

  If ValidaDatos = True Then
     If ActionForm = NewContact Then
     
       If ValidaExisteContacto(Text2.Text, Text3.Text) = False Then
        ContactCount = ContactCount + 1
        WriteStringINI App.Path & "\" & App.EXEName & ".ini", "Contacts", "ContactName" & ContactCount, Text1.Text
        WriteStringINI App.Path & "\" & App.EXEName & ".ini", "Contacts", "IpContact" & ContactCount, Text2.Text
        WriteStringINI App.Path & "\" & App.EXEName & ".ini", "Contacts", "PortContact" & ContactCount, Text3.Text
        WriteStringINI App.Path & "\" & App.EXEName & ".ini", "Service", "ContactCount", Str(ContactCount)
       End If
      
        Unload Me
        
     Else
     
       If Text1.Text = ContactNameIn And _
          Text2.Text = IpContactIn And _
          Text3.Text = PortContactIn Then
          
         Unload Me
         
       Else
          If Text2.Text = IpContactIn And _
             Text3.Text = PortContactIn Then
               
               WriteStringINI App.Path & "\" & App.EXEName & ".ini", "Contacts", "ContactName" & IniItem, Text1.Text
               
          Else
            If ValidaExisteContacto(Text2.Text, Text3.Text) = False Then
            
                WriteStringINI App.Path & "\" & App.EXEName & ".ini", "Contacts", "ContactName" & IniItem, Text1.Text
                WriteStringINI App.Path & "\" & App.EXEName & ".ini", "Contacts", "IpContact" & IniItem, Text2.Text
                WriteStringINI App.Path & "\" & App.EXEName & ".ini", "Contacts", "PortContact" & IniItem, Text3.Text
            
            
            End If
         End If
         
         Unload Me
       End If
       
     End If
  End If
  
End Sub

Private Sub Command2_Click()
On Error Resume Next

Unload Me

End Sub

