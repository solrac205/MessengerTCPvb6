Attribute VB_Name = "ChatMessengerTools"
Public Type TIPONOTIFICARICONO
  cbSize As Long
  hWnd As Long
  uId As Long
  uFlags As Long
  ucallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type


'***************************************************************************************************
'Constantes de emoticons
'Si se agregan emoticons hay que agregarlos aca y cambiar el valor de la constante NumMaxIconsIndex
'***************************************************************************************************
Public Const FELIZ = "<:)>"               '0
Public Const PREOCUPADO = "<:S>"          '1
Public Const TRISTE = "<:(>"              '2
Public Const BESO = "<:*>"                '3
Public Const CARCAJADA = "<:D>"           '4
Public Const FLOR = "<:F>"                '5
Public Const LLORAR = "<;(>"              '6
Public Const LENGUA = "<:P>"              '7
Public Const DIABLO = "<:6>"              '8
Public Const ENOJO = "<:#>"               '9
Public Const ANGEL = "<:A>"               '10
Public Const SILENCIO = "<:-)>"           '11
Public Const DORMIR = "<(SLEEP)>"         '12
Public Const MULA = "<:M>"                '13

Public Const NumMaxIconsIndex As Integer = 13
Public IconsPosibles As Variant
'***************************************************************************************************


Public ContenidoClip As String



Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function Shell_NotifyIcon Lib "shell32" _
    Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
    pnid As TIPONOTIFICARICONO) As Boolean
    
Public Declare Function WinExec& Lib "kernel32" _
    (ByVal lpCmdLine As String, ByVal nCmdShow As Long)
'--------------------

Public Enum TypeActionContact
   NewContact = 1
   EditContact = 2
End Enum

Public Enum TypeObjectSystem
   craDirectory = 1
   craFile = 2
End Enum

Public Enum ActionMessenger
  OpenMessenger = 1
  EnterMessage = 2
End Enum

Public NotificaPantalla As TIPONOTIFICARICONO
Private UltimaImagen As PictureBox


Private Declare Function PathFileExistsW Lib "shlwapi.dll" ( _
    ByVal pszPath As Long) As Boolean
'// Declaración de API para verificación de existencia de Directorios.

Private Declare Function SHCreateDirectory Lib "shell32" ( _
    ByVal hWnd As Long, _
    ByVal pszPath As Long) As Long
'// Declaración de API para creación de Estructuras Completas de Directorios


Public UserSendMessage As String
Public IPuserSelected As String
Public PortUserSelected As String



'Declaración del Api SetLayeredWindowAttributes que establece _
 la transparencia al form
  
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
                (ByVal hWnd As Long, _
                 ByVal crKey As Long, _
                 ByVal bAlpha As Byte, _
                 ByVal dwFlags As Long) As Long
  
  
'Recupera el estilo de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                (ByVal hWnd As Long, _
                 ByVal nIndex As Long) As Long
  
  
'Declaración del Api SetWindowLong necesaria para aplicar un estilo _
 al form antes de usar el Api SetLayeredWindowAttributes
  
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
  
  
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
'Función para saber si formulario ya es transparente. _
 Se le pasa el Hwnd del formulario en cuestión
  
Public Function Is_Transparent(ByVal hWnd As Long) As Boolean
On Error Resume Next
  
Dim Msg As Long
  
    Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
          
       If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
          Is_Transparent = True
       Else
          Is_Transparent = False
       End If
  
    If Err Then
       Is_Transparent = False
    End If
  
End Function
  
'Función que aplica la transparencia, se le pasa el hwnd del form y un valor de 0 a 255
Public Function Aplicar_Transparencia(ByVal hWnd As Long, _
                                      Valor As Integer) As Long
  
Dim Msg As Long
  
On Error Resume Next
  
If Valor < 0 Or Valor > 255 Then
   Aplicar_Transparencia = 1
Else
   Msg = GetWindowLong(hWnd, GWL_EXSTYLE)
   Msg = Msg Or WS_EX_LAYERED
      
   SetWindowLong hWnd, GWL_EXSTYLE, Msg
      
   'Establece la transparencia
   SetLayeredWindowAttributes hWnd, 0, Valor, LWA_ALPHA
  
   Aplicar_Transparencia = 0
  
End If
  
  
If Err Then
   Aplicar_Transparencia = 2
End If
  
End Function





Public Function DeleteObjectSystem(PathAndObjectDelete As String, TypeDelete As TypeObjectSystem) As Boolean
On Error GoTo Err_DeleteObjectSystem

Dim Fs As Object

If PathAndObjectDelete = "" Or IsNull(PathAndObjectDelete) Then
  MsgBox "Valor de Path / Objeto es requerido", vbCritical, "Error en función DeleteObjectSystem"
  DeleteObjectSystem = False
  Exit Function
End If


Set Fs = CreateObject("Scripting.FileSystemObject")

Select Case TypeDelete
Case 1
   If PathFileExistsW(StrPtr(PathAndObjectDelete)) Then
      Fs.DeleteFolder (PathAndObjectDelete)
      DeleteObjectSystem = True
      
   Else
      MsgBox "Directorio no existe, no pudo ser eliminado", vbInformation, "Error en función DeleteObjectSystem"
      DeleteObjectSystem = False
      Set Fs = Nothing
      Exit Function
   End If
Case 2
   If Fs.FileExists(PathAndObjectDelete) Then
   
    Fs.DeleteFile PathAndObjectDelete, True
    DeleteObjectSystem = True
    
    
   Else
   
     MsgBox "Archivo no existe, no pudo ser eliminado", vbInformation, "Error en función DeleteObjectSystem"
     DeleteObjectSystem = False
     Set Fs = Nothing
     Exit Function
   
   End If

Case Else

   MsgBox "Opción no Valida, Verificar TypeObjectSystem seleccionado", vbInformation, "Error en función DeleteObjectSystem"
   DeleteObjectSystem = False
   Set Fs = Nothing
   Exit Function
End Select

Set Fs = Nothing

Exit_Err_DeleteObjectSystem:
Exit Function

Err_DeleteObjectSystem:
MsgBox Err.Description, vbInformation, "Error en función DeleteObjectSystem"
DeleteObjectSystem = False
Set Fs = Nothing
Resume Exit_Err_DeleteObjectSystem
End Function




Public Sub EliminaContactoINI(ItemIniContact As String)
On Error Resume Next

Dim FileIniOpenScan As Integer
Dim NewIniEdited As Integer
Dim ReadFile As String
Dim RegLocate As Boolean

RegLocate = False

FileIniOpenScan = FreeFile
Open App.Path & "\" & App.EXEName & ".ini" For Input As #FileIniOpenScan

NewIniEdited = FreeFile
Open App.Path & "\" & App.EXEName & "2.ini" For Output As #NewIniEdited

Do While Not EOF(FileIniOpenScan)
Line Input #FileIniOpenScan, ReadFile

If InStr(1, ReadFile, "ContactName" & ItemIniContact) > 0 Or _
   InStr(1, ReadFile, "IpContact" & ItemIniContact) > 0 Or _
   InStr(1, ReadFile, "PortContact" & ItemIniContact) > 0 Then
   
   RegLocate = True
    
End If

If RegLocate = False Then
   Print #NewIniEdited, ReadFile
Else
 If Not (InStr(1, ReadFile, "ContactName" & ItemIniContact) > 0 Or _
    InStr(1, ReadFile, "IpContact" & ItemIniContact) > 0 Or _
    InStr(1, ReadFile, "PortContact" & ItemIniContact) > 0) Then
   
    If Mid(ReadFile, 1, 11) = "ContactName" Then
      Print #NewIniEdited, Mid(ReadFile, 1, 11) & (Val(Mid(ReadFile, 12, InStr(12, ReadFile, "=") - 12)) - 1) & Mid(ReadFile, InStr(12, ReadFile, "="), Len(ReadFile) - InStr(12, ReadFile, "=") + 1)
    Else
      If Mid(ReadFile, 1, 9) = "IpContact" Then
           Print #NewIniEdited, Mid(ReadFile, 1, 9) & (Val(Mid(ReadFile, 10, InStr(10, ReadFile, "=") - 10)) - 1) & Mid(ReadFile, InStr(10, ReadFile, "="), Len(ReadFile) - InStr(10, ReadFile, "=") + 1)
      Else
        If Mid(ReadFile, 1, 11) = "PortContact" Then
           Print #NewIniEdited, Mid(ReadFile, 1, 11) & (Val(Mid(ReadFile, 12, InStr(12, ReadFile, "=") - 12)) - 1) & Mid(ReadFile, InStr(12, ReadFile, "="), Len(ReadFile) - InStr(12, ReadFile, "=") + 1)
        Else
           Print #NewIniEdited, ReadFile
        End If
      End If
    End If
      
 End If
   
End If

Loop

Close #NewIniEdited
Close #FileIniOpenScan

DeleteObjectSystem App.Path & "\" & App.EXEName & ".ini", craFile
Name App.Path & "\" & App.EXEName & "2.ini" As App.Path & "\" & App.EXEName & ".ini"

ContactCount = ContactCount - 1
WriteStringINI App.Path & "\" & App.EXEName & ".ini", "Service", "ContactCount", Str(ContactCount)

LoadContacts frmChatForm.TreeView1, frmChatForm.ImageList1

With frmChatForm
If .TreeView1.Enabled = False Then
       .Command3.Enabled = False
       .Command4.Enabled = False
Else
       .Command3.Enabled = True
       .Command4.Enabled = True
End If
.Label4.Caption = "Hablando Con:"
End With


End Sub


Public Function Nulidad(InputString As String) As Boolean
On Error GoTo Err_Nulidad
Dim i As Integer

If Len(InputString) = 0 Or _
   InputString = "" Or _
   IsNull(InputString) Then
   
  Nulidad = True
Else
   For i = 1 To Len(InputString)
       If Mid(InputString, i, 1) <> " " Then
           Nulidad = False
           Exit For
        Else
           Nulidad = True
       End If
   Next i
End If
Exit_Err_Nulidad:
Exit Function
Err_Nulidad:
MsgBox Err.Description, vbExclamation, "Error en DataComp"
Resume Exit_Err_Nulidad
End Function

Public Function ValidaExisteContactoSilen(IPContact As String, PortContact As String, ContactNameFun As String) As Boolean
On Error Resume Next
Dim ContactNameIni As String
Dim IPContactIni As String
Dim PortContactIni As String
Dim i

ValidaExisteContactoSilen = False

For i = 1 To ContactCount
    ContactNameIni = ""
    IPContactIni = ""
    PortContactIni = ""
    
    ContactNameIni = ReadStringINI(App.Path & "\" & App.EXEName & ".ini", "Contacts", "ContactName" & i, "")
    IPContactIni = ReadStringINI(App.Path & "\" & App.EXEName & ".ini", "Contacts", "IpContact" & i, "")
    PortContactIni = ReadStringINI(App.Path & "\" & App.EXEName & ".ini", "Contacts", "PortContact" & i, "")
    
    If IPContactIni = IPContact And _
       PortContactIni = PortContact Then
       ContactNameFun = ContactNameIni
       ValidaExisteContactoSilen = True
       Exit For
    End If
    
Next i



End Function

Public Function ValidaExisteContacto(IPContact As String, PortContact As String) As Boolean
On Error Resume Next
Dim ContactNameIni As String
Dim IPContactIni As String
Dim PortContactIni As String
Dim i

ValidaExisteContacto = False

For i = 1 To ContactCount
    ContactNameIni = ""
    IPContactIni = ""
    PortContactIni = ""
    
    ContactNameIni = ReadStringINI(App.Path & "\" & App.EXEName & ".ini", "Contacts", "ContactName" & i, "")
    IPContactIni = ReadStringINI(App.Path & "\" & App.EXEName & ".ini", "Contacts", "IpContact" & i, "")
    PortContactIni = ReadStringINI(App.Path & "\" & App.EXEName & ".ini", "Contacts", "PortContact" & i, "")
    
    If IPContactIni = IPContact And _
       PortContactIni = PortContact Then
       ValidaExisteContacto = True
       MsgBox "El contacto Existe con el Nombre de Usuario: " & ContactNameIni, vbInformation, "Contacto Existe"
       Exit For
    End If
    
Next i



End Function


Public Sub LoadContacts(ObjectDisplayContacts As TreeView, ImageDisplay As ImageList)
On Error Resume Next
Dim ContactName As String
Dim IPContact As String
Dim PortContact As String
Dim i As Integer

Dim Nod_A
Dim Nod_B

With ObjectDisplayContacts

.LineStyle = tvwRootLines
.ImageList = ImageDisplay

If ContactCount = 0 Then

    .Nodes.Clear
    Set Nod_A = .Nodes.Add(, , "root", "Sin contactos", 3, 3)
    .Enabled = False
    frmChatForm.Command1.Enabled = False

Else

    .Nodes.Clear
    .Enabled = True
    
    For i = 1 To ContactCount
    ContactName = ""
    IPContact = ""
    PortContact = ""
    
    ContactName = ReadStringINI(App.Path & "\" & App.EXEName & ".ini", "Contacts", "ContactName" & i, "")
    IPContact = ReadStringINI(App.Path & "\" & App.EXEName & ".ini", "Contacts", "IpContact" & i, "")
    PortContact = ReadStringINI(App.Path & "\" & App.EXEName & ".ini", "Contacts", "PortContact" & i, "")
    
    If Nulidad(ContactName) = False And _
       Nulidad(IPContact) = False And _
       Nulidad(PortContact) = False Then
       
           
        Set Nod_A = .Nodes.Add(, , i & "-" & IPContact & "|" & PortContact, ContactName, 2, 2)
      
    End If
    
    Next i
    
    If .Nodes.Count = 0 Then
    
        .Nodes.Clear
        Set Nod_A = .Nodes.Add(, , "root", "Sin contactos", 3, 3)
        .Enabled = False
    Else
    
       .Nodes(1).Selected = True
       
       
    End If

End If

End With

End Sub



Public Sub PictureMoveMouse(FormInvoqued As Form, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
Static rec As Boolean, Msg As Long, ValDev As Long

    Msg = x / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case Msg
            Case WM_LBUTTONDBLCLK:
                FormInvoqued.Show
                FormInvoqued.Text1.SetFocus
                
            Case WM_LBUTTONDOWN:
            Case WM_LBUTTONUP:
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
                 ' PopUp menu,2 significa Izq/Der botones en el menu, mnuAbout es BOLD
                 FormInvoqued.PopupMenu FormInvoqued.Main1
            End Select
        rec = False
    End If

End Sub

   
   
Public Sub CloseIconAPP()
On Error Resume Next
   NotificaPantalla.cbSize = Len(NotificaPantalla)
   NotificaPantalla.hWnd = UltimaImagen.hWnd
   NotificaPantalla.uId = 1&
   Shell_NotifyIcon NIM_DELETE, NotificaPantalla
   
End Sub


Public Sub StartIconAPP(ActionIcon As ActionMessenger, ImageIcon As PictureBox)
On Error Resume Next
    Set UltimaImagen = Nothing
    Set UltimaImagen = ImageIcon
    
    NotificaPantalla.cbSize = Len(NotificaPantalla)
    NotificaPantalla.hWnd = ImageIcon.hWnd
    NotificaPantalla.uId = 1&
    NotificaPantalla.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    NotificaPantalla.ucallbackMessage = WM_MOUSEMOVE
    NotificaPantalla.hIcon = ImageIcon.Picture
'---------------------------------
    Select Case ActionIcon
     Case OpenMessenger
          NotificaPantalla.szTip = LocalName & ":" & Chr(13) & "Messenger TCP V." & App.Major & "." & App.Minor & Chr$(0)  ' Es un string de "C" ( \0 )
     Case EnterMessage
          NotificaPantalla.szTip = "Usted tiene Mensajes Nuevos" & Chr$(0) ' Es un string de "C" ( \0 )
    End Select
    
    Shell_NotifyIcon NIM_ADD, NotificaPantalla
    App.TaskVisible = True
    
End Sub

Public Function CompruebaIp(IpComprobar) As Boolean
On Error GoTo Err_CompruebaIp
Dim CounterSeparator As Integer
Dim AcumSection As String
Dim Segment1 As Boolean
Dim Segment2 As Boolean
Dim Segment3 As Boolean
Dim Segment4 As Boolean

Segment1 = False
Segment2 = False
Segment3 = False
Segment4 = False

CounterSeparator = 0
AcumSection = ""

If Len(IpComprobar) <= 15 And _
   Mid(IpComprobar, 1, 1) <> "." And _
   Mid(IpComprobar, Len(IpComprobar), 1) <> "." Then

For i = 1 To Len(IpComprobar)
   AcumSection = AcumSection + Mid(IpComprobar, i, 1)
   If Mid(IpComprobar, i, 1) = "." Then
     CounterSeparator = CounterSeparator + 1
     Select Case CounterSeparator
       Case 1
             If (Len(AcumSection) - 1 <= 3) And _
                (IsNumeric(Mid(AcumSection, 1, Len(AcumSection) - 1))) Then
                If Val(Mid(AcumSection, 1, Len(AcumSection) - 1)) <= 255 Then
                   Segment1 = True
                   AcumSection = ""
                End If
             End If
       Case 2
             If (Len(AcumSection) - 1 <= 3) And _
                (IsNumeric(Mid(AcumSection, 1, Len(AcumSection) - 1))) Then
                If Val(Mid(AcumSection, 1, Len(AcumSection) - 1)) <= 255 Then
                   Segment2 = True
                   AcumSection = ""
                End If
             End If
       Case 3
            If (Len(AcumSection) - 1 <= 3) And _
                (IsNumeric(Mid(AcumSection, 1, Len(AcumSection) - 1))) Then
                If Val(Mid(AcumSection, 1, Len(AcumSection) - 1)) <= 255 Then
                   Segment3 = True
                   AcumSection = ""
                End If
             End If
    End Select
   End If
Next i

If (Len(AcumSection) <= 3) And _
   (IsNumeric(AcumSection)) Then
     If Val(AcumSection) <= 255 Then
         Segment4 = True
         AcumSection = ""
     End If
End If
If Segment1 = True And Segment2 = True And _
   Segment3 = True And Segment4 = True And _
   CounterSeparator = 3 Then
             
   CompruebaIp = True
Else
CompruebaIp = False
End If
Else
CompruebaIp = False
End If

Exit_Err_CompruebaIp:
Exit Function

Err_CompruebaIp:
CompruebaIp = False
Resume Exit_Err_CompruebaIp
End Function




Public Sub EditOrNewContact(TypeAction As TypeActionContact, Buttom As CommandButton)
On Error Resume Next

If TypeAction = NewContact Then
    DlgContacts.ActionForm = TypeAction
    DlgContacts.Caption = "Ingreso de nuevo Contacto"
    DlgContacts.Icon = Buttom.Picture
    DlgContacts.Show 1
    LoadContacts frmChatForm.TreeView1, frmChatForm.ImageList1
Else
    DlgContacts.ActionForm = TypeAction
    DlgContacts.Caption = "Editar datos de contacto"
    DlgContacts.Icon = Buttom.Picture
    DlgContacts.IniItem = Mid(frmChatForm.TreeView1.SelectedItem.Key, 1, InStr(1, frmChatForm.TreeView1.SelectedItem.Key, "-") - 1)
    DlgContacts.Text1.Text = frmChatForm.TreeView1.SelectedItem.Text
    DlgContacts.Text2.Text = Mid(frmChatForm.TreeView1.SelectedItem.Key, InStr(1, frmChatForm.TreeView1.SelectedItem.Key, "-") + 1, _
                             InStr(1, frmChatForm.TreeView1.SelectedItem.Key, "|") - InStr(1, frmChatForm.TreeView1.SelectedItem.Key, "-") - 1)
    DlgContacts.Text3.Text = Mid(frmChatForm.TreeView1.SelectedItem.Key, InStr(1, frmChatForm.TreeView1.SelectedItem.Key, "|") + 1, Len(frmChatForm.TreeView1.SelectedItem.Key) - InStr(1, frmChatForm.TreeView1.SelectedItem.Key, "|"))
    DlgContacts.ContactNameIn = DlgContacts.Text1.Text
    DlgContacts.IpContactIn = DlgContacts.Text2.Text
    DlgContacts.PortContactIn = DlgContacts.Text3.Text
    DlgContacts.Show 1
    LoadContacts frmChatForm.TreeView1, frmChatForm.ImageList1
    
End If
With frmChatForm
If .TreeView1.Enabled = False Then
       .Command3.Enabled = False
       .Command4.Enabled = False
Else
       .Command3.Enabled = True
       .Command4.Enabled = True
End If
End With

End Sub


Public Sub ValidateContactConect(ListContact As TreeView, ClientConect As Winsock)
On Error Resume Next
Dim i As Integer
Dim j As Integer

If ListContact.Enabled = True Then

For i = 1 To ListContact.Nodes.Count
  DoEvents
  
  If ClientConect.State <> 0 Then
     ClientConect.Close
  End If
  
  ClientConect.RemoteHost = Mid(ListContact.Nodes(i).Key, InStr(1, ListContact.Nodes(i).Key, "-") + 1, _
                            InStr(1, ListContact.Nodes(i).Key, "|") - InStr(1, ListContact.Nodes(i).Key, "-") - 1)
  ClientConect.RemotePort = Mid(ListContact.Nodes(i).Key, InStr(1, ListContact.Nodes(i).Key, "|") + 1, Len(ListContact.Nodes(i).Key) - InStr(1, ListContact.Nodes(i).Key, "|"))
  
  ClientConect.Protocol = sckTCPProtocol
  ClientConect.Connect
  
   For j = 1 To 50
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
   Next j
  
  
  
  If ClientConect.State = 7 Then
    
    If ListContact.Nodes(i).Selected = True Then
        frmChatForm.Label4.Caption = "Hablando Con: " & ListContact.Nodes(i).Text
    End If
    ListContact.Nodes(i).Image = 1
    ListContact.Nodes(i).SelectedImage = 1
    
  Else
    If ListContact.Nodes(i).Selected = True Then
       frmChatForm.Label4.Caption = "Hablando Con: "
       frmChatForm.Command1.Enabled = False
    End If
    
    ListContact.Nodes(i).Image = 2
    ListContact.Nodes(i).SelectedImage = 2
  End If
  
 If ClientConect.State <> 0 Then
    ClientConect.Close
 End If
 
Next i


End If

End Sub


Public Function LimpiaEmotIcon(ByRef LineSegment As String) As String
Dim j As Integer


    LimpiaEmotIcon = ""
    LimpiaEmotIcon = LineSegment
    
      For j = 0 To NumMaxIconsIndex
              
        LimpiaEmotIcon = Replace(Replace(LimpiaEmotIcon, IconsPosibles(j), ""), LCase(IconsPosibles(j)), "")
        
      Next j

End Function

Public Sub PintaEmotIcons(ByRef ObjectToPaintIcon As RichTextBox, ByRef LineSegment As String)
Dim LongitudMensaje As Integer
Dim SegmentoAEvaluar As String
Dim i As Integer
Dim j As Integer
Dim NumIconsInLine As Integer
Dim ArrayStringIcons As String
Dim ValoresEncontrados As Variant



If LineaEmotIcon(UCase(LineSegment), NumIconsInLine, ArrayStringIcons) = True Then

ValoresEncontrados = Split(ArrayStringIcons, "|")

    For i = 0 To NumIconsInLine - 1
    
        ContenidoClip = ""
        ContenidoClip = Clipboard.GetText
        Clipboard.Clear
        Clipboard.SetData LoadPicture(App.Path & "\icons\A" & ValoresEncontrados(i) & ".gif")
        ObjectToPaintIcon.SelStart = Len(ObjectToPaintIcon.Text)
        ObjectToPaintIcon.Locked = False
        SendMessage ObjectToPaintIcon.hWnd, WM_PASTE, 0, 0
        ObjectToPaintIcon.Locked = True
        Clipboard.Clear
        Clipboard.SetText ContenidoClip
        ContenidoClip = ""
   Next i
   
Else
   MsgBox "Error no es una linea de Iconos", vbCritical, "Error de Uso"
End If

End Sub
Public Function LineaEmotIcon(ByRef LineSegment As String, Optional ByRef NumIcons As Integer, _
                              Optional ByRef PosArrayIcons As String) As Boolean
'//primero verificar todo el bloque enviado
'//si el bloque contiene solo emoticons entonces puede insertar imagenes
Dim LongitudMensaje As Integer
Dim i As Integer
Dim SegmentoAEvaluar As String
Dim EncontroEnArray As Boolean
Dim AcumuladorLongitudes As Integer
Dim j As Integer

AcumuladorLongitudes = 0
NumIcons = 0
PosArrayIcons = ""

LongitudMensaje = Len(LineSegment)

For i = 1 To LongitudMensaje
 '*************************************************************************
    If Mid(LineSegment, i, 1) = "<" Then
      SegmentoAEvaluar = ""
    End If
 '*************************************************************************
 
    SegmentoAEvaluar = SegmentoAEvaluar & Mid(LineSegment, i, 1)
    EncontroEnArray = False
    
 '*************************************************************************
    If Mid(LineSegment, i, 1) = ">" Then
        
        For j = 0 To NumMaxIconsIndex
           If IconsPosibles(j) = SegmentoAEvaluar Then
              EncontroEnArray = True
              Exit For
           End If
        Next j
        
        If EncontroEnArray = True Then
          AcumuladorLongitudes = AcumuladorLongitudes + Len(SegmentoAEvaluar)
          NumIcons = NumIcons + 1
          PosArrayIcons = PosArrayIcons & j & "|"
          SegmentoAEvaluar = ""
          EncontroEnArray = False
        End If
    
    End If
 '*************************************************************************
 
Next i

If AcumuladorLongitudes = LongitudMensaje Then
   LineaEmotIcon = True
Else
   LineaEmotIcon = False
   NumIcons = 0
   PosArrayIcons = ""
End If

End Function
