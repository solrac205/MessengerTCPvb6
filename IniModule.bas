Attribute VB_Name = "IniModule"
Option Explicit
Global TCPListenPort As Integer
Global ContactCount As Integer
Global LocalName As String
Global MSGMensajeInactividad As String

Public Declare Function SystemParametersInfo Lib "user32.dll" Alias _
    "SystemParametersInfoA" ( _
        ByVal uiAction As Long, _
        ByVal uiParam As Long, _
        pvParam As Any, _
        ByVal fWinIni As Long) As Long
  
' Mensaje para obtener el área de la pantalla sin contar el taskBar
Public Const SPI_GETWORKAREA = 48
  
Public Enum DatosDeArchivo
 '// Colección para identificación del dato de Fecha de Archivo que queremos consultar
 '// del archivo consultado en funcion DateToFile
     FechaCreacion = 1
     FechaModificacion = 2
     FechaUltimoAcceso = 3
End Enum

  
' Estrucura Rect que retorna la posición y dimensiones del área de trabajo
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
    
Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
    
'//mas adelante talvez introducir emoticons
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

Public Const WM_PASTE = &H302
'//mas adelante talvez introducir emoticons
'Clipboard.Clear
'Clipboard.SetData ImageList4.ListImages(ListView1.SelectedItem.Index).Picture
'SendMessage RichTextBox2.hWnd, WM_PASTE, 0, 0

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LEE EL INIFILE
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReadStringINI(cINIFile As String, cSection As String, cKey As String, CDefault As String) As String
On Error GoTo Err_ReadStringIni

Dim cTemp As String
Dim IOK%
ReadStringINI = True

cTemp = Space$(255)
IOK% = GetPrivateProfileString(cSection, cKey, CDefault, cTemp, Len(cTemp), cINIFile)

If IOK% > 0 Then
   ReadStringINI = Left$(cTemp, IOK%)
Else
   ReadStringINI = CDefault
End If
Exit Function

Err_ReadStringIni:
  ReadStringINI = CDefault
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ESCRIBE EL INIFILE
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function WriteStringINI(cINIFile As String, cSection As String, cKey As String, lpString As String)
On Error GoTo Err_WriteStringIni

Dim IOK%
WriteStringINI = True

IOK% = WritePrivateProfileString(cSection, cKey, lpString, cINIFile)

If IOK% > 0 Then
Else
   
   WriteStringINI = False
End If
Exit Function

Err_WriteStringIni:
  WriteStringINI = False
End Function

Public Function DateToFile(ByRef FileToEvaluate As String, ByRef TextObject As TextBox, ByRef DateToConsult As DatosDeArchivo) As String
'// Extracción de datos de fechas de un archivo cualquiera que este sea, utilizando tres parametros de entrada y dandonos como resultado un string.
'// Los Parametros de entrada son: Archivo a Evaluar, Objeto Tipo TextBox y tipo de Fecha de consulta
On Error GoTo Err_DateToFile
Dim FileToAcces
Dim File
Dim Drive
Dim Folder
Dim SubFol As TextBox
Dim i As Integer

TextObject.Text = ""
Set SubFol = TextObject

Set FileToAcces = CreateObject("Scripting.FileSystemObject")
Set Drive = FileToAcces.Drives(Mid(FileToEvaluate, 1, 1))
Set Folder = Drive.RootFolder
For i = 1 To Len(Mid(FileToEvaluate, 4, InStr(1, FileToEvaluate, FileToAcces.GetBaseName(FileToEvaluate) & "." & FileToAcces.GetExtensionName(FileToEvaluate)) - 5))
  If Mid(Mid(FileToEvaluate, 4, InStr(1, FileToEvaluate, FileToAcces.GetBaseName(FileToEvaluate) & "." & FileToAcces.GetExtensionName(FileToEvaluate)) - 5), i, 1) = "\" Then
    If Dir(Folder & "\" & SubFol.Text, vbDirectory) = "" Then
     SubFol.Text = SubFol.Text & "."
    End If
    Set Folder = Folder.SubFolders(SubFol.Text)
    SubFol.Text = ""
  Else
    SubFol.Text = SubFol.Text & Mid(Mid(FileToEvaluate, 4, InStr(1, FileToEvaluate, FileToAcces.GetBaseName(FileToEvaluate) & "." & FileToAcces.GetExtensionName(FileToEvaluate)) - 5), i, 1)
  End If
Next i
If Dir(Folder & "\" & SubFol.Text, vbDirectory) = "" Then
 SubFol.Text = SubFol.Text & "."
End If
Set Folder = Folder.SubFolders(SubFol.Text)
Set File = Folder.Files(FileToAcces.GetBaseName(FileToEvaluate) & "." & FileToAcces.GetExtensionName(FileToEvaluate))
Select Case DateToConsult
Case 1
DateToFile = File.DateCreated
Case 2
DateToFile = File.DateLastModified
Case 3
DateToFile = File.DateLastAccessed
End Select

Exit_Err_DateToFile:
Exit Function

Err_DateToFile:
DateToFile = "File Or Path Not Found"
Resume Exit_Err_DateToFile
End Function


Public Sub InitConfValues()
On Error GoTo Err_InitConfValues

TCPListenPort = Int(ReadStringINI(App.Path & "\" & App.EXEName & ".ini", "Service", "ListenPort", ""))
ContactCount = Int(ReadStringINI(App.Path & "\" & App.EXEName & ".ini", "Service", "ContactCount", ""))
LocalName = ReadStringINI(App.Path & "\" & App.EXEName & ".ini", "Service", "LocalName", "")
MSGMensajeInactividad = ReadStringINI(App.Path & "\" & App.EXEName & ".ini", "Service", "MsgInactividad", "")

Exit_Err_InitConfValues:
Exit Sub

Err_InitConfValues:
  MsgBox "Error en valores iniciales", vbCritical, "Error en Carga de Datos IniFile"
  End
Resume Exit_Err_InitConfValues
End Sub
