Attribute VB_Name = "Funciones_API"
Option Explicit

'Constantes de retorno de error para ShellExecute
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_OOM = 8
Private Const SE_ERR_PNF = 3
Private Const SE_ERR_SHARE = 26

Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long

'Declaración de ShellExecute
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal _
                        lpFile As String, ByVal lpParameters As String, ByVal _
                        lpDirectory As String, ByVal nShowCmd As Long) As Long
                        
Private Const SW_SHOWNORMAL = 1

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long

'Minimiza una ventana
Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
'La región está habilitada?
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
'Función que habilita una región !!!
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
'Mueve una ventana
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'Destruye una ventna
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
'Establece la posición del cursor en dichas coordenadas
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'Muestra u oculta el cursor
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SetDoubleClickTime Lib "user32" (ByVal wCount As Long) As Long
Private Declare Function GetDoubleClickTime Lib "user32" () As Long

Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

' función SendMessage
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

'Obtiene el Handle de una ventana a partir de una coordenada
Declare Function WindowFromPoint _
    Lib "user32" ( _
        ByVal xPoint As Long, _
        ByVal yPoint As Long) As Long

'Recupera la coordenada del cursor
Declare Function GetCursorPos _
    Lib "user32" ( _
        lpPoint As POINTAPI) As Long

'Recupera el nombre de la clase de Ventana  a partir de su handle
Declare Function GetClassName _
    Lib "user32" _
    Alias "GetClassNameA" ( _
        ByVal hwnd As Long, _
        ByVal lpClassName As String, _
        ByVal nMaxCount As Long) As Long

'REcupera el Handle de la ventana padre de una ventana
Declare Function GetParent _
    Lib "user32" ( _
        ByVal hwnd As Long) As Long

'Recupera el Handle de la ventana padre que está en primer nivel
Declare Function GetAncestor Lib "user32.dll" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long

'Función Api GetAsyncKeyState
Declare Function GetAsyncKeyState _
    Lib "user32" ( _
        ByVal vKey As Long) As Integer

'Udt
'##############
 Type ProcData
    HwndWin As Long
    captionWin As String
    Estado As String
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Type POINTAPI
    X As Long
    Y As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type WINDOWPLACEMENT
    length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type
 
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As _
                                                            WINDOWPLACEMENT) As Long
'Declaración del Api keybd_event para la presión de tecla
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, _
    ByVal dwExtraInfo As Long)

'Constantes para las teclas y otros
Const KEYEVENTF_KEYUP = &H2
Const INPUT_MOUSE = 0
Const INPUT_KEYBOARD = 1
Const INPUT_HARDWARE = 2
Const KEYEVENTF_EXTENDEDKEY = &H1

Enum eButtons
    eNothing = 1
    eRightClick = 2
    eDoubleRight = 4
    eLeftClick = 8
    eDoubleLeft = 16
End Enum

Private Type MOUSEINPUT
  dX As Long
  dY As Long
  mouseData As Long
  dwFlags As Long
  time As Long
  dwExtraInfo As Long
End Type

Private Type KEYBDINPUT
  wVk As Integer
  wScan As Integer
  dwFlags As Long
  time As Long
  dwExtraInfo As Long
End Type
Private Type HARDWAREINPUT
  uMsg As Long
  wParamL As Integer
  wParamH As Integer
End Type
Private Type GENERALINPUT
  dwType As Long
  xi(0 To 23) As Byte
End Type

Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dY As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'Para manipular archivos INI
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long

Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4
Const MOUSEEVENTF_MIDDLEDOWN = &H20
Const MOUSEEVENTF_MIDDLEUP = &H40
Const MOUSEEVENTF_MOVE = &H1
Const MOUSEEVENTF_ABSOLUTE = &H8000
Const MOUSEEVENTF_RIGHTDOWN = &H8
Const MOUSEEVENTF_RIGHTUP = &H10

' constantes para SendMessage
Const BM_SETSTATE = &HF3
Const WM_LBUTTONDOWN = &H201  ' botón izquierdo abajo
Const WM_LBUTTONUP = &H202  ' izquierdo arriba
Const WM_LBUTTONDBLCLK = &H203   ' izquierdo doble click
Const WM_RBUTTONDOWN = &H204   'Botón derecho abajo
Const WM_RBUTTONUP = &H205   'derecho arriba

'Para usar el portapapeles con la api
Const WM_COPY As Long = &H301 'Copiar al portapapeles
Const WM_CUT As Long = &H300 'Cortar (solo si hay algo seleccionado)
Const WM_PASTE As Long = &H302 'Pegar
Const EM_SETSEL As Long = &HB1 'Seleccionar texto

Public HwndPadre As Long
Public EnumChild As Long
Public TxtChild As String

Public Sub Ejecutar(ByVal hwnd As Long, ByVal Archivo As String)
    'Abre el archivo
    Select Case ShellExecute(hwnd, "Open", Archivo, vbNullString, _
                            vbNullString, SW_SHOWNORMAL)
        'Retorno de ShellExecute
        Case 0
            MsgBox "Fuera de memoria o de recursos", vbCritical
        Case ERROR_BAD_FORMAT
            MsgBox "Formato de archivo no váldio", vbCritical
        Case SE_ERR_ACCESSDENIED
            MsgBox "Acceso denegado al Intentar abrir el archivo", vbCritical
        Case SE_ERR_ASSOCINCOMPLETE
            MsgBox "Extensión incompleta o inválida del nombre del archivo", vbCritical
        Case SE_ERR_DDEBUSY
            MsgBox "DDE ocupado ", vbCritical
        Case SE_ERR_DDEFAIL
            MsgBox "Caída en la Transacción DDE", vbCritical
        Case SE_ERR_DDETIMEOUT
            MsgBox "Petición DDE fuera de tiempo", vbCritical
        Case SE_ERR_DLLNOTFOUND
            MsgBox "DLL no encontrada", vbCritical
        Case ERROR_FILE_NOT_FOUND, SE_ERR_FNF
            MsgBox "Archivo no encontrado", vbCritical
        Case SE_ERR_NOASSOC
            MsgBox "El archivo no está asociado a ninguna aplicación", vbCritical
        Case SE_ERR_OOM
            MsgBox "Fuera de memoria", vbCritical
        Case ERROR_PATH_NOT_FOUND, SE_ERR_PNF
            MsgBox "Ruta no encontrada", vbCritical
        Case SE_ERR_SHARE
            MsgBox "Sharing violation", vbCritical
    End Select
End Sub

'Funcion de lectura de un archivo INI
Public Function LeeINI(ByVal Filename As String, ByVal Key_Value As String, ByVal Key_Name As String, Optional ByVal Default As String) As String
    On Error GoTo ErrOut
    Dim Size As Integer
    Dim Value As String
    'Comprobamos que el archivo existe:
    If Dir(Filename) = "" Then Err.Raise 53
    'Se define el tamaño maximo de caracteres
    'que podra tener la variable Value:
    Value = Space(256)
    'Se utiliza la función para obtener
    'el valor de la clave:
    Size = GetPrivateProfileString(Key_Value, Key_Name, Default, Value, Len(Value), Filename)

    'Si el tamaño es mayor a 0 entonces
    'se ha encontrado el valor de la clave:
    If Size > 0 Then
        Value = VBA.Left$(Value, Size)
    
    'Devolvemos el valor de la clave:
    If VBA.Right$(VBA.Trim$(Value), 1) = Chr(0) Then Value = VBA.Left$(VBA.Trim$(Value), Len(VBA.Trim$(Value)) - 1)
    
    LeeINI = VBA.Trim$(Value)
    Exit Function
    End If
ErrOut:
    LeeINI = Default
End Function

'Metodo de escritura. Si el archivo al que se refiere el parametro Archivo
'Si no existe se crea automaticamente:
Public Sub GrabaINI(ByVal Filename As String, ByVal Key_Value As String, ByVal Key_Name As String, ByVal Value As String)
    On Error GoTo ErrOut
    Dim Size As Integer
    'Escribimos el valor de la clave:
    Size = WritePrivateProfileString(Key_Value, Key_Name, Value, Filename)
ErrOut:
End Sub

Function ObtenerWinDir(Optional SoloLetraUnidad As Boolean = True) As String
Dim sSave As String, ret As Long
'Create a buffer
sSave = Space(255)
'Get the system directory
ret = GetSystemDirectory(sSave, 255)
'Remove all unnecessary chr$(0)'s
sSave = Left$(sSave, ret)
If SoloLetraUnidad Then ObtenerWinDir = Left(sSave, 2)
If Not SoloLetraUnidad Then ObtenerWinDir = sSave
End Function

Sub Pulsar_Tecla(ByVal Tecla As Byte)
    Call keybd_event(Tecla, 0, 0, 0)
    Call keybd_event(Tecla, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub SendKey(bKey As Byte) 'Idéntica a la anterior no hay diferencias
    Dim GInput(0 To 1) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    KInput.wVk = bKey  'the key we're going to press
    KInput.dwFlags = 0 'press the key
    'copy the structure into the input array's buffer.
    GInput(0).dwType = INPUT_KEYBOARD   ' keyboard input
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)
    'do the same as above, but for releasing the key
    KInput.wVk = bKey  ' the key we're going to realease
    KInput.dwFlags = KEYEVENTF_KEYUP  ' release the key
    GInput(1).dwType = INPUT_KEYBOARD  ' keyboard input
    CopyMemory GInput(1).xi(0), KInput, Len(KInput)
    'send the input now
    Call SendInput(2, GInput(0), Len(GInput(0)))
End Sub

'A diferencia de la otra funcion esta simula un click en la posición del cursor
'sin utilizar hwnd
Sub EnviarClickXY(Optional dX As Long = 0, Optional dY As Long = 0, Optional Buttons As eButtons = eNothing, Optional ByVal OcultarCursor As Boolean = True, Optional UseAbsolute As Boolean = False)
    Const MOUSEEVENTF_ABSOLUTE = &H8000, MOUSEEVENTF_LEFTDOWN = &H2, MOUSEEVENTF_LEFTUP = &H4
    Const MOUSEEVENTF_MIDDLEDOWN = &H20, MOUSEEVENTF_MIDDLEUP = &H40, MOUSEEVENTF_MOVE = &H1
    Const MOUSEEVENTF_RIGHTDOWN = &H8, MOUSEEVENTF_RIGHTUP = &H10
    Dim lCoordBit As Long
    Dim PosCur As POINTAPI
    
    If OcultarCursor Then
        GetCursorPos PosCur
        ShowCursor 0
    End If
    
    If UseAbsolute Then
        'Add the bit for using absolute coordinates
        lCoordBit = MOUSEEVENTF_ABSOLUTE
    End If
    
    Select Case Buttons
    Case eNothing
        'Just move the mouse
        SetCursorPos dX, dY
    Case eRightClick
        SetCursorPos dX, dY
        DoEvents
        'Send a right click
        mouse_event MOUSEEVENTF_RIGHTDOWN Or lCoordBit, 0&, 0&, 0&, 0&
        mouse_event MOUSEEVENTF_RIGHTUP Or lCoordBit, 0&, 0&, 0&, 0&

    Case eDoubleRight
        SetCursorPos dX, dY
        DoEvents
        'Send a right click
        mouse_event MOUSEEVENTF_RIGHTDOWN Or lCoordBit, 0&, 0&, 0&, 0&
        mouse_event MOUSEEVENTF_RIGHTUP Or lCoordBit, 0&, 0&, 0&, 0&
        DoEvents
        'Send a right click
        mouse_event MOUSEEVENTF_RIGHTDOWN Or lCoordBit, 0&, 0&, 0&, 0&
        mouse_event MOUSEEVENTF_RIGHTUP, 0&, 0&, 0&, 0&
    Case eLeftClick
        SetCursorPos dX, dY
        DoEvents
        'Send a left click
        mouse_event MOUSEEVENTF_LEFTDOWN Or lCoordBit, 0&, 0&, 0&, 0&
        mouse_event MOUSEEVENTF_LEFTUP Or lCoordBit, 0&, 0&, 0&, 0&
    Case eDoubleLeft
        SetCursorPos dX, dY
        DoEvents
        'Send a left click
        mouse_event MOUSEEVENTF_LEFTDOWN Or lCoordBit, 0&, 0&, 0&, 0&
        mouse_event MOUSEEVENTF_LEFTUP Or lCoordBit, 0&, 0&, 0&, 0&
        DoEvents
        'Send a left click
        mouse_event MOUSEEVENTF_LEFTDOWN Or lCoordBit, 0&, 0&, 0&, 0&
        mouse_event MOUSEEVENTF_LEFTUP Or lCoordBit, 0&, 0&, 0&, 0&
    End Select

If OcultarCursor Then
    SetCursorPos PosCur.X, PosCur.Y
    ShowCursor 1
End If
End Sub

'Enviar click de mouse al Hwnd indicado
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub EnviarClick(ByVal handle As Long, ByVal BotónClick As Integer)
    Call SendMessage(handle, BM_SETSTATE, 0, ByVal 0&)
Select Case BotónClick
Case 1 'Click
    Call SendMessage(handle, &H201, 0, ByVal 0&)
    Call SendMessage(handle, &H202, 0, ByVal 0&)
Case 2 'Click derecho
    Call SendMessage(handle, &H204, 0, ByVal 0&)
    Call SendMessage(handle, &H205, 0, ByVal 0&)
Case 3 'Doble click
    Call SendMessage(handle, &H201, 0, ByVal 0&)
    Call SendMessage(handle, &H202, 0, ByVal 0&)
    Sleep (GetDoubleClickTime)
    Call SendMessage(handle, &H201, 0, ByVal 0&)
    Call SendMessage(handle, &H202, 0, ByVal 0&)
End Select
'   Call SendMessage(handle, BM_SETSTATE, 1, ByVal 0&)
End Sub

Public Function GetWindowsPlacementx(ByVal hwnd As Long) As ProcData
Dim infoWin As WINDOWPLACEMENT
Dim winenum As New clsEnum
Dim p As ProcData
infoWin.length = Len(infoWin)
Call GetWindowPlacement(hwnd, infoWin)
p.captionWin = winenum.GetWindowCaption(hwnd)
With p
        'Guarda el estado de la ventana
        Select Case infoWin.showCmd
            Case 1
                .Estado = "Normal"
            Case 2
                .Estado = "Minimizado"
            Case 3
                .Estado = "Maximizado"
        End Select
                'Guarda la posición de la ventana
                .Left = infoWin.rcNormalPosition.Left
                .Top = infoWin.rcNormalPosition.Top
                .Right = infoWin.rcNormalPosition.Right
                .Bottom = infoWin.rcNormalPosition.Bottom
End With
GetWindowsPlacementx = p
End Function

Public Sub PegarTextoHwnd(ByVal hwnd As Long, ByVal Texto As String, ByRef TmpTextBox As TextBox)
TmpTextBox.Text = Trim(Texto)
If TmpTextBox.SelLength = 0 Then SendMessage TmpTextBox.hwnd, EM_SETSEL, 0&, -1&
SendMessage TmpTextBox.hwnd, WM_COPY, 0&, 0&  'Copio al portapapeles
SendMessage hwnd, WM_PASTE, 0&, 0&  'Pego en el destino
End Sub

Public Sub PegarTexto(ByVal hwnd As Long, ByVal Texto As String)
Clipboard.Clear
Clipboard.SetText Trim(Texto)
'TomarFoco hwnd
'Ctrl + v (pegar el texto enviado al clipboard)
'Esto es equivalente a -> SendKeys "^(v)" pero ha funcionado mejor que este
Call keybd_event(17, 0, 0, 0) 'Mantengo presionado la tecla ctrl
Call keybd_event(86, 0, 0, 0) 'Mantengo presionado la tecla v
Call keybd_event(17, 0, KEYEVENTF_KEYUP, 0) 'Suelto la tecla ctrl
Call keybd_event(86, 0, KEYEVENTF_KEYUP, 0) 'Suelto la tecla v
End Sub

Public Function TomarFoco(ByVal hwnd As Long) As Long
TomarFoco = SetFocus(hwnd)
End Function

Public Sub OrdenarDataGrid(ByVal ColIndex As Integer, _
                           rs As ADODB.Recordset, _
                           DataGrid As DataGrid)

    Dim strColName As String
    Static bSortAsc As Boolean
    Static strPrevCol As String
    
    strColName = DataGrid.Columns(ColIndex).DataField
 
    If strColName = strPrevCol Then

        If bSortAsc Then
            rs.Sort = strColName & " DESC"
            bSortAsc = False
        Else
            rs.Sort = strColName
            bSortAsc = True
        End If

    Else
        rs.Sort = strColName
        bSortAsc = True
    End If
 strPrevCol = strColName
End Sub

Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim sSave As String
    'Get the windowtext length
    sSave = Space$(GetWindowTextLength(hwnd) + 1)
    'get the window text
    GetWindowText hwnd, sSave, Len(sSave)
    'remove the last Chr$(0)
    sSave = Left$(sSave, Len(sSave) - 1)
    If sSave <> "" And hwnd <> 0 Then
    If Trim(LCase(sSave)) = TxtChild Then
        EnumChild = hwnd
        EnumChildProc = 0
        Exit Function
    End If
    End If
    'continue enumeration
    EnumChildProc = 1
End Function

