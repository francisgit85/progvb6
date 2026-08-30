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

Private Declare Function SystemParametersInfo Lib "user32" _
    Alias "SystemParametersInfoA" (ByVal uAction As Long, _
    ByVal uParam As Long, ByVal lpvParam As Any, _
    ByVal fuWinIni As Long) As Long

Const SPI_SETDEFAULTINPUTLANG = &H5A
Const SPIF_SENDCHANGE = &H2

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" _
    (ByVal hwnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    
Declare Function LoadKeyboardLayout Lib "user32" _
    Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long

Declare Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Declare Function AttachThreadInput Lib "user32" _
    (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long

Declare Function ActivateKeyboardLayout Lib "user32" _
    (ByVal hkl As Long, ByVal flags As Long) As Long

Private Declare Function GetKeyboardLayoutName Lib "user32" _
    Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long

' API para enumerar ventanas
Public Declare Function EnumWindows Lib "user32" _
    (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SendInput Lib "user32" _
    (ByVal nInputs As Long, pInputs As Any, ByVal cbSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal length As Long)

Public Declare Function GetKeyboardLayout Lib "user32" _
    (ByVal idThread As Long) As Long
    
Public Declare Function GetForegroundWindow Lib "user32" () As Long
    
    
Private Type KEYBDINPUT
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Type KINPUT
    dwType As Long
    ki As KEYBDINPUT
End Type


Private Declare Function MapVirtualKey Lib "user32" _
    Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

' Constantes
Private Const INPUT_KEYBOARD = 1
Private Const KEYEVENTF_KEYUP = &H2
Public Const KEYEVENTF_UNICODE As Long = &H4
Private Const VK_SHIFT = &H10
Private Const VK_RETURN = &HD
Private Const VK_TAB = &H9
Private Const VK_SPACE = &H20
Private Const VK_ADD = &H6B
Private Const VK_SUBTRACT = &H6D
Private Const VK_MULTIPLY = &H6A
Private Const VK_DIVIDE = &H6F
Private Const VK_DECIMAL = &H6E
Private Const VK_CAPITAL = &H14

' Diccionario Virtual Key Codes (US Standard)

Const VK_OEM_1 = &HBA       ' ; :
Const VK_OEM_PLUS = &HBB    ' = +
Const VK_OEM_COMMA = &HBC   ' , <
Const VK_OEM_MINUS = &HBD   ' -_
Const VK_OEM_PERIOD = &HBE  ' . >
Const VK_OEM_2 = &HBF       ' / ?
Const VK_OEM_3 = &HC0       ' ` ~
Const VK_OEM_4 = &HDB       ' [ {
Const VK_OEM_5 = &HDC       ' \ |
Const VK_OEM_6 = &HDD       ' ] }
Const VK_OEM_7 = &HDE       ' ' "
Const VK_OEM_8 = &HDF       ' (varía según layout)
Const VK_OEM_102 = &HE2     ' <> o \ en teclados 102 teclas

' Paréntesis
' ( ? Shift + "9" (VK_9 = &H39)
' ) ? Shift + "0" (VK_0 = &H30)

' Otros símbolos con Shift:
' ! ? Shift + "1" (VK_1 = &H31)
' @ ? Shift + "2" (VK_2 = &H32)
' # ? Shift + "3" (VK_3 = &H33)
' $ ? Shift + "4" (VK_4 = &H34)
' % ? Shift + "5" (VK_5 = &H35)
' ^ ? Shift + "6" (VK_6 = &H36)
' & ? Shift + "7" (VK_7 = &H37)
' * ? Shift + "8" (VK_8 = &H38) o VK_MULTIPLY


'Declaración de ShellExecute
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal _
                        lpFile As String, ByVal lpParameters As String, ByVal _
                        lpDirectory As String, ByVal nShowCmd As Long) As Long
                        
                        
Private Declare Function CreateProcess Lib "kernel32" _
    Alias "CreateProcessA" (ByVal lpApplicationName As String, _
    ByVal lpCommandLine As String, lpProcessAttributes As Any, _
    lpThreadAttributes As Any, ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, ByVal lpEnvironment As String, _
    ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
                       
                        
Private Const SW_SHOWNORMAL = 1

Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long

'Minimiza una ventana
Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

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
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' función SendMessage
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
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
Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

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


Private Type HARDWAREINPUT
  uMsg As Long
  wParamL As Integer
  wParamH As Integer
End Type
Private Type GENERALINPUT
  dwType As Long
  xi(0 To 23) As Byte
End Type


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

Const VK_CONTROL = &H11
Const VK_V = &H56

Public HwndPadre As Long
Public EnumChild As Long
Public TxtChild As String

' En un módulo .BAS
Public m_TargetPID As Long
Public m_FoundHwnd As Long

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" _
    (lpExecInfo As SHELLEXECUTEINFO) As Long
   
Private Declare Function GetProcessId Lib "kernel32" (ByVal hProcess As Long) As Long
     
Private Declare Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" ( _
    ByVal hProcess As Long, _
    ByVal uExitCode As Long) As Long

' Constantes de acceso
Private Const PROCESS_TERMINATE = &H1
     
' Variable global para guardar el layout previo
Private hklPrevio As Long
Private tidDestino As Long
Private Const KLF_ACTIVATE = &H1

Public Sub KillProcessByPID(ByVal PID As Long)
    Dim hProcess As Long
    
    ' Abrir el proceso con permiso de terminación
    hProcess = OpenProcess(PROCESS_TERMINATE, 0, PID)
    
    If hProcess <> 0 Then
        ' Terminar el proceso
        Call TerminateProcess(hProcess, 0)
        ' Cerrar el handle
        Call CloseHandle(hProcess)
    Else
        MsgBox "No se pudo abrir el proceso con PID " & PID
    End If
End Sub

Function ObtenerPID(hwndVentana As Long) As Long
    Dim PID As Long
    GetWindowThreadProcessId hwndVentana, PID
    ObtenerPID = PID
End Function

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim PID As Long
    GetWindowThreadProcessId hwnd, PID
    If PID = m_TargetPID And IsWindowVisible(hwnd) <> 0 Then
        m_FoundHwnd = hwnd
        EnumWindowsProc = 0   ' detener
    Else
        EnumWindowsProc = 1   ' continuar
    End If
End Function

Public Function FindWindowByProcessID(ByVal PID As Long) As Long
    m_TargetPID = PID
    m_FoundHwnd = 0
    EnumWindows AddressOf EnumWindowsProc, 0
    FindWindowByProcessID = m_FoundHwnd
End Function

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

Public Function EjecutarConPID(ByRef RutaExe As String, Optional ByVal Params As String = "") As Long
    Dim sei As SHELLEXECUTEINFO
    Dim ret As Long

    sei.cbSize = Len(sei)
    sei.fMask = SEE_MASK_NOCLOSEPROCESS
    sei.hwnd = 0
    sei.lpVerb = "open"
    sei.lpFile = RutaExe
    sei.lpParameters = Params
    sei.lpDirectory = vbNullString
    sei.nShow = SW_SHOWNORMAL

    ret = ShellExecuteEx(sei)

    If ret <> 0 Then
        ' Obtener el PID real
        Dim pi As Long
        pi = GetProcessId(sei.hProcess)
        EjecutarConPID = pi
        CloseHandle sei.hProcess
    Else
        MsgBox "Error al ejecutar: " & Err.LastDllError, vbCritical
        EjecutarConPID = 0
    End If
End Function

'Funcion de lectura manual de un archivo INI (respeta los espacios de cadenas)
Public Function LeeINI(ByVal Filename As String, ByVal Section As String, ByVal Key_Name As String, Optional ByVal Default As String) As String
    On Error GoTo ErrOut
    Dim Linea As String
    Dim SecActual As String
    Dim f As Integer

    'Comprobamos que el archivo existe:
    If Dir(Filename) = "" Then Err.Raise 53

    f = FreeFile
    Open Filename For Input As #f

    SecActual = ""
    Do While Not EOF(f)
        Line Input #f, Linea

        'Detectar sección
        If Left$(Trim$(Linea), 1) = "[" And Right$(Trim$(Linea), 1) = "]" Then
            SecActual = Mid$(Linea, 2, Len(Linea) - 2)
        ElseIf SecActual = Section Then
            'Buscar clave dentro de la sección actual
            If InStr(1, Linea, Key_Name & "=", vbTextCompare) = 1 Then
                'Extraer valor tal cual, incluyendo espacios
                LeeINI = Mid$(Linea, Len(Key_Name) + 2)
                Close #f
                Exit Function
            End If
        End If
    Loop

    Close #f
    LeeINI = Default
    Exit Function

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
    Dim KINPUT As KEYBDINPUT
    KINPUT.wVk = bKey  'the key we're going to press
    KINPUT.dwFlags = 0 'press the key
    'copy the structure into the input array's buffer.
    GInput(0).dwType = INPUT_KEYBOARD   ' keyboard input
    CopyMemory GInput(0).xi(0), KINPUT, Len(KINPUT)
    'do the same as above, but for releasing the key
    KINPUT.wVk = bKey  ' the key we're going to realease
    KINPUT.dwFlags = KEYEVENTF_KEYUP  ' release the key
    GInput(1).dwType = INPUT_KEYBOARD  ' keyboard input
    CopyMemory GInput(1).xi(0), KINPUT, Len(KINPUT)
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

Public Sub CopiarAlPortapapeles(ByVal Texto As String)
    Clipboard.Clear
    Clipboard.SetText Texto
End Sub

Public Sub SimularCtrlV()
    ' Presionar Ctrl
    keybd_event VK_CONTROL, 0, 0, 0
    
    ' Presionar y soltar V
    EnviarTecla VK_V
    
    ' Soltar Ctrl
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
End Sub

Public Sub PegarTextoHwnd(ByVal hwnd As Long, ByVal Texto As String, ByRef TmpTextBox As TextBox)
TmpTextBox.Text = Texto
If TmpTextBox.SelLength = 0 Then SendMessage TmpTextBox.hwnd, EM_SETSEL, 0&, -1&
SendMessage TmpTextBox.hwnd, WM_COPY, 0&, 0&  'Copio al portapapeles
SendMessage hwnd, WM_PASTE, 0&, 0&  'Pego en el destino
End Sub

Public Sub EnviarTecla(vk As Integer, Optional shiftNeeded As Boolean = False)
    If shiftNeeded Then keybd_event VK_SHIFT, 0, 0, 0
    keybd_event vk, 0, 0, 0
    
    keybd_event vk, 0, KEYEVENTF_KEYUP, 0
    If shiftNeeded Then keybd_event VK_SHIFT, 0, KEYEVENTF_KEYUP, 0
End Sub

Public Sub ForzarMinusculas()
    ' Si Bloq Mayús está activo, lo desactiva
    If (GetKeyState(VK_CAPITAL) And 1) <> 0 Then
        ' Simula pulsar Bloq Mayús
        keybd_event VK_CAPITAL, 0, 0, 0
        keybd_event VK_CAPITAL, 0, KEYEVENTF_KEYUP, 0
    End If
End Sub

Public Sub PegarTexto(ByVal hwnd As Long, ByVal Texto As String, IntervaloTecla As Long)
    SetForegroundWindow hwnd
    
    Dim i As Long, c As String
    For i = 1 To Len(Texto)
        c = Mid$(Texto, i, 1)
        
    Select Case c
        ' Números
        Case "0" To "9"
            EnviarTecla Asc(c)
    
        ' Espacio y Enter
        Case " "
            EnviarTecla VK_SPACE
        Case vbCr, vbLf
            EnviarTecla VK_RETURN
    
        Case ":"
            EnviarTecla VK_OEM_1, True   ' Shift + ; ? :
        Case ";"
            EnviarTecla VK_OEM_1, False  ' ;
    
        Case "\"
            EnviarTecla VK_OEM_5, False  ' \
        Case "|"
            EnviarTecla VK_OEM_5, True   ' Shift + \ ? |
    
        Case "="
            EnviarTecla VK_OEM_PLUS, False
        Case "+"
            EnviarTecla VK_OEM_PLUS, True
    
        Case "/"
            EnviarTecla VK_OEM_2, False
        Case "?"
            EnviarTecla VK_OEM_2, True
    
        Case """"
            EnviarTecla VK_OEM_7, True   ' Shift + ' ? "
        Case "'"
            EnviarTecla VK_OEM_7, False
    
        Case "-"
            EnviarTecla VK_OEM_MINUS, False
        Case "_"
            EnviarTecla VK_OEM_MINUS, True
    
        Case ","
            EnviarTecla VK_OEM_COMMA, False
        Case "<"
            EnviarTecla VK_OEM_COMMA, True
    
        Case "."
            EnviarTecla VK_OEM_PERIOD, False
        Case ">"
            EnviarTecla VK_OEM_PERIOD, True
    
        Case "["
            EnviarTecla VK_OEM_4, False
        Case "{"
            EnviarTecla VK_OEM_4, True
    
        Case "]"
            EnviarTecla VK_OEM_6, False
        Case "}"
            EnviarTecla VK_OEM_6, True
    
        Case "`"
            EnviarTecla VK_OEM_3, False
        Case "~"
            EnviarTecla VK_OEM_3, True
    
        Case "*"
            EnviarTecla Asc("8"), True  ' Shift + 8 en US ? *
        
    Case Else
         If c Like "[A-Za-z]" Then
            Dim vk As Integer
            vk = Asc(UCase(c))   ' VK_A = 65, VK_B = 66, etc.
            If c = UCase(c) Then
                EnviarTecla vk, True   ' Mayúscula ? con Shift
            Else
                EnviarTecla vk, False  ' Minúscula ? sin Shift
            End If
        ElseIf c Like "[0-9]" Then
            Dim vkNum As Integer
            vkNum = Asc(c)             ' VK_0 = 48, VK_1 = 49, etc.
            EnviarTecla vkNum, False
        End If
    End Select

    Sleep IntervaloTecla
    Next i
    
End Sub

Sub EnviarUnicode(ByVal ch As String)
    Dim inp As KINPUT
    
    ' Pulsar la tecla
    inp.dwType = INPUT_KEYBOARD
    inp.ki.wVk = 0
    inp.ki.wScan = AscW(ch)
    inp.ki.dwFlags = KEYEVENTF_UNICODE
    SendInput 1, inp, Len(inp)
    
    ' Soltar la tecla
    inp.ki.dwFlags = KEYEVENTF_UNICODE Or KEYEVENTF_KEYUP
    SendInput 1, inp, Len(inp)
End Sub

Public Function HwndFromPID(ByVal PID As Long) As Long
    m_TargetPID = PID
    m_FoundHwnd = 0
    EnumWindows AddressOf EnumWindowsProc, 0
    HwndFromPID = m_FoundHwnd
End Function
Public Function TomarFoco(ByVal hwnd As Long) As Long
TomarFoco = SetFocus(hwnd)
End Function

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

