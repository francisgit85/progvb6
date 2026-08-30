Attribute VB_Name = "ModLayout"
Option Explicit

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
Private Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long

Private Const WM_INPUTLANGCHANGEREQUEST As Long = &H50
Private Const KLF_ACTIVATE As Long = 1

Public Const HKL_ENG_US As String = "00000409"
Public Const HKL_ESP_ES As String = "00000C0A" ' Español España
Public Const HKL_ESP_ES_TRA As String = "0000040A" ' Español España / Tradicional
Public Const HKL_ESP_LA As String = "0000080A" ' Español Latinoamérica

Public Function LayoutActualKLID() As String
    Dim buffer As String
    buffer = String$(9, vbNullChar) ' 8 dígitos + terminador

    If GetKeyboardLayoutName(buffer) <> 0 Then
        LayoutActualKLID = Left$(buffer, 8)
    Else
        LayoutActualKLID = ""
    End If
End Function

Public Sub CambiarTecladoForzado(ByVal ID_Idioma As String)
    Dim hwndActivo As Long
    Dim processId As Long
    Dim targetThreadId As Long
    Dim currentThreadId As Long
    Dim hkl As Long
    
    hwndActivo = GetForegroundWindow()
    If hwndActivo = 0 Then Exit Sub
    
    ' Obtener IDs de los hilos de la ventana remota (cmd.exe) y de nuestra app VB6
    targetThreadId = GetWindowThreadProcessId(hwndActivo, processId)
    currentThreadId = GetCurrentThreadId()
    
    ' Cargar el Layout en el sistema
    hkl = LoadKeyboardLayout(ID_Idioma, KLF_ACTIVATE)
    
    ' 1. Enviar mensaje genérico de cambio de idioma a la ventana objetivo
    PostMessage hwndActivo, WM_INPUTLANGCHANGEREQUEST, 0, hkl
    
    ' 2. Vincular el hilo de VB6 con el hilo de la consola objetivo para transferir el estado de teclado
    If targetThreadId <> currentThreadId Then
        AttachThreadInput currentThreadId, targetThreadId, True
        ActivateKeyboardLayout hkl, KLF_ACTIVATE
        AttachThreadInput currentThreadId, targetThreadId, False
    Else
        ActivateKeyboardLayout hkl, KLF_ACTIVATE
    End If
End Sub
