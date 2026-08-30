Attribute VB_Name = "modAutomation"
Option Explicit

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function AllowSetForegroundWindow Lib "user32" (ByVal dwProcessId As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const SW_RESTORE As Long = 9
Private Const ASFW_ANY As Long = -1

Private m_ColHwndProcesados As New Collection
Private m_ColPreExistentes As New Collection
Private m_HwndEncontrado As Long

Public Sub LimpiarHistorialVentanas()
    Set m_ColHwndProcesados = New Collection
End Sub

Public Function EsHwndConocido(ByVal hwnd As Long) As Boolean
    On Error Resume Next
    Dim tmp As Variant
    tmp = m_ColHwndProcesados("H" & hwnd)
    EsHwndConocido = (Err.Number = 0)
    Err.Clear
End Function

Public Sub RegistrarHwnd(ByVal hwnd As Long)
    If hwnd <> 0 And Not EsHwndConocido(hwnd) Then
        On Error Resume Next
        m_ColHwndProcesados.Add hwnd, "H" & hwnd
        Err.Clear
    End If
End Sub

' --- CAPTURA DE VENTANAS PREVIAS ---
Private Function EnumPreExistentesProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    If IsWindowVisible(hwnd) <> 0 And GetWindowTextLength(hwnd) > 0 Then
        On Error Resume Next
        m_ColPreExistentes.Add hwnd, "H" & hwnd
        Err.Clear
    End If
    EnumPreExistentesProc = 1
End Function

Public Sub TomarSnapshotVentanas()
    Set m_ColPreExistentes = New Collection
    EnumWindows AddressOf EnumPreExistentesProc, 0
End Sub

' --- BÚSQUEDA DE LA NUEVA VENTANA APARECIDA ---
Private Function EnumNuevasProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    If IsWindowVisible(hwnd) <> 0 And GetWindowTextLength(hwnd) > 0 Then
        ' Verificar si esta ventana NO estaba en la lista de previas y tampoco procesada
        On Error Resume Next
        Dim tmp As Variant
        tmp = m_ColPreExistentes("H" & hwnd)
        
        ' Si tira error es porque NO existía antes -> Es la ventana NUEVA
        If Err.Number <> 0 And Not EsHwndConocido(hwnd) Then
            m_HwndEncontrado = hwnd
            Err.Clear
            EnumNuevasProc = 0 ' Detener búsqueda
            Exit Function
        End If
        Err.Clear
    End If
    EnumNuevasProc = 1
End Function

Public Function EsperarNuevaVentana(ByVal timeoutMs As Long) As Long
    Dim inicio As Long
    inicio = GetTickCount()
    m_HwndEncontrado = 0
    
    Do
        EnumWindows AddressOf EnumNuevasProc, 0
        
        If m_HwndEncontrado <> 0 Then
            EsperarNuevaVentana = m_HwndEncontrado
            Exit Function
        End If
        
        Sleep 50
        DoEvents
    Loop While (GetTickCount() - inicio) < timeoutMs
    
    EsperarNuevaVentana = 0
End Function

Public Function LanzarProceso(ByVal RutaExe As String, ByVal Parametros As String) As Long
    On Error Resume Next
    Dim cmdLine As String
    cmdLine = Trim$(RutaExe & " " & Parametros)
    
    If Len(cmdLine) > 0 Then
        AllowSetForegroundWindow ASFW_ANY
        LanzarProceso = Shell(cmdLine, vbNormalFocus)
    Else
        LanzarProceso = 0
    End If
End Function

Public Sub ActivarVentana(ByVal hwndTarget As Long)
    If hwndTarget = 0 Then Exit Sub
    
    Dim targetThreadId As Long
    Dim currentThreadId As Long
    Dim processId As Long

    AllowSetForegroundWindow ASFW_ANY

    targetThreadId = GetWindowThreadProcessId(hwndTarget, processId)
    currentThreadId = GetCurrentThreadId()

    ShowWindow hwndTarget, SW_RESTORE
    BringWindowToTop hwndTarget

    If targetThreadId <> currentThreadId And targetThreadId <> 0 Then
        AttachThreadInput currentThreadId, targetThreadId, True
        SetForegroundWindow hwndTarget
        AttachThreadInput currentThreadId, targetThreadId, False
    Else
        SetForegroundWindow hwndTarget
    End If
    
    Sleep 150 ' Dar tiempo a la GUI para asentar el foco activo
End Sub
