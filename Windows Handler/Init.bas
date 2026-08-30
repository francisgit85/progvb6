Attribute VB_Name = "Init"
Option Explicit

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOWMINIMIZED = 2
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Global CargadoPorParametros As Boolean

Public Sub Main()
    Dim args() As String
    Dim sFile As String
    Dim nVeces As Long, i As Long
    Dim tiempoEspera As Long
    
    args = Split(Trim$(Command$), " ")
    
    If UBound(args) < 0 Then
        ' Sin parámetros ? modo interactivo
        Form1.Show
        Exit Sub
    End If
   
    sFile = args(0)
    nVeces = 1
    tiempoEspera = 0
    
    ' Parsear segundo y tercer parámetro si existen
    If UBound(args) >= 1 Then
        If Left$(args(1), 2) = "/x" Then
            nVeces = Val(Mid$(args(1), 3))
        ElseIf InStr(args(1), "veces:") > 0 Then
            nVeces = Val(Mid$(args(1), InStr(args(1), ":") + 1))
        End If
    End If
    
    If UBound(args) >= 2 Then
        If Left$(args(2), 2) = "/t" Then
            tiempoEspera = Val(Mid$(args(2), 3))
        ElseIf InStr(args(2), "tiempoEspera:") > 0 Then
            tiempoEspera = Val(Mid$(args(2), InStr(args(2), ":") + 1))
        End If
    End If
    
    ' Ejecutar secuencia N veces
    If LCase$(Right$(sFile, 4)) = ".sca" Then
        
        CargadoPorParametros = True
        Load Form1

        ' Mostrar minimizado
        ShowWindow Form1.hwnd, SW_SHOWMINIMIZED
    
        Form1.AbrirProyecto App.Path & "\" & sFile
        
        For i = 1 To nVeces
            Form1.CmdSecClick
            
            ' Esperar entre ejecuciones si corresponde
            If tiempoEspera > 0 And i < nVeces Then
                Sleep tiempoEspera * 1000   ' convertir a milisegundos
            End If
        Next i
        
        End   ' cerrar al terminar
    End If
End Sub

