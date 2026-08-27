VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B389CD47-E20E-4D96-A4EC-576F2B1F43BF}#1.0#0"; "Hook-Menu-2.ocx"
Object = "{58D464EF-A3A5-471F-9E1E-5F5FE61295C6}#1.0#0"; "BotonXp.ocx"
Object = "{95D85F43-414D-432F-909E-2ED57BBC389C}#1.2#0"; "MCLHotkey.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   ClientHeight    =   8220
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14370
   Icon            =   "Asistente.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Asistente.frx":345A
   ScaleHeight     =   548
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   958
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Zinisoft Windows Handler"
   WindowState     =   2  'Maximized
   Begin MCLHotkey.VBHotKey VBHotKey1 
      Left            =   135
      Top             =   5715
      _ExtentX        =   794
      _ExtentY        =   794
      VKey            =   116
   End
   Begin BotonXp.BotónXp CmdNuevo 
      Height          =   375
      Left            =   765
      TabIndex        =   40
      Top             =   4905
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&Nuevo"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Asistente.frx":AE74
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   5265
      Top             =   2655
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox tmpTxt 
      Height          =   330
      Left            =   13770
      TabIndex        =   39
      Top             =   4905
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CheckBox ChkBloq 
      Caption         =   "Bloquear mouse y teclado (recomendado para secuencias largas)"
      Height          =   240
      Left            =   8460
      TabIndex        =   37
      Top             =   4995
      Width           =   5010
   End
   Begin VB.Frame Frame3 
      Caption         =   "Directorio raíz de la aplicación"
      Height          =   1545
      Left            =   765
      TabIndex        =   32
      Top             =   7965
      Width           =   7530
      Begin VB.TextBox txtRuta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1710
         MaxLength       =   2
         TabIndex        =   36
         Top             =   1110
         Width           =   465
      End
      Begin VB.OptionButton OptRuta3 
         Caption         =   "Otra ubicación:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   270
         TabIndex        =   35
         Top             =   1155
         Width           =   1455
      End
      Begin VB.OptionButton optRuta2 
         Caption         =   "Utilizar el directorio raíz del sistema operativo"
         Enabled         =   0   'False
         Height          =   195
         Left            =   270
         TabIndex        =   34
         Top             =   840
         Width           =   3570
      End
      Begin VB.OptionButton optRuta1 
         Caption         =   $"Asistente.frx":AE90
         Enabled         =   0   'False
         Height          =   420
         Left            =   270
         TabIndex        =   33
         Top             =   315
         Width           =   6900
      End
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   5580
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "sca"
      Filter          =   "Archivos de secuencia de comandos automatizados (*.sca)|*.sca"
   End
   Begin VB.Frame Frame2 
      Caption         =   "Observador de clases:"
      Height          =   2895
      Left            =   8415
      TabIndex        =   4
      Top             =   5370
      Width           =   5505
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   4275
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2385
         Width           =   825
      End
      Begin VB.TextBox txtY 
         Height          =   285
         Left            =   4275
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1890
         Width           =   825
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2385
         Width           =   825
      End
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1890
         Width           =   825
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtNC 
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txthwndp 
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   1230
      End
      Begin VB.TextBox txthwnd 
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   3015
         TabIndex        =   29
         Top             =   2430
         Width           =   150
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   4095
         TabIndex        =   28
         Top             =   2430
         Width           =   150
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Posición relativa dentro de la región:"
         Height          =   195
         Left            =   330
         TabIndex        =   26
         Top             =   2430
         Width           =   2565
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Posición general del cursor:"
         Height          =   195
         Left            =   945
         TabIndex        =   25
         Top             =   1935
         Width           =   1950
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   4095
         TabIndex        =   14
         Top             =   1935
         Width           =   150
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   3015
         TabIndex        =   13
         Top             =   1935
         Width           =   150
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Caption:"
         Height          =   195
         Left            =   795
         TabIndex        =   11
         Top             =   1485
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de clase:"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   1125
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hwnd parent:"
         Height          =   195
         Left            =   450
         TabIndex        =   7
         Top             =   765
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hwnd:"
         Height          =   195
         Left            =   945
         TabIndex        =   5
         Top             =   405
         Width           =   465
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4905
      Top             =   7290
   End
   Begin VB.Frame Frame1 
      Caption         =   "Eventos"
      Height          =   2490
      Left            =   765
      TabIndex        =   0
      Top             =   5370
      Width           =   7530
      Begin VB.OptionButton OptEnter 
         Caption         =   "Enviar ENTER"
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   405
         TabIndex        =   44
         Top             =   2160
         Width           =   1500
      End
      Begin VB.CheckBox ChkXY 
         Caption         =   "En coordenadas específicas dentro de la ventana:"
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   810
         TabIndex        =   38
         Top             =   690
         Width           =   3930
      End
      Begin VB.OptionButton OptTexto 
         Caption         =   "Escribir Texto:"
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   405
         TabIndex        =   24
         Top             =   1605
         Width           =   1365
      End
      Begin VB.TextBox TxtCy 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6075
         MaxLength       =   4
         TabIndex        =   23
         Top             =   720
         Width           =   780
      End
      Begin VB.TextBox txtCx 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4950
         MaxLength       =   4
         TabIndex        =   22
         Top             =   720
         Width           =   780
      End
      Begin VB.OptionButton OptTab 
         Caption         =   "Enviar TAB"
         Enabled         =   0   'False
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   405
         TabIndex        =   19
         Top             =   1890
         Width           =   1320
      End
      Begin VB.TextBox txtTexto 
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Top             =   1560
         Width           =   5550
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Doble click"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4905
         TabIndex        =   17
         Top             =   315
         Width           =   1725
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Enviar cierre de ventana"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   405
         TabIndex        =   16
         Top             =   1260
         Width           =   2040
      End
      Begin VB.OptionButton OptClick2 
         Caption         =   "Click Derecho"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3180
         TabIndex        =   2
         Top             =   315
         Width           =   1725
      End
      Begin VB.OptionButton OptClick1 
         Caption         =   "Click"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1455
         TabIndex        =   1
         Top             =   315
         Value           =   -1  'True
         Width           =   1725
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5895
         TabIndex        =   21
         Top             =   765
         Width           =   150
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4770
         TabIndex        =   20
         Top             =   765
         Width           =   150
      End
   End
   Begin BotonXp.BotónXp CmdInsertar 
      Height          =   375
      Left            =   2490
      TabIndex        =   41
      Top             =   4905
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&Instertar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Asistente.frx":AF4B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BotonXp.BotónXp CmdEliminar 
      Height          =   375
      Left            =   4215
      TabIndex        =   42
      Top             =   4905
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&Eliminar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Asistente.frx":AF67
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BotonXp.BotónXp CmdSec 
      Height          =   375
      Left            =   5940
      TabIndex        =   43
      Top             =   4905
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   661
      BTYPE           =   14
      TX              =   "&Comenzar secuencia"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Asistente.frx":AF83
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataGridLib.DataGrid dgEventos 
      Align           =   1  'Align Top
      Height          =   4830
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   8520
      _Version        =   393216
      BackColor       =   -2147483624
      HeadLines       =   1
      RowHeight       =   18
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   12915
      Picture         =   "Asistente.frx":AF9F
      Stretch         =   -1  'True
      Top             =   10170
      Width           =   2265
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuNuevo 
         Caption         =   "&Nuevo proyecto"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuAP 
         Caption         =   "&Abrir proyecto"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuGP 
         Caption         =   "&Guardar proyecto"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuGPC 
         Caption         =   "&Guardar proyecto como..."
      End
      Begin VB.Menu raya1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTmp 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTmp 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTmp 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTmp 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu raya2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m As Integer 'Menu recientes
Dim Archivo As String 'Archivo de proyecto abierto
Dim PasoHecho As Boolean 'Bandera de paso completado en la secuencia
Dim hwndChild As Long
Dim pos As Boolean, Bloquear As Boolean
Dim rsSucesos As New Recordset, t As Recordset
Attribute rsSucesos.VB_VarHelpID = -1
Dim WithEvents EnumWin As clsEnum
Attribute EnumWin.VB_VarHelpID = -1

Private Sub ChkBloq_Click()
Bloquear = CBool(ChkBloq.Value)
End Sub

Private Sub ChkXY_Click()
If ChkXY.Value = 0 Then
txtCx.Text = "": TxtCy.Text = ""
rsSucesos.Update
rsSucesos("clickposx") = 0
rsSucesos("clickposy") = 0
rsSucesos.UpdateBatch adAffectCurrent
End If
End Sub

Private Sub CmdEliminar_Click()
Dim ap As Double
ap = rsSucesos.AbsolutePosition
'Si saco uno del medio, reajustar el numero de pila
If rsSucesos.RecordCount - 1 = 0 Or rsSucesos.AbsolutePosition <= 0 Or ap <= 0 Then Exit Sub
rsSucesos.Delete
rsSucesos.UpdateBatch
If rsSucesos.RecordCount = 0 Then
 CmdEliminar.Enabled = False
 Exit Sub
End If
Dim i As Integer
rsSucesos.MoveFirst
For i = 1 To rsSucesos.RecordCount
rsSucesos.Update "id", i
rsSucesos.UpdateBatch adAffectCurrent
rsSucesos.MoveNext
Next
If Not rsSucesos.BOF And rsSucesos.RecordCount <> 1 Then rsSucesos.AbsolutePosition = CDbl(ap + IIf(ap = 1, 1, -1))
End Sub

Private Sub CmdInsertar_Click()
Dim bmk As Long, i As Long, k As Long
Timer1.Enabled = False
Set dgEventos.DataSource = Nothing
With rsSucesos
bmk = .AbsolutePosition
.UpdateBatch
Set t = New Recordset
Call AppendCampos(t)
t.Open
.MoveFirst
i = 1
Do While Not .EOF
If i = bmk Then 'Dejo un lugar para el nuevo registro
t.AddNew "id", i
t.Update
i = i + 1 'Hacer un salto de ID para continuar la numeración
End If
If i <> bmk Then 'Seguir cargando los demás registros por debajo del nuevo
t.AddNew
t(0) = i
For k = 1 To .Fields.Count - 1
t(k) = .Fields(k) 'Continuar con el volcado de los datos en los nuevos registros
Next
t.Update
End If
.MoveNext
i = i + 1
Loop
.Close
Set rsSucesos = Nothing
Set rsSucesos = New Recordset
Call AppendCampos(rsSucesos)
Set rsSucesos = t.Clone
.Sort = "id"
t.Close
Set t = Nothing
Set dgEventos.DataSource = rsSucesos
BindGrid
rsSucesos.AbsolutePosition = bmk
Timer1.Enabled = True
End With
End Sub

Private Sub cmdNuevo_Click()
AgregarNuevo (rsSucesos.RecordCount + 1)
rsSucesos.MoveLast
End Sub

Private Sub AgregarNuevo(ByVal id As Integer)
Dim ap As Integer, idant As String
rsSucesos.AddNew
rsSucesos("id") = id
rsSucesos("evento") = 1 'Evento por defecto: click izquierdo
rsSucesos.Update
If rsSucesos.AbsolutePosition = 1 Then Exit Sub
'Copiar por defecto el titulo de ventana ya cargado
If rsSucesos.RecordCount > 1 Then
rsSucesos.MovePrevious
idant = rsSucesos("textoventana")
rsSucesos.MoveNext
rsSucesos.Update "textoventana", idant
rsSucesos.UpdateBatch adAffectCurrent
End If
optRuta1.Value = True
End Sub

Private Function VerificarRutas() As Boolean
Dim r As Long
Dim uo As String
rsSucesos.MoveFirst
If Val(rsSucesos("unidadorigen")) = 1 Then uo = Left(App.Path, 2)
If Val(rsSucesos("unidadorigen")) = 2 Then uo = ObtenerWinDir
If Val(rsSucesos("unidadorigen")) = 0 Then uo = rsSucesos("unidadorigen")
VerificarRutas = True
For r = 1 To rsSucesos.RecordCount
If Trim(rsSucesos("Ruta")) <> "" Then
If Dir(uo + "\" + rsSucesos("ruta")) = "" Then
MsgBox "No se ha encontrado la ruta de acceso para " + rsSucesos("ruta") + vbCrLf + _
"Verifique el directorio raíz y la ruta de acceso para esta aplicación", vbCritical
VerificarRutas = False
End If
End If
rsSucesos.MoveNext
Next
End Function

Private Sub GetWindowsChild(ByVal HwndPadre As Long)
EnumChild = 0
EnumChildWindows HwndPadre, AddressOf EnumChildProc, ByVal 0&
If Not PasoHecho And EnumChild <> 0 Then
If IsWindowEnabled(EnumChild) And IsWindow(EnumChild) And IsWindowVisible(EnumChild) Then
'Enviar click en función del handle y el caption (textosubclase)
If rsSucesos("evento") = 1 Or rsSucesos("evento") = 2 Or rsSucesos("evento") = 3 Then
EnumWin.SetActiveWindows HwndPadre
EnviarClick EnumChild, rsSucesos("evento")
PasoHecho = True
End If
If Trim(rsSucesos("texto")) <> "" Then
PegarTextoHwnd EnumChild, Trim(rsSucesos("texto")), tmpTxt
PasoHecho = True
End If
End If
End If
End Sub

Private Sub cmdSec_Click()
'Verificar todas las rutas de las aplicaciones a  ejecutar
If Not VerificarRutas Then Exit Sub
Timer1.Enabled = False
Dim TextoVentana As String
Dim TextoSubClase As String
Dim Ruta As String, unidad As String
Dim px As Long, py As Long, pasos As Long
Dim v As RECT
Dim fwc As Long 'Handle ventana actual
rsSucesos.Sort = "id asc"
BlockInput Bloquear 'Bloquear mouse y teclado
rsSucesos.MoveFirst

For pasos = 1 To rsSucesos.RecordCount
px = rsSucesos("clickposx")
py = rsSucesos("clickposy")
Ruta = Trim(rsSucesos("ruta"))
unidad = rsSucesos("unidadorigen")
TextoVentana = Trim(LCase(rsSucesos("TextoVentana")))
TextoSubClase = Trim(LCase(rsSucesos("TextoSubClase")))

If Ruta <> "" Then
If Val(unidad) = 1 Then Ejecutar Me.hwnd, Left(App.Path, 2) + Ruta
If Val(unidad) = 2 Then Ejecutar Me.hwnd, ObtenerWinDir + "\" + Ruta
If Val(unidad) = 0 Then Ejecutar Me.hwnd, unidad + "\" + Ruta
End If

PasoHecho = False
TxtChild = TextoSubClase

Sleep (100)

Do
fwc = EnumWin.FindWindowByCaption(TextoVentana)

If TextoSubClase <> Empty And IsWindowEnabled(fwc) And _
IsWindowVisible(fwc) And IsWindow(fwc) Then

EnumWin.SetActiveWindows fwc

If Trim(rsSucesos("texto")) <> "" Then
PegarTexto fwc, Trim(rsSucesos("texto"))
PasoHecho = True
End If

If rsSucesos("evento") = 4 Then 'Cierre de ventana
EnumWin.CloseWindowByHandle fwc
PasoHecho = True
ElseIf rsSucesos("evento") = 5 And IIf(EnumChild = 0, True, CBool(IsWindowEnabled(EnumChild))) Then
Pulsar_Tecla (9) 'TAB
PasoHecho = True
ElseIf rsSucesos("evento") = 6 And IIf(EnumChild = 0, True, CBool(IsWindowEnabled(EnumChild))) Then
Pulsar_Tecla (13)  'Enter
PasoHecho = True
End If

If px <> 0 And py <> 0 And IIf(EnumChild = 0, True, CBool(IsWindowEnabled(EnumChild))) Then
'Enviar click "común" en función de la posición absoluta del cursor
EnumWin.SetActiveWindows fwc
Call GetWindowRect(fwc, v)
If rsSucesos("evento") = 1 Then EnviarClickXY px + v.Left, py + v.Top, eLeftClick
If rsSucesos("evento") = 2 Then EnviarClickXY px + v.Left, py + v.Top, eRightClick
If rsSucesos("evento") = 3 Then EnviarClickXY px + v.Left, py + v.Top, eDoubleLeft
PasoHecho = True
End If

If px = 0 And py = 0 Then  'Todo en función del handle
Call GetWindowsChild(fwc)
'Debug.Print pasos, TxtChild, EnumChild
End If

End If
'DoEvents
Loop Until PasoHecho = True

rsSucesos.MoveNext
DoEvents
Next
BlockInput False 'Una vez terminado desbloquear mouse y teclado
Timer1.Enabled = True
End Sub

Private Sub dgEventos_ColEdit(ByVal ColIndex As Integer)
If ColIndex = 2 Then
optRuta1.Enabled = True
optRuta2.Enabled = True
OptRuta3.Enabled = True
End If
End Sub

Private Sub dgEventos_HeadClick(ByVal ColIndex As Integer)
OrdenarDataGrid ColIndex, rsSucesos, dgEventos
End Sub

Private Sub dgEventos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 117 Then SendKeys ("^(v)")
End Sub

Private Sub dgEventos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If rsSucesos.RecordCount > 0 Then CmdEliminar.Enabled = True

If Not rsSucesos.EOF And Not rsSucesos.BOF Then
If Val(rsSucesos("unidadorigen")) = 1 Then optRuta1.Value = True
If Val(rsSucesos("unidadorigen")) = 2 Then optRuta2.Value = True
If Val(rsSucesos("unidadorigen")) = 0 Then
    OptRuta3.Value = True
    txtRuta.Text = rsSucesos("unidadorigen")
    txtRuta.Tag = txtRuta.Text
End If

If rsSucesos("Evento") = 1 Then OptClick1.Value = True
If rsSucesos("Evento") = 2 Then OptClick2.Value = True
If rsSucesos("Evento") = 3 Then Option2.Value = True
If rsSucesos("Evento") = 4 Then Option1.Value = True
If rsSucesos("Evento") = 5 Then OptTab.Value = True
If rsSucesos("Evento") = 6 Then OptEnter.Value = True

If Trim(rsSucesos("texto")) <> "" Then
txtTexto.Text = Trim(rsSucesos("texto"))
OptTexto.Value = True
Else
txtTexto.Text = ""
End If

If rsSucesos("clickposx") <> 0 And rsSucesos("clickposy") <> 0 Then
txtCx.Text = rsSucesos("clickposx"): TxtCy.Text = rsSucesos("clickposy")
ChkXY.Value = 1
Else
ChkXY.Value = 0
txtCx.Text = "": TxtCy.Text = ""
End If

End If
End Sub

Private Sub Form_Load()
Set EnumWin = New clsEnum
OptClick1.Value = False
Call AbrirRecordset
BindGrid
Me.Caption = Me.Tag
Call CargarMenuRecientes
IniciarScroll dgEventos

Dim Ctl As Control
For Each Ctl In Me.Controls
Dim cn As String
cn = TypeName(Ctl)
If cn <> "ctxHookMenu" And cn <> "CommonDialog" And cn <> "Timer" And cn <> "Menu" And cn <> "Image" And cn <> "VBHotKey" Then
Ctl.BackColor = Me.BackColor
End If
Next

'Argumentos de la línea de comandos
If Command <> "" Then
Me.Visible = False
Dim args() As String
args = Split(Command, "/")
Dim i As Long
Dim a As String
For i = 0 To UBound(args)
If Left(args(i), 1) = "a" Then a = Mid(args(i), 2, Len(args(i)))
If Trim(args(i)) = "b" Then Bloquear = True
Next
Call NuevoProyecto(Trim(a))
Call cmdSec_Click
Unload Me
End If

End Sub

Private Sub CargarMenuRecientes()
Dim i As Integer
Dim e As String
For i = 1 To 4
e = LeeINI(App.Path + "\config.ini", "MENU", "RECIENTE" + CStr(i))
If Trim(Dir(e)) <> "" And Trim(e) <> "" Then
m = m + 1
    raya1.Visible = True
    mnuTmp(m).Caption = CStr(m) + " " + e
    mnuTmp(m).Visible = True
End If
Next
End Sub

Private Sub AbrirRecordset()
Set rsSucesos = New Recordset
Call AppendCampos(rsSucesos)
If rsSucesos.State = adStateOpen Then rsSucesos.Close
rsSucesos.Open Archivo, , adOpenDynamic, adLockBatchOptimistic
Set dgEventos.DataSource = rsSucesos
If Trim(Archivo) <> "" Then rsSucesos.Requery
End Sub

Private Sub AppendCampos(ByRef rs As Recordset)
With rs.Fields
rs.CursorLocation = adUseClient
.Append "Id", adInteger
.Append "UnidadOrigen", adVarChar, 2
.Append "Ruta", adVarChar, 100
.Append "TextoVentana", adVarChar, 100
.Append "TextoSubClase", adVarChar, 50
.Append "Evento", adInteger
.Append "Texto", adVarChar, 50
.Append "ClickPosX", adInteger
.Append "ClickPosY", adInteger
End With
End Sub

Private Sub BindGrid()
With dgEventos
Dim c As Long
For c = 0 To rsSucesos.Fields.Count - 1
.Columns(c).DataField = rsSucesos.Fields(c).Name
Next
.Columns("id").Locked = True
.Columns("id").Width = 25
.Columns("Ruta").Caption = "Ruta de la aplicación (sin directorio raíz)"
.Columns(1).Width = 320
.Columns(2).Width = 270
.Columns(3).Width = 350
.Columns("UnidadOrigen").Visible = False
.Columns("Evento").Visible = False
.Columns("Texto").Visible = False
.Columns("ClickPosX").Visible = False
.Columns("ClickPosY").Visible = False
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
VBHotKey1.StopHotkey
Set EnumWin = Nothing
DetenerScroll dgEventos
Set dgEventos.DataSource = Nothing
rsSucesos.Close
Set rsSucesos = Nothing
GrabaINI App.Path + "\config.ini", "MENU", "RECIENTE1", Mid(mnuTmp(1).Caption, 3, Len(mnuTmp(1).Caption))
GrabaINI App.Path + "\config.ini", "MENU", "RECIENTE2", Mid(mnuTmp(2).Caption, 3, Len(mnuTmp(2).Caption))
GrabaINI App.Path + "\config.ini", "MENU", "RECIENTE3", Mid(mnuTmp(3).Caption, 3, Len(mnuTmp(3).Caption))
GrabaINI App.Path + "\config.ini", "MENU", "RECIENTE4", Mid(mnuTmp(4).Caption, 3, Len(mnuTmp(4).Caption))
End
End Sub

Private Sub mnuAP_Click()
cdg.CancelError = True
cdg.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
On Error Resume Next
cdg.ShowOpen
If Err.Number = 32755 Then Exit Sub
If cdg.Filename <> "" Then
Archivo = cdg.Filename
AbrirRecordset
BindGrid
Me.Caption = Me.Tag & " - [" & Archivo & "]"
Call AgregarMenuReciente
End If
End Sub

Private Sub AgregarMenuReciente() 'Agrega los menus para los archivos recientes
Dim Existe As Boolean
Dim i As Integer
Dim c As String
For i = 1 To 4
c = LCase(Mid(Trim(mnuTmp(i).Caption), 3, Len(Trim(mnuTmp(i).Caption))))
If c = LCase(Archivo) Then Existe = True: Exit For
Next
If Not Existe Then
m = m + 1
If m > 4 Then m = 1
c = CStr(m) + " " + Archivo
raya1.Visible = True
mnuTmp(m).Caption = c
mnuTmp(m).Visible = True
End If
End Sub

Private Sub mnuGP_Click()
If Archivo <> "" And Dir(Archivo) <> "" Then
Kill (Archivo)
rsSucesos.Save Archivo
rsSucesos.UpdateBatch adAffectAllChapters
Me.Caption = Me.Tag & " - [" & Archivo & "]"
End If
End Sub

Private Sub mnuGPC_Click()
cdg.CancelError = True
cdg.flags = cdlOFNOverwritePrompt
On Error Resume Next
cdg.ShowSave
If Err.Number = 32755 Then Exit Sub
If cdg.Filename <> "" Then
Archivo = cdg.Filename
If Dir(Archivo) <> "" Then Kill (Archivo)
rsSucesos.Save Archivo
rsSucesos.UpdateBatch adAffectAllChapters
Me.Caption = Me.Tag & " - [" & Archivo & "]"
End If
End Sub

Private Sub mnuTmp_Click(Index As Integer)
Archivo = Mid(mnuTmp(Index).Caption, 3, Len(mnuTmp(Index).Caption))
If Trim(Dir(Archivo)) = "" Then
MsgBox "No se encuentra el archivo " + Archivo, vbCritical
Else
NuevoProyecto (Archivo)
End If
End Sub

Private Sub NuevoProyecto(ByVal RutaArchivo As String)
If Trim(Dir(RutaArchivo)) = "" Then
MsgBox RutaArchivo + vbCrLf + "El archivo no existe o no es válido...", vbCritical
End
End If
On Error GoTo Error
txtCx.Text = ""
TxtCy.Text = ""
txtTexto.Text = ""
rsSucesos.Close
Set rsSucesos = Nothing
Set dgEventos.DataSource = Nothing
Archivo = RutaArchivo
AbrirRecordset
BindGrid
Me.Caption = IIf(Trim(RutaArchivo = ""), Me.Tag, Me.Tag & " - [" & Archivo & "]")
CmdNuevo.Enabled = True
mnuGP.Enabled = False

Exit Sub
Error:
MsgBox Err.Description + "  " + Archivo, vbCritical, Err.Number
End
End Sub

Private Sub mnuNuevo_Click()
NuevoProyecto ""
End Sub

Private Sub mnuSalir_Click()
Unload Me
End Sub

Private Sub OptClick1_Click()
    rsSucesos.Update "Evento", 1     'Click izquierdo
    rsSucesos.UpdateBatch adAffectCurrent
End Sub

Private Sub OptClick2_Click()
    rsSucesos.Update "Evento", 2 'Click derecho
    rsSucesos.UpdateBatch adAffectCurrent
End Sub

Private Sub Option2_Click()
    rsSucesos.Update "Evento", 3 'Doble click
    rsSucesos.UpdateBatch adAffectCurrent
End Sub

Private Sub Option1_Click()
    rsSucesos.Update "Evento", 4 'Cerrar ventana
    rsSucesos.UpdateBatch adAffectCurrent
End Sub

Private Sub optRuta1_Click()
    rsSucesos.Update "UnidadOrigen", "1"
    rsSucesos.UpdateBatch adAffectCurrent
End Sub

Private Sub optRuta2_Click()
    rsSucesos.Update "UnidadOrigen", "2"
    rsSucesos.UpdateBatch adAffectCurrent
End Sub

Private Sub OptRuta3_Click()
txtRuta.Enabled = True
End Sub

Private Sub txtRuta_Validate(Cancel As Boolean)
If txtRuta.Text = Empty Then
Cancel = True
MsgBox "Debe escribir una letra de unidad seguido de dos puntos, por ejemplo d:", vbExclamation
txtRuta.Text = txtRuta.Tag
Exit Sub
Else
 rsSucesos.Update "UnidadOrigen", Trim(txtRuta.Text)
 rsSucesos.UpdateBatch adAffectCurrent
End If
End Sub

Private Sub OptTab_Click()
    rsSucesos.Update "Evento", 5  'enviar tab
    rsSucesos.UpdateBatch adAffectCurrent
End Sub

Private Sub OptEnter_Click()
    rsSucesos.Update "Evento", 6 'Enviar enter
    rsSucesos.UpdateBatch adAffectCurrent
End Sub

Private Sub OptClickXY_Click()
txtCx.Enabled = True: TxtCy.Enabled = True
End Sub

Private Sub OptTexto_Click()
txtTexto.Enabled = True
End Sub

Private Sub txtCx_Validate(Cancel As Boolean)
    rsSucesos.Update "clickposx", CInt(Val(txtCx.Text))
    rsSucesos.UpdateBatch adAffectCurrent
End Sub

Private Sub txtCy_Validate(Cancel As Boolean)
    rsSucesos.Update "clickposy", CInt(Val(TxtCy.Text))
    rsSucesos.UpdateBatch adAffectCurrent
End Sub

Private Sub txtTexto_Validate(Cancel As Boolean)
    rsSucesos.Update "Texto", Trim(txtTexto.Text)  'Texto para escribir
    rsSucesos.UpdateBatch adAffectCurrent
End Sub

Private Sub Timer1_Timer()
pos = CBool(rsSucesos.AbsolutePosition > 0) And Not rsSucesos.EOF And Not rsSucesos.BOF
mnuGP.Enabled = rsSucesos.RecordCount > 0 And Trim(Archivo) <> ""
CmdNuevo.Enabled = pos Xor rsSucesos.RecordCount = 0
CmdInsertar.Enabled = pos And rsSucesos.RecordCount > 0
OptClick1.Enabled = pos: OptClick2.Enabled = pos
Option2.Enabled = pos: Option1.Enabled = pos
OptTexto.Enabled = pos: OptTab.Enabled = pos
OptEnter.Enabled = pos
CmdEliminar.Enabled = pos
txtTexto.Enabled = pos And OptTexto.Value
ChkXY.Enabled = (OptClick1.Value Or OptClick2.Value Or Option2.Value) And pos And rsSucesos.RecordCount > 0
txtCx.Enabled = ChkXY.Value And pos
TxtCy.Enabled = ChkXY.Value And pos

If Not rsSucesos.EOF And Not rsSucesos.BOF Then
optRuta1.Enabled = (pos And rsSucesos("ruta") <> "") Or (pos And dgEventos.Columns(2) <> "")
optRuta2.Enabled = (pos And rsSucesos("ruta") <> "") Or (pos And dgEventos.Columns(2) <> "")
OptRuta3.Enabled = (pos And rsSucesos("ruta") <> "") Or (pos And dgEventos.Columns(2) <> "")
txtRuta.Enabled = rsSucesos("ruta") <> "" And OptRuta3.Value And pos
End If

CmdSec.Enabled = rsSucesos.RecordCount > 0
  
  Dim ret As Long, handle As Long, hParent
  Dim Cor As POINTAPI
  Dim buffer As String
  Dim ClassName As String
  Dim Caption_Ventana As String
    
    'Obtiene la coordenada del Mouse
    ret = GetCursorPos(Cor)
    'Recuperamos el HWND de la ventana asociada a esa coordenada
    handle = WindowFromPoint(Cor.X, Cor.Y)
    
    'Handle de la ventana padre
    hParent = GetParent(handle)
    'Llenamos un Buffer
    ClassName = Space$(128)
    'Recupera el Classname y lo devuelve en el Buffer
    ret = GetClassName(handle, ClassName, 128)
    
    'Extraemos el nombre de la clase
    ClassName = LCase(Left$(ClassName, ret))
    
    'Cantidad de caracteres del texto
    Caption_Ventana = EnumWin.GetWindowCaption(handle)
    
    txthwnd.Text = handle
    txthwndp.Text = hParent
    txtNC.Text = ClassName
    txtCaption.Text = Caption_Ventana
    txtX.Text = Cor.X
    txtY.Text = Cor.Y
    
    'Posición relativa del cursor dentro de la ventana
    Dim infoWin As RECT
    GetWindowRect GetAncestor(handle, 2), infoWin
    'Posición relativa dentro de la ventana principal
    Text1.Text = Cor.Y - infoWin.Top
    Text2.Text = Cor.X - infoWin.Left
End Sub

Private Sub VBHotKey1_HotkeyPressed()
Clipboard.Clear
'If Trim(txtCaption.Text) = "" Then Call EnviarClickXY(CLng(Val(txtX.Text)), CLng(Val(txtY.Text)), eLeftClick)
Clipboard.SetText txtCaption.Text
End Sub
