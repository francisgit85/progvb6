VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B389CD47-E20E-4D96-A4EC-576F2B1F43BF}#1.0#0"; "hook-menu-2.ocx"
Object = "{58D464EF-A3A5-471F-9E1E-5F5FE61295C6}#1.0#0"; "BotonXp.ocx"
Object = "{95D85F43-414D-432F-909E-2ED57BBC389C}#1.2#0"; "MCLHotkey.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   ClientHeight    =   11640
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   14205
   Icon            =   "Asistente.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   776
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   947
   Tag             =   "Zinisoft Windows Handler"
   WindowState     =   2  'Maximized
   Begin ComCtl2.UpDown udTiempoEspera 
      Height          =   375
      Left            =   22920
      TabIndex        =   57
      Top             =   10200
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   327681
      Max             =   999
      Enabled         =   -1  'True
   End
   Begin VB.ComboBox cboModoEnvio 
      Height          =   315
      Left            =   22800
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   9720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox cboEditor 
      Height          =   315
      Left            =   22800
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   9240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox chkEditor 
      Height          =   255
      Left            =   22800
      TabIndex        =   50
      Top             =   8880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtEditor 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   25800
      TabIndex        =   45
      Top             =   5040
      Visible         =   0   'False
      Width           =   1575
   End
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
      TabIndex        =   39
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
      MICON           =   "Asistente.frx":345A
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
      Left            =   22080
      Top             =   8880
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
      Left            =   24240
      TabIndex        =   38
      Top             =   5040
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CheckBox ChkBloq 
      Caption         =   "Bloquear mouse y teclado (recomendado para secuencias largas)"
      Height          =   240
      Left            =   7800
      TabIndex        =   36
      Top             =   6120
      Width           =   5010
   End
   Begin VB.Frame Frame3 
      Caption         =   "Directorio raíz de la aplicación"
      Height          =   1545
      Left            =   720
      TabIndex        =   31
      Top             =   11640
      Visible         =   0   'False
      Width           =   7530
      Begin VB.TextBox txtRuta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1710
         MaxLength       =   2
         TabIndex        =   35
         Top             =   1110
         Width           =   465
      End
      Begin VB.OptionButton OptRuta3 
         Caption         =   "Otra ubicación:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   270
         TabIndex        =   34
         Top             =   1155
         Width           =   1455
      End
      Begin VB.OptionButton optRuta2 
         Caption         =   "Utilizar el directorio raíz del sistema operativo"
         Enabled         =   0   'False
         Height          =   195
         Left            =   270
         TabIndex        =   33
         Top             =   840
         Width           =   3570
      End
      Begin VB.OptionButton optRuta1 
         Caption         =   $"Asistente.frx":3476
         Enabled         =   0   'False
         Height          =   420
         Left            =   270
         TabIndex        =   32
         Top             =   315
         Width           =   6900
      End
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   5580
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "sca"
      Filter          =   "Archivos de secuencia de comandos automatizados (*.sca)|*.sca"
   End
   Begin VB.Frame Frame2 
      Caption         =   "Observador de clases:"
      Height          =   2895
      Left            =   21840
      TabIndex        =   3
      Top             =   5400
      Width           =   5505
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   4275
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2385
         Width           =   825
      End
      Begin VB.TextBox txtY 
         Height          =   285
         Left            =   4275
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1890
         Width           =   825
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2385
         Width           =   825
      End
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1890
         Width           =   825
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtNC 
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox txthwndp 
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   1230
      End
      Begin VB.TextBox txthwnd 
         Height          =   285
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   3015
         TabIndex        =   28
         Top             =   2430
         Width           =   150
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   4095
         TabIndex        =   27
         Top             =   2430
         Width           =   150
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Posición relativa dentro de la región:"
         Height          =   195
         Left            =   330
         TabIndex        =   25
         Top             =   2430
         Width           =   2565
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Posición general del cursor:"
         Height          =   195
         Left            =   945
         TabIndex        =   24
         Top             =   1935
         Width           =   1950
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   4095
         TabIndex        =   13
         Top             =   1935
         Width           =   150
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   3015
         TabIndex        =   12
         Top             =   1935
         Width           =   150
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Caption:"
         Height          =   195
         Left            =   795
         TabIndex        =   10
         Top             =   1485
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de clase:"
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   1125
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hwnd parent:"
         Height          =   195
         Left            =   450
         TabIndex        =   6
         Top             =   765
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Hwnd:"
         Height          =   195
         Index           =   0
         Left            =   945
         TabIndex        =   4
         Top             =   405
         Width           =   465
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4905
      Top             =   7440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Eventos"
      Height          =   5850
      Left            =   720
      TabIndex        =   0
      Top             =   5400
      Width           =   20490
      Begin VB.TextBox txtTiempoEspera 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9480
         TabIndex        =   54
         Top             =   345
         Width           =   975
      End
      Begin VB.OptionButton OptTiempoEspera 
         Caption         =   "Esperar"
         Height          =   195
         Left            =   8640
         TabIndex        =   53
         Top             =   345
         Width           =   1725
      End
      Begin VB.OptionButton OptNinguno 
         Caption         =   "Ninguno"
         Height          =   195
         Left            =   6990
         TabIndex        =   52
         Top             =   315
         Width           =   1725
      End
      Begin VB.Frame fraModoTexto 
         Caption         =   "Modo envío"
         Height          =   1095
         Left            =   16800
         TabIndex        =   47
         Top             =   4560
         Width           =   3375
         Begin VB.TextBox txtIntervaloTecla 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   58
            Text            =   "100"
            Top             =   285
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.OptionButton optPegarTexto 
            Caption         =   "Pegar texto"
            Height          =   375
            Left            =   240
            TabIndex        =   49
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton optPulsarTecla 
            Caption         =   "Pulsar tecla"
            Height          =   375
            Left            =   240
            TabIndex        =   48
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblIntervaloTecla 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ms."
            Height          =   195
            Left            =   2280
            TabIndex        =   59
            Top             =   360
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VB.CheckBox chkCerrarApp 
         Caption         =   "Cerrar la aplicación al finalizar secuencia"
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   405
         TabIndex        =   46
         Top             =   5280
         Width           =   4575
      End
      Begin VB.OptionButton OptEnter 
         Caption         =   "Enviar ENTER"
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   405
         TabIndex        =   43
         Top             =   5040
         Width           =   1500
      End
      Begin VB.CheckBox ChkXY 
         Caption         =   "En coordenadas específicas dentro de la ventana:"
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   810
         TabIndex        =   37
         Top             =   690
         Width           =   3930
      End
      Begin VB.OptionButton OptTexto 
         Caption         =   "Texto:"
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   405
         TabIndex        =   23
         Top             =   1560
         Width           =   885
      End
      Begin VB.TextBox TxtCy 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6075
         MaxLength       =   4
         TabIndex        =   22
         Top             =   720
         Width           =   780
      End
      Begin VB.TextBox txtCx 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4950
         MaxLength       =   4
         TabIndex        =   21
         Top             =   720
         Width           =   780
      End
      Begin VB.OptionButton OptTab 
         Caption         =   "Enviar TAB"
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   405
         TabIndex        =   18
         Top             =   4770
         Width           =   1320
      End
      Begin VB.TextBox txtTexto 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3000
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   17
         Top             =   1522
         Width           =   18375
      End
      Begin VB.OptionButton OptDobleClick 
         Caption         =   "Doble click"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5145
         TabIndex        =   16
         Top             =   345
         Width           =   1725
      End
      Begin VB.OptionButton optCerrarVentana 
         Caption         =   "Enviar cierre de ventana"
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   405
         TabIndex        =   15
         Top             =   1080
         Width           =   2280
      End
      Begin VB.OptionButton OptClickDerecho 
         Caption         =   "Click Derecho"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3300
         TabIndex        =   2
         Top             =   345
         Width           =   1725
      End
      Begin VB.OptionButton optClickIzquierdo 
         Caption         =   "Click"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   840
         TabIndex        =   1
         Top             =   345
         Value           =   -1  'True
         Width           =   1725
      End
      Begin VB.Label lblSg 
         AutoSize        =   -1  'True
         Caption         =   "sg."
         Height          =   195
         Left            =   10500
         TabIndex        =   55
         Top             =   435
         Width           =   210
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5895
         TabIndex        =   20
         Top             =   765
         Width           =   150
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4770
         TabIndex        =   19
         Top             =   765
         Width           =   150
      End
   End
   Begin BotonXp.BotónXp CmdInsertar 
      Height          =   375
      Left            =   2470
      TabIndex        =   40
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
      MICON           =   "Asistente.frx":3531
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
      Left            =   4175
      TabIndex        =   41
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
      MICON           =   "Asistente.frx":354D
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
      Left            =   5880
      TabIndex        =   42
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
      MICON           =   "Asistente.frx":3569
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid DgEventos 
      Height          =   4335
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   28695
      _ExtentX        =   50615
      _ExtentY        =   7646
      _Version        =   393216
      BackColorBkg    =   16777215
      Appearance      =   0
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   0
      TabIndex        =   60
      Top             =   11280
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   15960
      Picture         =   "Asistente.frx":3585
      Stretch         =   -1  'True
      Top             =   13080
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

' ====================================================
'  CONSTANTES Y TIPOS
' ====================================================
Private Const MAX_RECIENTES   As Long = 4
Private Const TIMEOUT_VENTANA As Long = 5000
Private Const MAX_EVENTOS     As Long = 500
Private Const SEC_META        As String = "Meta"
Private Const SEC_EVENTO      As String = "Evento"

' Columnas de la grilla
Private Const COL_ID             As Long = 0
Private Const COL_HABILITADO     As Long = 1
Private Const COL_RUTA           As Long = 2
Private Const COL_PARAMETROS     As Long = 3
Private Const COL_TEXTO_VENTANA  As Long = 4
Private Const COL_TEXTO_SUBCLASE As Long = 5
Private Const COL_EVENTO         As Long = 6
Private Const COL_MODO_ENVIO     As Long = 7   ' nuevo: 1=Tecla, 2=Pegar, 0=sin definir
Private Const COL_TEXTO          As Long = 8
Private Const COL_TIEMPO_ESPERA  As Long = 9   ' nuevo: segundos de espera (solo si Evento=evTiempoEspera)
Private Const COL_CLICK_X        As Long = 10
Private Const COL_CLICK_Y        As Long = 11
Private Const COL_TOTAL          As Long = 12  ' total de columnas

' Valores de ModoEnvioTexto guardados en TEvento
Private Const MODO_PULSAR_TECLA As Long = 1
Private Const MODO_PEGAR_TEXTO  As Long = 2

Private Const CB_SHOWDROPDOWN As Long = &H14F

Private Enum eEvento
    evClickIzquierdo = 1
    evClickDerecho = 2
    evDobleClick = 3
    evCerrarVentana = 4
    evTexto = 5
    evTab = 6
    evEnter = 7
    evTiempoEspera = 8     ' nuevo: espera bloqueante
    evNinguno = 9          ' nuevo: no aplica / ninguno
End Enum

Private Type TEvento
    Id             As Long
    habilitado     As Boolean
    Ruta           As String
    parametros     As String
    TextoVentana   As String
    TextoSubClase  As String
    Evento         As Long
    Texto          As String
    ClickPosX      As Long
    ClickPosY      As Long
    ModoEnvioTexto As Long   ' 1=PulsarTecla  2=PegarTexto  0=sin definir
    IntervaloTecla As Long   ' milisegundos entre cada tecla (solo si ModoEnvioTexto=1)
    tiempoEspera   As Long   ' segundos de espera (solo si Evento=evTiempoEspera)
End Type

' ====================================================
'  ESTADO DEL MÓDULO
' ====================================================
Private m_Eventos()    As TEvento
Private m_Count        As Long
Private m_FilaActual   As Long
Private m_CargandoFila As Boolean   ' bandera para suprimir _Click durante carga

Private m_Archivo  As String
Private m_Bloquear As Boolean
Private WithEvents m_EnumWin As clsEnum
Attribute m_EnumWin.VB_VarHelpID = -1
Private m_LastPID  As Long
Private UltFila    As Long
Private UltCol     As Long

Const FILE_INI = "config.ini"
Public IsByParam As Boolean
Public g_GrabarLog As Boolean

Private Const LOG_FILE As String = "WindowsHandler.log"

Private Sub EscribirLog(ByVal accion As String, Optional ByVal parametros As String = "")
    If Not g_GrabarLog Then Exit Sub  ' si está desactivado, no hace nada
    
    Dim f As Integer
    f = FreeFile
    On Error Resume Next
    Open App.Path & "\" & LOG_FILE For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & accion & _
              IIf(parametros <> "", " | " & parametros, "")
    Close #f
End Sub


' ====================================================
'  HELPERS DE ARRAY
' ====================================================
Private Function HayEventos() As Boolean
    EscribirLog "HayEventos", "m_Count=" & m_Count
    HayEventos = (m_Count > 0)
End Function

Private Function AppendEvento() As Long
    EscribirLog "AppendEvento"
    m_Count = m_Count + 1
    ReDim Preserve m_Eventos(1 To m_Count)
    With m_Eventos(m_Count)
        .Id = m_Count
        .habilitado = True
        .Evento = evClickIzquierdo
        .ModoEnvioTexto = 0
    End With
    AppendEvento = m_Count
End Function

Private Sub RemoveEvento(ByVal pos As Long)
    EscribirLog "RemoveEvento", "pos=" & pos
    If pos < 1 Or pos > m_Count Then Exit Sub
    Dim i As Long
    For i = pos To m_Count - 1
        m_Eventos(i) = m_Eventos(i + 1)
        m_Eventos(i).Id = i
    Next i
    m_Count = m_Count - 1
    If m_Count > 0 Then
        ReDim Preserve m_Eventos(1 To m_Count)
    Else
        ReDim m_Eventos(1 To 1)
    End If
End Sub

Private Sub InsertEvento(ByVal pos As Long)
    EscribirLog "InsertEvento", "pos=" & pos
    m_Count = m_Count + 1
    ReDim Preserve m_Eventos(1 To m_Count)
    Dim i As Long
    For i = m_Count To pos + 1 Step -1
        m_Eventos(i) = m_Eventos(i - 1)
    Next i
    Dim nuevo As TEvento
    nuevo.Id = pos
    nuevo.habilitado = True
    nuevo.Evento = evClickIzquierdo
    nuevo.ModoEnvioTexto = 0
    m_Eventos(pos) = nuevo
    For i = pos To m_Count
        m_Eventos(i).Id = i
    Next i
End Sub

Private Sub InicializarArray()
    EscribirLog "InicializarArray"
    m_Count = 0
    m_FilaActual = 0
    ReDim m_Eventos(1 To 1)
End Sub

Private Function FilaValida() As Boolean
    EscribirLog "FilaValida", "m_FilaActual=" & m_FilaActual
    FilaValida = (m_FilaActual >= 1 And m_FilaActual <= m_Count)
End Function

Private Function EsEventoTexto(ByVal tipoEvento As Long) As Boolean
    EscribirLog "EsEventoTexto", "tipoEvento=" & tipoEvento
    EsEventoTexto = (tipoEvento = evTexto)
End Function

' ====================================================
'  HELPERS DE EVENTO (descripciones para la grilla)
' ====================================================
Private Function DescripcionEvento(ByVal ev As Long) As String
    Select Case ev
        Case evClickIzquierdo: DescripcionEvento = "Click"
        Case evClickDerecho:   DescripcionEvento = "Click derecho"
        Case evDobleClick:     DescripcionEvento = "Doble click"
        Case evCerrarVentana:  DescripcionEvento = "Cerrar"
        Case evTexto:          DescripcionEvento = "Texto"
        Case evTab:            DescripcionEvento = "Tab"
        Case evEnter:          DescripcionEvento = "Enter"
        Case evTiempoEspera:   DescripcionEvento = "Esperar"
        Case evNinguno:        DescripcionEvento = "Ninguno"
        Case Else:             DescripcionEvento = CStr(ev)
    End Select
End Function

Private Function EventoDesdeIndiceCombo(ByVal indice As Long) As Long
    If indice < 0 Or indice >= cboEditor.ListCount Then
        EventoDesdeIndiceCombo = evClickIzquierdo
    Else
        EventoDesdeIndiceCombo = cboEditor.ItemData(indice)
    End If
End Function

Private Sub SeleccionarEventoEnCombo(ByVal ev As Long)
    Dim i As Long
    m_CargandoFila = True
    For i = 0 To cboEditor.ListCount - 1
        If cboEditor.ItemData(i) = ev Then
            cboEditor.ListIndex = i
            Exit For
        End If
    Next i
    m_CargandoFila = False
End Sub

Private Function TextoHabilitado(ByVal habilitado As Boolean) As String
    If habilitado Then TextoHabilitado = Chr$(252) Else TextoHabilitado = ""
End Function

Private Sub InicializarComboEventos()
    With cboEditor
        .Clear
        .AddItem "Ninguno":        .ItemData(.NewIndex) = evNinguno
        .AddItem "Click":          .ItemData(.NewIndex) = evClickIzquierdo
        .AddItem "Click derecho":  .ItemData(.NewIndex) = evClickDerecho
        .AddItem "Doble click":    .ItemData(.NewIndex) = evDobleClick
        .AddItem "Cerrar":         .ItemData(.NewIndex) = evCerrarVentana
        .AddItem "Texto":          .ItemData(.NewIndex) = evTexto
        .AddItem "Tab":            .ItemData(.NewIndex) = evTab
        .AddItem "Enter":          .ItemData(.NewIndex) = evEnter
        .AddItem "Esperar": .ItemData(.NewIndex) = evTiempoEspera
    End With
End Sub

' ====================================================
'  EDITORES SUPERPUESTOS (txtEditor / chkEditor / cboEditor)
' ====================================================
Private Sub OcultarEditores()
    txtEditor.Visible = False
    chkEditor.Visible = False
    cboEditor.Visible = False
    cboModoEnvio.Visible = False
    udTiempoEspera.Visible = False
End Sub

Private Sub PosicionarControlSobreCelda( _
    ByVal ctrl As Control, _
    ByVal fila As Long, _
    ByVal col As Long, _
    Optional ByVal anchoPx As Single = 0, _
    Optional ByVal altoPx As Single = 0)

    DgEventos.Row = fila
    DgEventos.col = col

    Dim X As Single, Y As Single, w As Single, h As Single
    X = ScaleX(DgEventos.CellLeft, vbTwips, vbPixels) + DgEventos.Left
    Y = ScaleY(DgEventos.CellTop, vbTwips, vbPixels) + DgEventos.Top
    w = ScaleX(DgEventos.CellWidth, vbTwips, vbPixels)
    h = ScaleY(DgEventos.CellHeight, vbTwips, vbPixels)

    If anchoPx > 0 Then w = anchoPx
    If altoPx > 0 Then h = altoPx

    ctrl.Move X, Y, w, h
    
    ctrl.Visible = True
    ctrl.ZOrder 0
End Sub

Private Sub RefrescarCeldaTiempoEspera(ByVal fila As Long)
    If fila < 1 Or fila > m_Count Then Exit Sub
    With DgEventos
        .Row = fila
        .col = COL_TIEMPO_ESPERA

        If m_Eventos(fila).Evento = evTiempoEspera Then
            .TextMatrix(fila, COL_TIEMPO_ESPERA) = CStr(m_Eventos(fila).tiempoEspera)
        Else
            .TextMatrix(fila, COL_TIEMPO_ESPERA) = "0"
        End If
    End With
End Sub

Private Sub MostrarEditorTiempoEspera(ByVal fila As Long)
    If fila < 1 Or fila > m_Count Then Exit Sub

    ' Solo mostrar si el evento es Esperar
    If m_Eventos(fila).Evento <> evTiempoEspera Then
        OcultarEditores
        Exit Sub
    End If

    OcultarEditores

    Dim X As Single, Y As Single, w As Single, h As Single
    ObtenerRectCelda fila, COL_TIEMPO_ESPERA, X, Y, w, h

    ' TextBox ocupa casi toda la celda
    txtEditor.Move X, Y, w, h
    txtEditor.Text = CStr(m_Eventos(fila).tiempoEspera)
    txtEditor.Visible = True

    ' UpDown angosto a la derecha
    udTiempoEspera.Move X + w - (udTiempoEspera.Width), Y, 20, h
    udTiempoEspera.Visible = True
    udTiempoEspera.ZOrder 0
    txtEditor.SetFocus
End Sub

Private Sub MostrarEditorCheckbox(ByVal fila As Long)
    If fila < 1 Or fila > m_Count Then Exit Sub
    OcultarEditores
    m_CargandoFila = True
    chkEditor.Value = IIf(m_Eventos(fila).habilitado, vbChecked, vbUnchecked)
    m_CargandoFila = False
    PosicionarControlSobreCelda chkEditor, fila, COL_HABILITADO, 18, 18
    chkEditor.SetFocus
End Sub

Private Sub MostrarEditorComboEvento(ByVal fila As Long)
    If fila < 1 Or fila > m_Count Then Exit Sub
    
    OcultarEditores
    SeleccionarEventoEnCombo m_Eventos(fila).Evento

    Dim X As Single, Y As Single, w As Single, h As Single
    ObtenerRectCelda fila, COL_EVENTO, X, Y, w, h

    ' ComboBox Style=2: no se puede fijar Height ? Move sin h
    cboEditor.Move X, Y, w
    cboEditor.Visible = True
    cboEditor.ZOrder 0
    cboEditor.SetFocus
  

    'SendMessage cboEditor.hwnd, CB_SHOWDROPDOWN, 1, 0
End Sub

Private Sub MostrarEditorComboModoEnvio(ByVal fila As Long)
    If fila < 1 Or fila > m_Count Then Exit Sub
    If m_Eventos(fila).Evento <> evTexto Then Exit Sub  ' solo aplica a Texto

    OcultarEditores
    cboModoEnvio.Clear
    cboModoEnvio.AddItem "Pulsar tecla": cboModoEnvio.ItemData(cboModoEnvio.NewIndex) = MODO_PULSAR_TECLA
    cboModoEnvio.AddItem "Pegar texto": cboModoEnvio.ItemData(cboModoEnvio.NewIndex) = MODO_PEGAR_TEXTO

    ' Seleccionar el valor actual
    Dim i As Long
    For i = 0 To cboModoEnvio.ListCount - 1
        If cboModoEnvio.ItemData(i) = m_Eventos(fila).ModoEnvioTexto Then
            cboModoEnvio.ListIndex = i
            Exit For
        End If
    Next i

    ' Posicionar combo sobre la celda
    Dim X As Single, Y As Single, w As Single, h As Single
    ObtenerRectCelda fila, COL_MODO_ENVIO, X, Y, w, h
    cboModoEnvio.Move X, Y, w
    cboModoEnvio.Visible = True
    cboModoEnvio.ZOrder 0
    cboModoEnvio.SetFocus
End Sub

Private Sub RefrescarCeldaHabilitado(ByVal fila As Long)
    If fila < 1 Or fila > m_Count Then Exit Sub
    DgEventos.Row = fila
    DgEventos.col = COL_HABILITADO
    DgEventos.CellFontName = "Wingdings"

    If m_Eventos(fila).habilitado Then

        DgEventos.TextMatrix(fila, COL_HABILITADO) = "ü"   ' tilde
        'DgEventos.CellForeColor = vbGreen
    Else
        DgEventos.TextMatrix(fila, COL_HABILITADO) = "û"   ' cruz
        'DgEventos.CellForeColor = vbRed
    End If
End Sub

Private Sub AlternarHabilitado(ByVal fila As Long)
    If fila < 1 Or fila > m_Count Then Exit Sub
    m_Eventos(fila).habilitado = Not m_Eventos(fila).habilitado
    DgEventos.Row = fila
    DgEventos.col = COL_HABILITADO
    RefrescarCeldaHabilitado fila
End Sub

Private Sub RefrescarCeldaEvento(ByVal fila As Long)
    If fila < 1 Or fila > m_Count Then Exit Sub
    DgEventos.TextMatrix(fila, COL_EVENTO) = DescripcionEvento(m_Eventos(fila).Evento)
End Sub

Private Sub RefrescarCeldaModoEnvio(ByVal fila As Long)
    If fila < 1 Or fila > m_Count Then Exit Sub
    Dim Texto As String
    Select Case m_Eventos(fila).ModoEnvioTexto
        Case MODO_PULSAR_TECLA: Texto = "Pulsar tecla"
        Case MODO_PEGAR_TEXTO:  Texto = "Pegar texto"
        Case Else:              Texto = ""
    End Select
    DgEventos.TextMatrix(fila, COL_MODO_ENVIO) = Texto
End Sub

Private Sub ObtenerRectCelda( _
    ByVal fila As Long, _
    ByVal col As Long, _
    ByRef X As Single, _
    ByRef Y As Single, _
    ByRef w As Single, _
    ByRef h As Single, _
    Optional ByVal anchoExtra As Single = 0, _
    Optional ByVal altoExtra As Single = 0)
    DgEventos.Row = fila
    DgEventos.col = col
    If Me.ScaleMode = vbPixels Then
        X = DgEventos.Left + ScaleX(DgEventos.CellLeft, vbTwips, vbPixels)
        Y = DgEventos.Top + ScaleY(DgEventos.CellTop, vbTwips, vbPixels)
        w = ScaleX(DgEventos.CellWidth, vbTwips, vbPixels)
        h = ScaleY(DgEventos.CellHeight, vbTwips, vbPixels)
    Else
        X = DgEventos.Left + DgEventos.CellLeft
        Y = DgEventos.Top + DgEventos.CellTop
        w = DgEventos.CellWidth
        h = DgEventos.CellHeight
    End If
    If anchoExtra > 0 Then w = anchoExtra
    If altoExtra > 0 Then h = altoExtra
End Sub

Private Sub CargarDesdeArchivo(ByVal Archivo As String)
    EscribirLog "CargarDesdeArchivo", "Archivo=" & Archivo
    Dim n As Long, i As Long, sec As String
    n = CLng(Val(LeeINI(Archivo, SEC_META, "Count", "0")))
    If n < 0 Then n = 0
    If n > MAX_EVENTOS Then n = MAX_EVENTOS

    m_Count = n
    If n = 0 Then ReDim m_Eventos(1 To 1): Exit Sub

    ReDim m_Eventos(1 To n)
    For i = 1 To n
        sec = SEC_EVENTO & CStr(i)
        With m_Eventos(i)
            .Id = CLng(Val(LeeINI(Archivo, sec, "Id", CStr(i))))
            .habilitado = (CLng(Val(LeeINI(Archivo, sec, "Habilitado", "1"))) <> 0)
            .Ruta = LeeINI(Archivo, sec, "Ruta", "")
            .parametros = LeeINI(Archivo, sec, "Parametros", "")
            .TextoVentana = LeeINI(Archivo, sec, "TextoVentana", "")
            .TextoSubClase = LeeINI(Archivo, sec, "TextoSubClase", "")
            .Evento = CLng(Val(LeeINI(Archivo, sec, "Evento", "1")))
            .Texto = LeeINI(Archivo, sec, "Texto", "")
            .ClickPosX = CLng(Val(LeeINI(Archivo, sec, "ClickPosX", "0")))
            .ClickPosY = CLng(Val(LeeINI(Archivo, sec, "ClickPosY", "0")))
            .ModoEnvioTexto = CLng(Val(LeeINI(Archivo, sec, "ModoEnvioTexto", "0")))
            .IntervaloTecla = CLng(Val(LeeINI(Archivo, sec, "IntervaloTecla", "100")))
            .tiempoEspera = CLng(Val(LeeINI(Archivo, sec, "TiempoEspera", "0")))
        End With
    Next i
End Sub

Private Sub GuardarEnArchivo(ByVal Archivo As String)
    EscribirLog "GuardarEnArchivo", "Archivo=" & Archivo & ", m_Count=" & m_Count
    If Dir$(Archivo) <> "" Then Kill Archivo
    GrabaINI Archivo, SEC_META, "Count", CStr(m_Count)
    ' Guardar la opción en el archivo del proyecto actual (.sca)
    GrabaINI Archivo, "Proyecto", "CerrarAlFinalizar", chkCerrarApp.Value
    Dim i As Long, sec As String
    For i = 1 To m_Count
        sec = SEC_EVENTO & CStr(i)
        With m_Eventos(i)
            GrabaINI Archivo, sec, "Id", CStr(.Id)
            GrabaINI Archivo, sec, "Habilitado", IIf(.habilitado, "1", "0")
            GrabaINI Archivo, sec, "Ruta", .Ruta
            GrabaINI Archivo, sec, "Parametros", .parametros
            GrabaINI Archivo, sec, "TextoVentana", .TextoVentana
            GrabaINI Archivo, sec, "TextoSubClase", .TextoSubClase
            GrabaINI Archivo, sec, "Evento", CStr(.Evento)
            GrabaINI Archivo, sec, "Texto", .Texto
            GrabaINI Archivo, sec, "ClickPosX", CStr(.ClickPosX)
            GrabaINI Archivo, sec, "ClickPosY", CStr(.ClickPosY)
            GrabaINI Archivo, sec, "ModoEnvioTexto", CStr(.ModoEnvioTexto)
            GrabaINI Archivo, sec, "IntervaloTecla", CStr(.IntervaloTecla)
            GrabaINI Archivo, sec, "TiempoEspera", CStr(.tiempoEspera)
        End With
    Next i
End Sub

' ====================================================
'  FORM LOAD / UNLOAD
' ====================================================
Public Sub LoadForm()
    EscribirLog "LoadForm"
    IsByParam = True
    Call Form_Load
End Sub

Private Sub cboModoEnvio_Click()
    Dim fila As Long
    fila = DgEventos.Row
    If fila < 1 Or fila > m_Count Then Exit Sub
    m_Eventos(fila).ModoEnvioTexto = cboModoEnvio.ItemData(cboModoEnvio.ListIndex)
    DgEventos.TextMatrix(fila, COL_MODO_ENVIO) = cboModoEnvio.Text
End Sub

Private Sub Form_Load()
    CargarConfiguracionGeneral
    EscribirLog "Form_Load"
    Set m_EnumWin = New clsEnum
    InicializarArray
    InicializarComboEventos
    ConfigurarFlexGrid
    CargarRecientes
    AjustarColumnas
    fraModoTexto.Visible = False
    Me.Caption = Me.Tag
    mnuGP.Enabled = True
End Sub

Private Sub Form_Resize()
   With lblVersion
        .Caption = "Versión: " & App.Major & "." & App.Minor & "." & App.Revision
        .Visible = True
        .AutoSize = True
        .ForeColor = vbBlack
        .Left = 0
        .Top = Me.ScaleHeight - .Height
    End With
End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    RestaurarLayoutVentana
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EscribirLog "Form_Unload", "Cancel=" & Cancel
    On Error Resume Next
    RestaurarLayoutVentana
    VBHotKey1.StopHotkey
    Set DgEventos.DataSource = Nothing
    Set m_EnumWin = Nothing
End Sub

' ====================================================
'  CONFIGURACIÓN GENERAL DEL PROYECTO ACTUAL
' ====================================================
Private Sub CargarConfiguracionProyecto()
   EscribirLog "CargarConfiguracionProyecto"
    If m_Archivo <> "" Then
        chkCerrarApp.Value = CInt(LeeINI(m_Archivo, "Proyecto", "CerrarAlFinalizar", "0"))
    Else
        chkCerrarApp.Value = vbUnchecked
    End If
End Sub

' ====================================================
'  RECIENTES
' ====================================================
Private Sub CargarRecientes()
    EscribirLog "CargarRecientes"
    Dim i As Long, Archivo As String
    For i = 1 To MAX_RECIENTES
        Archivo = LeeINI(App.Path & "\" & FILE_INI, "MENU", "RECIENTE" & CStr(i))
        If Trim$(Archivo) <> "" And Dir$(Archivo) <> "" Then
            mnuTmp(i).Caption = CStr(i) & " " & Archivo
            mnuTmp(i).Visible = True
        End If
    Next i
End Sub

' ====================================================
'  CARGAR CONFIGURACION GENERAL DEL PROGRAMA
' ====================================================
Private Sub CargarConfiguracionGeneral()
    EscribirLog "CargarConfiguracionGeneral"
    Dim valor As String
    valor = LeeINI(App.Path & "\" & FILE_INI, "General", "GrabarLog", 0)
    g_GrabarLog = CBool(valor)
End Sub

Private Sub AgregarMenuReciente(ByVal Archivo As String)
    EscribirLog "AgregarMenuReciente", "Archivo=" & Archivo
    Dim i As Long, Existe As Boolean
    For i = 1 To MAX_RECIENTES
        If LCase$(Trim$(mnuTmp(i).Caption)) = LCase$(CStr(i) & " " & Archivo) Then
            Existe = True: Exit For
        End If
    Next i
    If Existe Then Exit Sub
    For i = MAX_RECIENTES To 2 Step -1
        mnuTmp(i).Caption = mnuTmp(i - 1).Caption
        mnuTmp(i).Visible = mnuTmp(i - 1).Visible
    Next i
    mnuTmp(1).Caption = "1 " & Archivo
    mnuTmp(1).Visible = True
    GuardarRecientesINI
End Sub

Private Sub GuardarRecientesINI()
    EscribirLog "GuardarRecientesINI"
    Dim i As Long, Archivo As String
    For i = 1 To MAX_RECIENTES
        Archivo = ""
        If mnuTmp(i).Visible Then
            Archivo = Mid$(mnuTmp(i).Caption, InStr(1, mnuTmp(i).Caption, " ") + 1)
        End If
        GrabaINI App.Path & "\" & FILE_INI, "MENU", "RECIENTE" & CStr(i), Archivo
    Next i
End Sub

Private Sub mnuTmp_Click(Index As Integer)
    EscribirLog "mnuTmp_Click", "Index=" & Index & ", Caption=" & mnuTmp(Index).Caption
    Dim Ruta As String
    Ruta = Trim$(Mid$(mnuTmp(Index).Caption, InStr(1, mnuTmp(Index).Caption, " ") + 1))
    If Ruta = "" Then Exit Sub
    If Dir$(Ruta) = "" Then MsgBox "Archivo inexistente", vbExclamation: Exit Sub
    AbrirProyecto Ruta
End Sub


' ====================================================
'  ABRIR PROYECTO
' ====================================================
Public Sub AbrirProyecto(ByVal RutaArchivo As String)
    EscribirLog "AbrirProyecto", "RutaArchivo=" & RutaArchivo
    On Error GoTo ErrHandler
    If Dir$(RutaArchivo) = "" Then MsgBox "No existe el archivo", vbCritical: Exit Sub
    m_Archivo = RutaArchivo
    CargarDesdeArchivo m_Archivo
    CargarEventosDesdeArray
    CargarConfiguracionProyecto
    Me.Caption = Me.Tag & " - [" & m_Archivo & "]"
    Exit Sub
ErrHandler:
    EscribirLog "AbrirProyecto_Error", "Error=" & Err.Description
    MsgBox Err.Description, vbCritical
End Sub

' ====================================================
'  CONFIGURAR FLEXGRID
' ====================================================
Private Sub ConfigurarFlexGrid()
    EscribirLog "ConfigurarFlexGrid"
    With DgEventos
        .Cols = COL_TOTAL
        .TextMatrix(0, COL_ID) = "Id"
        .TextMatrix(0, COL_HABILITADO) = "Habilitado"
        .TextMatrix(0, COL_RUTA) = "Ruta"
        .TextMatrix(0, COL_PARAMETROS) = "Parametros"
        .TextMatrix(0, COL_TEXTO_VENTANA) = "TextoVentana"
        .TextMatrix(0, COL_TEXTO_SUBCLASE) = "TextoSubClase"
        .TextMatrix(0, COL_EVENTO) = "Evento"
        .TextMatrix(0, COL_MODO_ENVIO) = "ModoEnvio"
        .TextMatrix(0, COL_TEXTO) = "Texto"
        .TextMatrix(0, COL_TIEMPO_ESPERA) = "TiempoEspera"
        .TextMatrix(0, COL_CLICK_X) = "ClickPosX"
        .TextMatrix(0, COL_CLICK_Y) = "ClickPosY"
        
        Set .Font = New StdFont
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        txtEditor.Font.Name = .Font.Name
        txtEditor.Font.Size = .Font.Size
        cboEditor.Font.Name = .Font.Name
        cboEditor.Font.Size = .Font.Size - 2
        cboModoEnvio.Font.Name = cboEditor.Font.Name
        cboModoEnvio.Font.Size = cboEditor.Font.Size
        .BackColor = RGB(245, 245, 245)
        .ForeColor = RGB(30, 30, 30)
        .GridColor = RGB(200, 200, 200)
        .BackColorFixed = RGB(220, 220, 220)
        .ForeColorFixed = vbBlack
        .SelectionMode = flexSelectionFree
        .BackColorSel = RGB(173, 216, 230)
        .ForeColorSel = vbBlack
        .CellBackColor = RGB(240, 240, 240)
    End With
    
    ' Marca visual de checkbox en celdas Habilitado (Wingdings)
    DgEventos.col = COL_HABILITADO
    DgEventos.CellFontName = "Wingdings"
    
    DgEventos.ColWidth(COL_MODO_ENVIO) = 0
End Sub

Private Sub AjustarColumnas()
    EscribirLog "AjustarColumnas"
    Dim c As Long, r As Long, maxAnchoPx As Long, Texto As String
    Const ANCHO_MIN As Long = 80
    Const ANCHO_MAX As Long = 600
    Const margen    As Long = 12
    For c = 0 To DgEventos.Cols - 1
        If c = COL_HABILITADO Then
            GoTo SiguienteCol
        End If
        maxAnchoPx = Me.TextWidth(" " & DgEventos.TextMatrix(0, c) & " ")
        For r = 1 To DgEventos.Rows - 1
            Texto = DgEventos.TextMatrix(r, c)
            If Len(Texto) > 0 Then
                maxAnchoPx = Max(maxAnchoPx, Me.TextWidth(" " & Texto & " "))
            End If
        Next r
        maxAnchoPx = maxAnchoPx + margen
        If maxAnchoPx < ANCHO_MIN Then maxAnchoPx = ANCHO_MIN
        If maxAnchoPx > ANCHO_MAX Then maxAnchoPx = ANCHO_MAX
        DgEventos.ColWidth(c) = ScaleX(maxAnchoPx, vbPixels, vbTwips) * 1.5
SiguienteCol:
    Next c
    DgEventos.Refresh
End Sub

Private Function Max(a As Long, b As Long) As Long
    If a > b Then Max = a Else Max = b
End Function


' ====================================================
'  CARGAR GRILLA DESDE ARRAY
' ====================================================
Private Sub CargarEventosDesdeArray()
    EscribirLog "CargarEventosDesdeArray", "m_Count=" & m_Count
    OcultarEditores
    DgEventos.Rows = 1
    If m_Count = 0 Then Exit Sub
    
    Dim i As Long, valorModo As String
    For i = 1 To m_Count
        With m_Eventos(i)

            DgEventos.AddItem _
                CStr(.Id) & vbTab & _
                TextoHabilitado(.habilitado) & vbTab & _
                .Ruta & vbTab & _
                .parametros & vbTab & _
                .TextoVentana & vbTab & _
                .TextoSubClase & vbTab & _
                DescripcionEvento(.Evento) & vbTab & _
                IIf(.Evento = evTexto, _
                    IIf(.ModoEnvioTexto = MODO_PULSAR_TECLA, "Pulsar tecla", "Pegar texto"), _
                    "") & vbTab & _
                .Texto & vbTab & _
                CStr(.tiempoEspera) & vbTab & _
                CStr(.ClickPosX) & vbTab & _
                CStr(.ClickPosY)
        
'                'ChrW(&H2713) <- es cuadernito con lapiz
'                If m_Eventos(i).habilitado Then
'                    DgEventos.TextMatrix(i, COL_HABILITADO) = "ü"   ' tilde
'                Else
'                    DgEventos.TextMatrix(i, COL_HABILITADO) = "û"   ' cruz
'
'                End If
                RefrescarCeldaHabilitado (i)
                RefrescarCeldaTiempoEspera (i)

        End With
        DgEventos.Row = i
        DgEventos.col = COL_HABILITADO
        DgEventos.CellFontName = "Wingdings"
    Next i
    AjustarColumnas
End Sub

' ====================================================
'  NUEVO / INSERTAR / ELIMINAR
' ====================================================
Private Sub cmdNuevo_Click()
    EscribirLog "cmdNuevo_Click"
    AppendEvento
    CargarEventosDesdeArray
    DgEventos.Row = m_Count
End Sub

Private Sub CmdInsertar_Click()
    EscribirLog "CmdInsertar_Click", "fila=" & DgEventos.Row
    Dim fila As Long
    fila = DgEventos.Row
    If fila = 0 Then Exit Sub
    InsertEvento fila
    CargarEventosDesdeArray
    DgEventos.Row = fila
End Sub

Private Sub CmdEliminar_Click()
    EscribirLog "CmdEliminar_Click", "fila=" & DgEventos.Row
    Dim fila As Long
    fila = DgEventos.Row
    If fila = 0 Then Exit Sub

    UltFila = 0
    UltCol = 0
    OcultarEditores

    RemoveEvento fila
    CargarEventosDesdeArray

    If m_Count > 0 Then
        DgEventos.Row = IIf(fila <= m_Count, fila, m_Count)
        m_FilaActual = DgEventos.Row
        MostrarEvento m_FilaActual
    Else
        m_FilaActual = 0
        fraModoTexto.Visible = False
    End If
End Sub


' ====================================================
'  MOSTRAR EVENTO EN CONTROLES
' ====================================================
Private Sub MostrarEvento(ByVal fila As Long)
    EscribirLog "MostrarEvento", "fila=" & fila
    If fila < 1 Or fila > m_Count Then
        fraModoTexto.Visible = False
        Exit Sub
    End If

    m_CargandoFila = True

    Dim ev As TEvento
    ev = m_Eventos(fila)

    txtTexto.Text = ev.Texto
    txtCx.Text = CStr(ev.ClickPosX)
    TxtCy.Text = CStr(ev.ClickPosY)

    optClickIzquierdo.Value = (ev.Evento = evClickIzquierdo)
    OptClickDerecho.Value = (ev.Evento = evClickDerecho)
    OptDobleClick.Value = (ev.Evento = evDobleClick)
    optCerrarVentana.Value = (ev.Evento = evCerrarVentana)
    OptTexto.Value = (ev.Evento = evTexto)
    OptTab.Value = (ev.Evento = evTab)
    OptEnter.Value = (ev.Evento = evEnter)
    OptTiempoEspera.Value = (ev.Evento = evTiempoEspera)
    OptNinguno.Value = (ev.Evento = evNinguno)

    If ev.Evento = evTiempoEspera Then
        txtTiempoEspera.Text = CStr(ev.tiempoEspera)
        txtTiempoEspera.Enabled = True
        fraModoTexto.Visible = False
    Else
        txtTiempoEspera.Text = "0"        ' limpiar valor
        txtTiempoEspera.Enabled = False   ' grisar/deshabilitar
        fraModoTexto.Visible = (ev.Evento = evTexto)
    End If
    If ev.Evento = evTexto Then
        If ev.Evento = evTexto Then
            If ev.ModoEnvioTexto <> MODO_PEGAR_TEXTO Then
                If ev.ModoEnvioTexto = 0 Then
                    m_Eventos(fila).ModoEnvioTexto = MODO_PULSAR_TECLA
                End If
                optPulsarTecla.Value = True
                optPulsarTecla.Caption = "Pulsar tecla cada:"
                txtIntervaloTecla.Text = ev.IntervaloTecla
                txtIntervaloTecla.Visible = True
                lblIntervaloTecla.Visible = True
                optPegarTexto.Value = False
            Else
                optPulsarTecla.Value = False
                optPulsarTecla.Caption = "Pulsar tecla"
                txtIntervaloTecla.Visible = False
                lblIntervaloTecla.Visible = False
                optPegarTexto.Value = True
                
            End If
            fraModoTexto.Visible = True
        Else
            fraModoTexto.Visible = False
        End If
    Else
    fraModoTexto.Visible = False
    End If

    
    m_CargandoFila = False
End Sub


' ====================================================
'  CLICK EN LA GRILLA
' ====================================================
Private Sub DgEventos_Click()
    EscribirLog "DgEventos_Click", "Row=" & DgEventos.Row
    Dim fila As Long
    fila = DgEventos.Row
    If fila = 0 Then Exit Sub
    m_FilaActual = fila
    MostrarEvento fila
    DgEventos_PaintCells
End Sub


' ====================================================
'  ACTUALIZAR ARRAY <- CELDA EDITADA
' ====================================================
Private Sub ActualizarCelda(ByVal fila As Long, ByVal col As Long, ByVal valor As String)
    EscribirLog "ActualizarCelda", "fila=" & fila & ", col=" & col & ", valor=" & valor
    If fila < 1 Or fila > m_Count Then Exit Sub
    With m_Eventos(fila)
        Select Case col
            Case COL_ID:             .Id = CLng(Val(valor))
            Case COL_HABILITADO:     .habilitado = (Trim$(valor) <> "")
            Case COL_RUTA:           .Ruta = valor
            Case COL_PARAMETROS:     .parametros = valor
            Case COL_TEXTO_VENTANA:  .TextoVentana = valor
            Case COL_TEXTO_SUBCLASE: .TextoSubClase = valor
            Case COL_EVENTO:         .Evento = CLng(Val(valor))
            Case COL_TEXTO:          .Texto = valor
            Case COL_CLICK_X:        .ClickPosX = CLng(Val(valor))
            Case COL_CLICK_Y:        .ClickPosY = CLng(Val(valor))
            Case COL_MODO_ENVIO:     .ModoEnvioTexto = CLng(Val(valor))
            Case COL_TIEMPO_ESPERA:  .tiempoEspera = CLng(Val(valor))
        End Select
    End With

    If col = COL_HABILITADO Then
        RefrescarCeldaHabilitado fila
    ElseIf col = COL_EVENTO Then
        RefrescarCeldaEvento fila
    Else
        DgEventos.TextMatrix(fila, col) = valor
    End If
    
    MostrarEvento fila
End Sub


' ====================================================
'  OPTIONBUTTONS DE TIPO DE EVENTO
' ====================================================
Private Sub ActualizarEvento(ByVal nuevoEvento As eEvento)
    EscribirLog "ActualizarEvento", "nuevoEvento=" & nuevoEvento
    If m_CargandoFila Then Exit Sub
    Dim fila As Long
    fila = DgEventos.Row
    If fila = 0 Or fila > m_Count Then Exit Sub
    m_Eventos(fila).Evento = nuevoEvento
    RefrescarCeldaEvento fila
    SincronizarFrameModo fila
End Sub

Private Sub SincronizarFrameModo(ByVal fila As Long)
    EscribirLog "SincronizarFrameModo", "fila=" & fila
    If EsEventoTexto(m_Eventos(fila).Evento) Then
        If m_Eventos(fila).ModoEnvioTexto = 0 Then
            m_Eventos(fila).ModoEnvioTexto = MODO_PULSAR_TECLA
        End If
        m_CargandoFila = True
        optPulsarTecla.Value = (m_Eventos(fila).ModoEnvioTexto = MODO_PULSAR_TECLA)
        optPegarTexto.Value = (m_Eventos(fila).ModoEnvioTexto = MODO_PEGAR_TEXTO)
        m_Eventos(fila).IntervaloTecla = CLng(Val(txtIntervaloTecla.Text))
        m_CargandoFila = False
        fraModoTexto.Visible = True
    Else
        fraModoTexto.Visible = False
    End If
End Sub

Private Sub optClickIzquierdo_Click()
    EscribirLog "optClickIzquierdo_Click"
    ActualizarEvento evClickIzquierdo
End Sub

Private Sub OptClickDerecho_Click()
    EscribirLog "OptClickDerecho_Click"
    ActualizarEvento evClickDerecho
End Sub

Private Sub OptDobleClick_Click()
    EscribirLog "OptDobleClick_Click"
    ActualizarEvento evDobleClick
End Sub

Private Sub optCerrarVentana_Click()
    EscribirLog "optCerrarVentana_Click"
    ActualizarEvento evCerrarVentana
End Sub

Private Sub OptTiempoEspera_Click()
    EscribirLog "OptTiempoEspera_Click"
    ActualizarEvento evTiempoEspera
End Sub

Private Sub OptNinguno_Click()
    EscribirLog "OptNinguno_Click"
    ActualizarEvento evNinguno
End Sub

Private Sub OptTexto_Click()
    EscribirLog "OptTexto_Click"
    ActualizarEvento evTexto
End Sub

Private Sub OptTab_Click()
    EscribirLog "OptTab_Click"
    ActualizarEvento evTab
End Sub

Private Sub OptEnter_Click()
    EscribirLog "OptEnter_Click"
    ActualizarEvento evEnter
End Sub

' ====================================================
'  OPTIONBUTTONS DE MODO DE ENVÍO
' ====================================================
Private Sub optPulsarTecla_Click()
    EscribirLog "optPulsarTecla_Click"
    If m_CargandoFila Then Exit Sub
    Dim fila As Long
    fila = DgEventos.Row
    If fila < 1 Or fila > m_Count Then Exit Sub
    m_Eventos(fila).ModoEnvioTexto = MODO_PULSAR_TECLA
    RefrescarCeldaModoEnvio fila
    optPulsarTecla.Caption = "Pulsar tecla cada:"
    txtIntervaloTecla.Visible = True
    lblIntervaloTecla.Visible = True
End Sub

Private Sub optPegarTexto_Click()
    EscribirLog "optPegarTexto_Click"
    If m_CargandoFila Then Exit Sub
    Dim fila As Long
    fila = DgEventos.Row
    If fila < 1 Or fila > m_Count Then Exit Sub
    m_Eventos(fila).ModoEnvioTexto = MODO_PEGAR_TEXTO
    RefrescarCeldaModoEnvio fila
    optPulsarTecla.Caption = "Pulsar tecla"
    txtIntervaloTecla.Visible = False
    lblIntervaloTecla.Visible = False
End Sub

' ====================================================
'  TEXTBOXES DE DETALLE
' ====================================================
Private Sub txtIntervaloTecla_Change()
    EscribirLog "txtIntervaloTecla_Change", "Valor=" & txtIntervaloTecla.Text
    If m_CargandoFila Then Exit Sub
    Dim fila As Long
    fila = DgEventos.Row
    If fila = 0 Or fila > m_Count Then Exit Sub
    m_Eventos(fila).IntervaloTecla = CLng(Val(txtIntervaloTecla.Text))
    'DgEventos.TextMatrix(fila, COL_TEXTO) = txtIntervaloTecla.Text ''' todavía no lo agrego a la grilla
End Sub


Private Sub txtTexto_Change()
    EscribirLog "txtTexto_Change", "Texto=" & txtTexto.Text
    If m_CargandoFila Then Exit Sub
    Dim fila As Long
    fila = DgEventos.Row
    If fila = 0 Or fila > m_Count Then Exit Sub
    m_Eventos(fila).Texto = txtTexto.Text
    DgEventos.TextMatrix(fila, COL_TEXTO) = txtTexto.Text
End Sub

Private Sub txtCx_Change()
    EscribirLog "txtCx_Change", "Valor=" & txtCx.Text
    If m_CargandoFila Then Exit Sub
    Dim fila As Long
    fila = DgEventos.Row
    If fila = 0 Or fila > m_Count Then Exit Sub
    m_Eventos(fila).ClickPosX = CLng(Val(txtCx.Text))
    DgEventos.TextMatrix(fila, COL_CLICK_X) = txtCx.Text
End Sub

Private Sub TxtCy_Change()
    EscribirLog "TxtCy_Change", "Valor=" & TxtCy.Text
    If m_CargandoFila Then Exit Sub
    Dim fila As Long
    fila = DgEventos.Row
    If fila = 0 Or fila > m_Count Then Exit Sub
    m_Eventos(fila).ClickPosY = CLng(Val(TxtCy.Text))
    DgEventos.TextMatrix(fila, COL_CLICK_Y) = TxtCy.Text
End Sub

Private Sub ChkXY_Click()
    EscribirLog "ChkXY_Click", "Valor=" & ChkXY.Value
    txtCx.Enabled = ChkXY.Value
    TxtCy.Enabled = ChkXY.Value
End Sub

' ====================================================
'  EDITORES INLINE SOBRE LA GRILLA
' ====================================================
Private Sub DgEventos_DblClick()
    EscribirLog "DgEventos_DblClick", "Row=" & DgEventos.Row & ", Col=" & DgEventos.col
    Dim fila As Long, col As Long
    fila = DgEventos.Row
    col = DgEventos.col
    If fila = 0 Or col = 0 Then Exit Sub

    OcultarEditores

    ' Mostrar editor genérico según columna
    MostrarEditor fila, col

    ' Acciones especiales
    Select Case col
        Case COL_RUTA
            cmdExplorar_Click
    End Select
End Sub

Private Sub MostrarEditor(ByVal fila As Long, ByVal col As Long)
    If fila < 1 Or fila > m_Count Then Exit Sub
    OcultarEditores

    Select Case col
        Case COL_HABILITADO
            ' Checkbox
            chkEditor.Value = IIf(m_Eventos(fila).habilitado, vbChecked, vbUnchecked)
            PosicionarControlSobreCelda chkEditor, fila, col, 18, 18
            chkEditor.SetFocus

        Case COL_EVENTO
            ' Combo de eventos
            SeleccionarEventoEnCombo m_Eventos(fila).Evento
            PosicionarControlSobreCelda cboEditor, fila, col
            cboEditor.SetFocus

        Case COL_MODO_ENVIO
            ' Solo aplica si es Texto
            If m_Eventos(fila).Evento = evTexto Then
                cboModoEnvio.Clear
                cboModoEnvio.AddItem "Pulsar tecla": cboModoEnvio.ItemData(cboModoEnvio.NewIndex) = MODO_PULSAR_TECLA
                cboModoEnvio.AddItem "Pegar texto": cboModoEnvio.ItemData(cboModoEnvio.NewIndex) = MODO_PEGAR_TEXTO
                ' Seleccionar actual
                Dim i As Long
                For i = 0 To cboModoEnvio.ListCount - 1
                    If cboModoEnvio.ItemData(i) = m_Eventos(fila).ModoEnvioTexto Then
                        cboModoEnvio.ListIndex = i
                        Exit For
                    End If
                Next i
                PosicionarControlSobreCelda cboModoEnvio, fila, col
                cboModoEnvio.SetFocus
            End If

        Case COL_TIEMPO_ESPERA
            ' Solo si el evento es Esperar
            If m_Eventos(fila).Evento = evTiempoEspera Then
                Dim X As Single, Y As Single, w As Single, h As Single
                ObtenerRectCelda fila, col, X, Y, w, h
                txtEditor.Move X, Y, w, h
                txtEditor.Text = CStr(m_Eventos(fila).tiempoEspera)
                txtEditor.Visible = True
                udTiempoEspera.Move X + w - udTiempoEspera.Width, Y, 20, h
                udTiempoEspera.Visible = True
                txtEditor.SetFocus
            End If

        Case Else
            ' TextBox genérico para otros campos editables
            PosicionarControlSobreCelda txtEditor, fila, col
            txtEditor.Text = DgEventos.TextMatrix(fila, col)
            txtEditor.SetFocus
    End Select
End Sub

Private Sub txtEditor_KeyDown(KeyCode As Integer, Shift As Integer)
    EscribirLog "txtEditor_KeyDown", "KeyCode=" & KeyCode & ", Shift=" & Shift
    If KeyCode = vbKeyReturn Then
        ActualizarCelda DgEventos.Row, DgEventos.col, txtEditor.Text
        txtEditor.Visible = False
    ElseIf KeyCode = vbKeyEscape Then
        txtEditor.Visible = False
    End If
End Sub

Private Sub txtEditor_LostFocus()
    EscribirLog "txtEditor_LostFocus"
    txtEditor.Visible = False
End Sub

Private Sub cboModoEnvio_LostFocus()
    EscribirLog "cboModoEnvio_LostFocus"
    cboModoEnvio.Visible = False
End Sub

Private Sub chkEditor_LostFocus()
    EscribirLog "chkEditor_LostFocus"
    chkEditor.Visible = False
End Sub

Private Sub cboEditor_LostFocus()
    EscribirLog "cboEditor_LostFocus"
    cboEditor.Visible = False
End Sub

Private Sub chkEditor_Click()
    EscribirLog "chkEditor_Click", "Value=" & chkEditor.Value
    If m_CargandoFila Then Exit Sub
    Dim fila As Long
    fila = DgEventos.Row
    If fila < 1 Or fila > m_Count Then Exit Sub
    m_Eventos(fila).habilitado = (chkEditor.Value = vbChecked)
    RefrescarCeldaHabilitado fila
    DgEventos.Row = fila
    DgEventos.col = COL_HABILITADO
    DgEventos.CellFontName = "Wingdings"
End Sub

Private Sub cboEditor_Click()
    EscribirLog "cboEditor_Click", "ListIndex=" & cboEditor.ListIndex
    If m_CargandoFila Then Exit Sub
    Dim fila As Long
    fila = DgEventos.Row
    If fila < 1 Or fila > m_Count Then Exit Sub
    If cboEditor.ListIndex < 0 Then Exit Sub
    Dim nuevoEvento As Long
    nuevoEvento = EventoDesdeIndiceCombo(cboEditor.ListIndex)
    m_Eventos(fila).Evento = nuevoEvento
    RefrescarCeldaEvento fila
    MostrarEvento fila
End Sub

' ====================================================
'  INTERACCIÓN CON LA GRILLA (mouse / paint)
' ====================================================
Private Sub DgEventos_PaintCells()
    Dim fila As Long, col As Long
    fila = DgEventos.Row
    col = DgEventos.col
    If fila = 0 Then Exit Sub

    If UltFila > 0 _
       And UltFila <= DgEventos.Rows - 1 _
       And UltCol <= DgEventos.Cols - 1 Then
        DgEventos.Row = UltFila
        DgEventos.col = UltCol
        DgEventos.CellBackColor = RGB(240, 240, 240)
        DgEventos.CellForeColor = RGB(30, 30, 30)
        If UltCol = COL_HABILITADO Then DgEventos.CellFontName = "Wingdings"
    End If

    DgEventos.Row = fila
    DgEventos.col = col
    DgEventos.CellBackColor = RGB(173, 216, 230)
    DgEventos.CellForeColor = vbBlack
    If col = COL_HABILITADO Then DgEventos.CellFontName = "Wingdings"

    UltFila = fila
    UltCol = col
End Sub

Private Sub DgEventos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim fila As Long, col As Long
    fila = DgEventos.MouseRow
    col = DgEventos.MouseCol
    If fila <= 0 Then Exit Sub
    DgEventos.Row = fila
    DgEventos.col = col
    Select Case col
        Case COL_HABILITADO
            AlternarHabilitado fila
        Case COL_EVENTO
            MostrarEditorComboEvento fila
        Case COL_MODO_ENVIO
            MostrarEditorComboModoEnvio fila
        Case COL_TIEMPO_ESPERA
            MostrarEditorTiempoEspera fila
        Case COL_RUTA
            'cmdExplorar_Click
        Case Else
            OcultarEditores
    End Select
End Sub

Private Sub DgEventos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DgEventos.MouseCol = COL_RUTA And DgEventos.MouseRow > 0 Then
        DgEventos.MousePointer = 3
    Else
        DgEventos.MousePointer = 0
    End If
End Sub


' ====================================================
'  EXPLORAR ARCHIVO
' ====================================================
Private Sub cmdExplorar_Click()
    On Error Resume Next
    cdg.CancelError = True
    cdg.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    cdg.Filter = "Programas y scripts|*.exe;*.lnk;*.bat;*.vbs"
    cdg.ShowOpen
    If Err.Number <> 0 Or Trim$(cdg.Filename) = "" Then Err.Clear: Exit Sub
    Dim fila As Long
    fila = DgEventos.Row
    If fila < 1 Or fila > m_Count Then Exit Sub
    m_Eventos(fila).Ruta = cdg.Filename
    DgEventos.TextMatrix(fila, COL_RUTA) = cdg.Filename
End Sub


' ====================================================
'  MENÚS
' ====================================================
Private Sub mnuAP_Click()
    On Error GoTo Salir
    cdg.CancelError = True
    cdg.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    cdg.Filter = "Secuencia (*.sca)|*.sca"
    cdg.ShowOpen
    If Trim$(cdg.Filename) = "" Then Exit Sub
    AbrirProyecto cdg.Filename
    AgregarMenuReciente cdg.Filename
Salir:
End Sub

Private Sub mnuGP_Click()
    If Trim$(m_Archivo) = "" Then
        cdg.CancelError = True
        cdg.flags = cdlOFNOverwritePrompt
        cdg.Filter = "Secuencia (*.sca)|*.sca"
        cdg.DefaultExt = "sca"
        cdg.ShowSave
        If Trim$(cdg.Filename) = "" Then Exit Sub
        m_Archivo = cdg.Filename
        Me.Caption = Me.Tag & " - [" & m_Archivo & "]"
        AgregarMenuReciente m_Archivo
    End If
    GuardarProyecto
End Sub

Private Sub mnuGPC_Click()
    On Error GoTo Salir
    cdg.CancelError = True
    cdg.flags = cdlOFNOverwritePrompt
    cdg.Filter = "Secuencia (*.sca)|*.sca"
    cdg.DefaultExt = "sca"
    cdg.ShowSave
    If Trim$(cdg.Filename) = "" Then Exit Sub
    m_Archivo = cdg.Filename
    GuardarProyecto
    Me.Caption = Me.Tag & " - [" & m_Archivo & "]"
    AgregarMenuReciente m_Archivo
Salir:
End Sub

Private Sub mnuNuevo_Click()
    LimpiarProyecto
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub


' ====================================================
'  GUARDAR / LIMPIAR PROYECTO
' ====================================================
Private Sub GuardarProyecto()
    On Error GoTo ErrHandler
    If Trim$(m_Archivo) = "" Then Exit Sub
    GuardarEnArchivo m_Archivo
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub LimpiarProyecto()
    InicializarArray
    CargarEventosDesdeArray
    fraModoTexto.Visible = False
    m_Archivo = ""
    Me.Caption = Me.Tag
End Sub


' ====================================================
'  VERIFICAR RUTAS
' ====================================================
Private Function VerificarRutas() As Boolean
    VerificarRutas = True
    If m_Count = 0 Then Exit Function
    Dim i As Long
    For i = 1 To m_Count
        If Not m_Eventos(i).habilitado Then GoTo Siguiente
        If Trim$(m_Eventos(i).Ruta) <> "" Then
            If Dir$(m_Eventos(i).Ruta) = "" Then
                MsgBox "No existe:" & vbCrLf & m_Eventos(i).Ruta, vbCritical
                VerificarRutas = False
                Exit Function
            End If
        End If
Siguiente:
    Next i
End Function


' ====================================================
'  EJECUTAR SECUENCIA
' ====================================================
Private Sub cmdSec_Click()
    If m_Count = 0 Then Exit Sub
    If Not VerificarRutas Then Exit Sub
    ForzarMinusculas
    EjecutarSecuencia
    If chkCerrarApp.Value Then KillProcessByPID ObtenerPID(Me.hwnd)
End Sub

Public Sub EjecutarSecuencia()
    Dim hwndVentana As Long, TiempoInicio As Long, i As Long
    On Error GoTo ErrHandler
    EscribirLog "EjecutarSecuencia_Inicio", "TotalEventos=" & m_Count

    Screen.MousePointer = vbHourglass
    BlockInput m_Bloquear

    ' Guardar layout original de tu aplicación
    GuardarLayoutVentana Me.hwnd

    i = 1
    Do While i <= m_Count

        If Not m_Eventos(i).habilitado Then
            EscribirLog "EjecutarSecuencia_SkipEvento", "Evento=" & i
            GoTo ContinuarLoop
        End If

        m_FilaActual = i
        EscribirLog "EjecutarSecuencia_Evento", "FilaActual=" & m_FilaActual & ", TipoEvento=" & m_Eventos(i).Evento

        Select Case m_Eventos(i).Evento
            Case evEnter, evTab, evCerrarVentana, evClickIzquierdo, evClickDerecho, evDobleClick
 
                ' Eventos directos
                EjecutarAccion hwndVentana

            Case evTiempoEspera
                EscribirLog "EjecutarSecuencia_TiempoEspera", "Segundos=" & m_Eventos(i).tiempoEspera
                Sleep m_Eventos(i).tiempoEspera * 1000

            Case Else
                ' Eventos que requieren ventana
                EjecutarAplicacionActual
                TiempoInicio = GetTickCount
                Do
                    hwndVentana = BuscarVentanaActual
                    ' Poner foco en la ventana y cambiar layout
                    SetForegroundWindow hwndVentana
                    CambiarLayoutVentana hwndVentana
                    If hwndVentana <> 0 Then
                        EjecutarAccion hwndVentana
                        Exit Do
                    End If
                    DoEvents

                    If GetTickCount - TiempoInicio > TIMEOUT_VENTANA Then
                        EscribirLog "EjecutarSecuencia_Timeout", "Evento=" & i
                        MsgBox "Timeout esperando ventana", vbExclamation
                        Exit Do
                    End If
                Loop
        End Select

ContinuarLoop:
        i = i + 1
    Loop

    ' Al terminar toda la secuencia: restaurar layout original y devolver foco
    Me.SetFocus
    RestaurarLayoutVentana
    SetForegroundWindow Me.hwnd

Salir:
    BlockInput False
    Screen.MousePointer = vbDefault
    EscribirLog "EjecutarSecuencia_Fin"
    Exit Sub

ErrHandler:
    EscribirLog "EjecutarSecuencia_Error", "Error=" & Err.Description
    MsgBox Err.Description, vbCritical
    Resume Salir
End Sub


' ====================================================
'  EJECUTAR ACCIÓN
' ====================================================
Private Sub EjecutarAccion(ByVal hwndObjetivo As Long)
    If m_FilaActual < 1 Or m_FilaActual > m_Count Then Exit Sub
    Dim ev As TEvento
    ev = m_Eventos(m_FilaActual)
 
    EscribirLog "EjecutarAccion", _
        "hwndVentana=" & hwndObjetivo & _
        ", Texto=" & ev.Texto & _
        ", ModoEnvioTexto=" & ev.ModoEnvioTexto

    If hwndObjetivo <> 0 Then m_EnumWin.SetActiveWindows hwndObjetivo

    Select Case ev.Evento
        Case evClickIzquierdo: EjecutarClick hwndObjetivo, eLeftClick
        Case evClickDerecho:   EjecutarClick hwndObjetivo, eRightClick
        Case evDobleClick:     EjecutarClick hwndObjetivo, eDoubleLeft
        Case evCerrarVentana:  KillProcessByPID m_LastPID
        Case evTab:            EnviarTecla vbKeyTab
        Case evEnter:          EnviarTecla vbKeyReturn
    End Select

    If ev.Evento = evTexto And Trim$(ev.Texto) <> "" Then
        Select Case ev.ModoEnvioTexto
            Case MODO_PULSAR_TECLA:
                PegarTexto hwndObjetivo, ev.Texto, ev.IntervaloTecla
            Case MODO_PEGAR_TEXTO
                CopiarAlPortapapeles ev.Texto
                SimularCtrlV
        End Select
    End If
End Sub

Private Sub EjecutarClick(ByVal hwndObjetivo As Long, ByVal TipoClick As Long)
    If m_FilaActual < 1 Or m_FilaActual > m_Count Then Exit Sub
    Dim px As Long, py As Long
    px = m_Eventos(m_FilaActual).ClickPosX
    py = m_Eventos(m_FilaActual).ClickPosY
    EscribirLog "EjecutarClick", "hwndVentana=" & hwndObjetivo & ", X=" & px & ", Y=" & py
    If px = 0 And py = 0 Then Exit Sub
    EnviarClickXY px, py, TipoClick
End Sub

Private Sub EjecutarAplicacionActual()
    EscribirLog "EjecutarAplicacionActual", "FilaActual=" & m_FilaActual & ", Ruta=" & m_Eventos(m_FilaActual).Ruta & ", Parametros=" & m_Eventos(m_FilaActual).parametros
    If m_FilaActual < 1 Or m_FilaActual > m_Count Then Exit Sub
    Dim Ruta As String, Params As String
    Ruta = Trim$(m_Eventos(m_FilaActual).Ruta)
    Params = Trim$(m_Eventos(m_FilaActual).parametros)
    If Ruta = "" Then Exit Sub
    m_LastPID = EjecutarConPID(Ruta, Params)
End Sub


' ====================================================
'  BUSCAR VENTANA
' ====================================================
Private Function BuscarVentanaActual() As Long
    If m_LastPID = 0 Then
        BuscarVentanaActual = 0
    Else
        BuscarVentanaActual = FindWindowByProcessID(m_LastPID)
    End If
End Function

Private Sub txtTiempoEspera_Change()
    EscribirLog "txtTiempoEspera_Change", "Valor=" & txtTiempoEspera.Text
    If m_CargandoFila Then Exit Sub
    Dim fila As Long
    fila = DgEventos.Row
    If fila = 0 Or fila > m_Count Then Exit Sub
    m_Eventos(fila).tiempoEspera = CLng(Val(txtTiempoEspera.Text))
    DgEventos.TextMatrix(fila, COL_TIEMPO_ESPERA) = txtTiempoEspera.Text
End Sub

Private Sub udTiempoEspera_Change()
    Dim fila As Long
    fila = DgEventos.Row
    If fila < 1 Or fila > m_Count Then Exit Sub

    m_Eventos(fila).tiempoEspera = udTiempoEspera.Value
    DgEventos.TextMatrix(fila, COL_TIEMPO_ESPERA) = CStr(udTiempoEspera.Value)
    txtEditor.Text = CStr(udTiempoEspera.Value)   ' sincroniza también el textbox
End Sub

' ====================================================
'  HOTKEY / TIMER
' ====================================================
Private Sub VBHotKey1_HotkeyPressed()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText txtCaption.Text
End Sub

Private Sub Timer1_Timer()
    ActualizarInformacionVentana
End Sub

Private Sub ActualizarInformacionVentana()
    Dim pt As POINTAPI, hwndActual As Long
    GetCursorPos pt
    hwndActual = WindowFromPoint(pt.X, pt.Y)
    txtX.Text = pt.X
    txtY.Text = pt.Y
    txthwnd.Text = hwndActual
    txtCaption.Text = m_EnumWin.GetWindowCaption(hwndActual)
End Sub


' ====================================================
'  HELPERS
' ====================================================
Private Function Nz(ByVal valor As Variant, _
                    Optional ByVal DefaultValue As String = "") As String
    If IsNull(valor) Then Nz = DefaultValue Else Nz = Trim$(CStr(valor))
End Function

Private Function ObtenerUnidad(ByVal Unidad As String) As String
    Select Case Trim$(Unidad)
        Case "1": ObtenerUnidad = Left$(App.Path, 2)
        Case "2": ObtenerUnidad = ObtenerWinDir()
        Case Else: ObtenerUnidad = Unidad
    End Select
End Function


