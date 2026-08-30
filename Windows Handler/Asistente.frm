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
   ClientHeight    =   10860
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   15120
   Icon            =   "Asistente.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   724
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1008
   Tag             =   "Zinisoft Windows Handler"
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid DgEventos 
      Height          =   4305
      Left            =   180
      TabIndex        =   44
      Top             =   150
      Width           =   28440
      _ExtentX        =   50165
      _ExtentY        =   7594
      _Version        =   393216
      BackColorBkg    =   16777215
      Appearance      =   0
   End
   Begin BotonXp.BotónXp CmdNuevo 
      Height          =   390
      Left            =   180
      TabIndex        =   39
      Top             =   4560
      Width           =   1300
      _ExtentX        =   2302
      _ExtentY        =   688
      BTYPE           =   14
      TX              =   "&Nuevo"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   15790320
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
   Begin BotonXp.BotónXp CmdInsertar 
      Height          =   390
      Left            =   1560
      TabIndex        =   40
      Top             =   4560
      Width           =   1300
      _ExtentX        =   2302
      _ExtentY        =   688
      BTYPE           =   14
      TX              =   "&Insertar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   15790320
      MPTR            =   1
      MICON           =   "Asistente.frx":3476
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
      Height          =   390
      Left            =   2940
      TabIndex        =   41
      Top             =   4560
      Width           =   1300
      _ExtentX        =   2302
      _ExtentY        =   688
      BTYPE           =   14
      TX              =   "&Eliminar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   15790320
      MPTR            =   1
      MICON           =   "Asistente.frx":3492
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
      Height          =   390
      Left            =   4380
      TabIndex        =   42
      Top             =   4560
      Width           =   2200
      _ExtentX        =   3889
      _ExtentY        =   688
      BTYPE           =   14
      TX              =   "&Comenzar secuencia"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   14196792
      BCOLO           =   12608000
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   14196792
      MPTR            =   1
      MICON           =   "Asistente.frx":34AE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox ChkBloq 
      BackColor       =   &H80000018&
      Caption         =   "Bloquear mouse y teclado durante ejecuciones largas"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   6800
      TabIndex        =   36
      Top             =   4620
      Width           =   4800
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Configuración del Evento Seleccionado "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5500
      Left            =   180
      TabIndex        =   0
      Top             =   5050
      Width           =   9800
      Begin VB.OptionButton optClickIzquierdo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Click Izquierdo"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1400
      End
      Begin VB.OptionButton OptClickDerecho 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Click Derecho"
         Height          =   255
         Left            =   1700
         TabIndex        =   2
         Top             =   360
         Width           =   1400
      End
      Begin VB.OptionButton OptDobleClick 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Doble Click"
         Height          =   255
         Left            =   3100
         TabIndex        =   16
         Top             =   360
         Width           =   1300
      End
      Begin VB.OptionButton OptNinguno 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ninguno"
         Height          =   255
         Left            =   4500
         TabIndex        =   52
         Top             =   360
         Width           =   1100
      End
      Begin VB.OptionButton OptTiempoEspera 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Esperar"
         Height          =   255
         Left            =   5700
         TabIndex        =   53
         Top             =   360
         Width           =   900
      End
      Begin VB.TextBox txtTiempoEspera 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   6650
         TabIndex        =   54
         Top             =   330
         Width           =   700
      End
      Begin VB.CheckBox ChkXY 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Coordenadas dentro de la ventana:"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   780
         Width           =   2900
      End
      Begin VB.TextBox txtCx 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3450
         MaxLength       =   4
         TabIndex        =   21
         Top             =   750
         Width           =   600
      End
      Begin VB.TextBox TxtCy 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4400
         MaxLength       =   4
         TabIndex        =   22
         Top             =   750
         Width           =   600
      End
      Begin VB.OptionButton OptTexto 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enviar Texto:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   1300
      End
      Begin VB.TextBox txtTexto 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2800
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   17
         Top             =   1500
         Width           =   9300
      End
      Begin VB.Frame fraModoTexto 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Modo de envío"
         Height          =   850
         Left            =   240
         TabIndex        =   47
         Top             =   4450
         Width           =   4800
         Begin VB.OptionButton optPulsarTecla 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pulsar tecla"
            Height          =   255
            Left            =   200
            TabIndex        =   48
            Top             =   300
            Width           =   1200
         End
         Begin VB.TextBox txtIntervaloTecla 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1450
            TabIndex        =   58
            Text            =   "100"
            Top             =   270
            Visible         =   0   'False
            Width           =   550
         End
         Begin VB.OptionButton optPegarTexto 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pegar texto (Clipboard)"
            Height          =   255
            Left            =   2600
            TabIndex        =   49
            Top             =   300
            Width           =   2000
         End
         Begin VB.Label lblIntervaloTecla 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ms."
            Height          =   255
            Left            =   2050
            TabIndex        =   59
            Top             =   300
            Visible         =   0   'False
            Width           =   400
         End
      End
      Begin VB.OptionButton OptTab 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enviar TAB"
         Height          =   255
         Left            =   5200
         TabIndex        =   18
         Top             =   4550
         Width           =   1300
      End
      Begin VB.OptionButton OptEnter 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enviar ENTER"
         Height          =   255
         Left            =   6600
         TabIndex        =   43
         Top             =   4550
         Width           =   1400
      End
      Begin VB.OptionButton optCerrarVentana 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cerrar ventana"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   8100
         TabIndex        =   15
         Top             =   4550
         Width           =   1400
      End
      Begin VB.CheckBox chkCerrarApp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cerrar la aplicación al finalizar secuencia"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   5200
         TabIndex        =   46
         Top             =   4950
         Width           =   4300
      End
      Begin VB.Label lblSg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "seg."
         Height          =   255
         Left            =   7420
         TabIndex        =   55
         Top             =   380
         Width           =   400
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "X:"
         Height          =   255
         Left            =   3250
         TabIndex        =   19
         Top             =   780
         Width           =   200
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y:"
         Height          =   255
         Left            =   4200
         TabIndex        =   20
         Top             =   780
         Width           =   200
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Observador de Clases "
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3300
      Left            =   10100
      TabIndex        =   3
      Top             =   5050
      Width           =   4840
      Begin VB.TextBox txthwnd 
         Height          =   285
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   350
         Width           =   1300
      End
      Begin VB.TextBox txthwndp 
         Height          =   285
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   1300
      End
      Begin VB.TextBox txtNC 
         Height          =   285
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1100
         Width           =   3200
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1480
         Width           =   3200
      End
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1900
         Width           =   700
      End
      Begin VB.TextBox txtY 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1900
         Width           =   700
      End
      Begin VB.TextBox txtRelX 
         Height          =   285
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2300
         Width           =   700
      End
      Begin VB.TextBox txtRelY 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2300
         Width           =   700
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hwnd:"
         Height          =   255
         Index           =   0
         Left            =   200
         TabIndex        =   4
         Top             =   380
         Width           =   1100
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hwnd Parent:"
         Height          =   255
         Left            =   200
         TabIndex        =   6
         Top             =   750
         Width           =   1100
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Clase:"
         Height          =   255
         Left            =   200
         TabIndex        =   8
         Top             =   1130
         Width           =   1100
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Caption:"
         Height          =   255
         Left            =   200
         TabIndex        =   10
         Top             =   1510
         Width           =   1100
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cursor General:"
         Height          =   255
         Left            =   200
         TabIndex        =   24
         Top             =   1930
         Width           =   1100
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "X"
         Height          =   255
         Left            =   2150
         TabIndex        =   12
         Top             =   1930
         Width           =   200
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y"
         Height          =   255
         Left            =   3150
         TabIndex        =   13
         Top             =   1930
         Width           =   200
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pos. Relativa:"
         Height          =   255
         Left            =   200
         TabIndex        =   25
         Top             =   2330
         Width           =   1100
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "X"
         Height          =   255
         Left            =   2150
         TabIndex        =   28
         Top             =   2330
         Width           =   200
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Y"
         Height          =   255
         Left            =   3150
         TabIndex        =   27
         Top             =   2330
         Width           =   200
      End
   End
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
      Left            =   10100
      Top             =   8500
      _ExtentX        =   794
      _ExtentY        =   794
      VKey            =   116
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   10600
      Top             =   8500
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
         Caption         =   $"Asistente.frx":34CA
         Enabled         =   0   'False
         Height          =   420
         Left            =   270
         TabIndex        =   32
         Top             =   315
         Width           =   6900
      End
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   11100
      Top             =   8500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "sca"
      Filter          =   "Archivos de secuencia de comandos automatizados (*.sca)|*.sca"
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   11600
      Top             =   8500
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   180
      TabIndex        =   60
      Top             =   10550
      Width           =   2000
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   12500
      Picture         =   "Asistente.frx":3585
      Stretch         =   -1  'True
      Top             =   8500
      Visible         =   0   'False
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
      Begin VB.Menu mnuTmp 
         Caption         =   ""
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTmp 
         Caption         =   ""
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTmp 
         Caption         =   ""
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTmp 
         Caption         =   ""
         Index           =   8
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
Private Const MAX_RECIENTES   As Long = 8
Private Const TIMEOUT_VENTANA As Long = 2500
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
    Habilitado     As Boolean
    Ruta           As String
    Parametros     As String
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
Public g_BorrarLog As Boolean
Public g_MostrarTimoutVentana As Boolean
Private HayCambios As Boolean
Private m_EventosOriginal() As TEvento
Private LayoutOriginal As String

Private Const LOG_FILE As String = "WindowsHandler.log"

Private Sub EscribirLog(ByVal accion As String, Optional ByVal Parametros As String = "")
    On Error Resume Next   ' evita que un error en el log dispare ErrHandler
    If Not g_GrabarLog Then Exit Sub  ' si está desactivado, no hace nada
    
    Dim f As Integer
    f = FreeFile
    On Error Resume Next
    Open App.Path & "\" & LOG_FILE For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & accion & _
              IIf(Parametros <> "", " | " & Parametros, "")
    Close #f
    On Error GoTo 0
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
        .Habilitado = True
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
    nuevo.Habilitado = True
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
    ReDim m_eventos_original(1 To 1)
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

Private Function TextoHabilitado(ByVal Habilitado As Boolean) As String
    If Habilitado Then TextoHabilitado = Chr$(252) Else TextoHabilitado = ""
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
    chkEditor.Value = IIf(m_Eventos(fila).Habilitado, vbChecked, vbUnchecked)
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

    If m_Eventos(fila).Habilitado Then

        DgEventos.TextMatrix(fila, COL_HABILITADO) = "ü"   ' tilde
        'DgEventos.CellForeColor = vbGreen
    Else
        DgEventos.TextMatrix(fila, COL_HABILITADO) = "û"   ' cruz
        'DgEventos.CellForeColor = vbRed
    End If
End Sub

Private Sub AlternarHabilitado(ByVal fila As Long)
    If fila < 1 Or fila > m_Count Then Exit Sub
    m_Eventos(fila).Habilitado = Not m_Eventos(fila).Habilitado
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
            .Habilitado = (CLng(Val(LeeINI(Archivo, sec, "Habilitado", "1"))) <> 0)
            .Ruta = LeeINI(Archivo, sec, "Ruta", "")
            .Parametros = LeeINI(Archivo, sec, "Parametros", "")
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
    
    'Copio el array para comparar luego si hubo cambios
    ReDim m_EventosOriginal(1 To m_Count)
    For i = 1 To m_Count
        m_EventosOriginal(i) = m_Eventos(i)
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
            GrabaINI Archivo, sec, "Habilitado", IIf(.Habilitado, "1", "0")
            GrabaINI Archivo, sec, "Ruta", .Ruta
            GrabaINI Archivo, sec, "Parametros", .Parametros
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

Private Sub cboModoEnvio_Click()
    Dim fila As Long
    fila = DgEventos.Row
    If fila < 1 Or fila > m_Count Then Exit Sub
    m_Eventos(fila).ModoEnvioTexto = cboModoEnvio.ItemData(cboModoEnvio.ListIndex)
    DgEventos.TextMatrix(fila, COL_MODO_ENVIO) = cboModoEnvio.Text
End Sub

Public Sub InicializarPorParametros()
    EscribirLog "IsByParam"
    IsByParam = True
    Timer1.Enabled = False
End Sub

Private Sub Form_Load()
    CargarConfiguracionGeneral
    
    If CargadoPorParametros = True Then
        Me.InicializarPorParametros
    End If
   
    LayoutOriginal = LayoutActualKLID
    EscribirLog "Cargado Layout Teclado [" & LayoutOriginal & "]"
   
    EscribirLog "Form_Load"
    Set m_EnumWin = New clsEnum
    InicializarArray
    InicializarComboEventos
    ConfigurarFlexGrid
    If Not IsByParam Then CargarRecientes
    If Not IsByParam Then AjustarColumnas
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
        .FontBold = True
        .Left = 0
        .Top = Me.ScaleHeight - .Height
    End With
End Sub

Private Sub Form_Terminate()
'    On Error Resume Next
'    EscribirLog "RestaurarLayoutVentana: Terminate()"
'    RestaurarLayoutVentana
End Sub

' ====================================================
'  CONFIGURACIÓN GENERAL DEL PROYECTO ACTUAL
' ====================================================
Private Sub CargarConfiguracionProyecto()
   EscribirLog "CargarConfiguracionProyecto"
    If m_Archivo <> "" Then
        chkCerrarApp.Value = CInt(LeeINI(m_Archivo, "Proyecto", "CerrarAlFinalizar", "0"))
        g_MostrarTimoutVentana = CInt(LeeINI(m_Archivo, "Proyecto", "MostrarVentanaTimeOut", "0"))
    Else
        chkCerrarApp.Value = vbUnchecked
        g_MostrarTimoutVentana = False
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
    Dim valor As String
    
    valor = LeeINI(App.Path & "\" & FILE_INI, "General", "GrabarLog", "1")
    g_GrabarLog = (valor = "1")
    
    valor = LeeINI(App.Path & "\" & FILE_INI, "General", "BorrarLog", "0")
    g_BorrarLog = (valor = "1")
    
    ' Si está activado borrar log y también grabar log
    If g_GrabarLog And g_BorrarLog Then
        On Error Resume Next
        Kill App.Path & "\" & LOG_FILE
        On Error GoTo 0
    End If
End Sub

Private Sub AgregarMenuReciente(ByVal Archivo As String)
    EscribirLog "AgregarMenuReciente", "Archivo=" & Archivo
    Dim i As Long, Existe As Boolean
    Dim rutaExistente As String
    
    ' Verificar si ya existe el archivo en la lista (sin prefijo numérico)
    For i = 1 To MAX_RECIENTES
        If mnuTmp(i).Visible Then
            rutaExistente = Mid$(mnuTmp(i).Caption, InStr(1, mnuTmp(i).Caption, " ") + 1)
            If LCase$(Trim$(rutaExistente)) = LCase$(Trim$(Archivo)) Then
                Existe = True: Exit For
            End If
        End If
    Next i
    
    If Existe Then Exit Sub
    
    ' Desplazar hacia abajo
    For i = MAX_RECIENTES To 2 Step -1
        mnuTmp(i).Caption = mnuTmp(i - 1).Caption
        mnuTmp(i).Visible = mnuTmp(i - 1).Visible
    Next i
    
    ' Insertar en primera posición
    mnuTmp(1).Caption = "1 " & Archivo
    mnuTmp(1).Visible = True
    
    ' Actualizar numeración
    For i = 2 To MAX_RECIENTES
        If mnuTmp(i).Visible Then
            mnuTmp(i).Caption = CStr(i) & " " & Mid$(mnuTmp(i).Caption, InStr(1, mnuTmp(i).Caption, " ") + 1)
        End If
    Next i
    
    GuardarRecientesINI
End Sub

Private Sub GuardarRecientesINI()
    EscribirLog "GuardarRecientesINI"
    Dim i As Long, Archivo As String
    For i = 1 To MAX_RECIENTES
        Archivo = ""
        If mnuTmp(i).Visible Then
            If InStr(1, mnuTmp(i).Caption, " ") > 0 Then
                Archivo = Mid$(mnuTmp(i).Caption, InStr(1, mnuTmp(i).Caption, " ") + 1)
            End If
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
                TextoHabilitado(.Habilitado) & vbTab & _
                .Ruta & vbTab & _
                .Parametros & vbTab & _
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
                optPulsarTecla.Caption = "Pulsar tecla"
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
            Case COL_HABILITADO:     .Habilitado = (Trim$(valor) <> "")
            Case COL_RUTA:           .Ruta = valor
            Case COL_PARAMETROS:     .Parametros = valor
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
    optPulsarTecla.Caption = "Pulsar tecla"
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
            chkEditor.Value = IIf(m_Eventos(fila).Habilitado, vbChecked, vbUnchecked)
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
    m_Eventos(fila).Habilitado = (chkEditor.Value = vbChecked)
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
    If fila < 1 Or fila > m_Count Then Exit Sub
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
    On Error GoTo salir
    cdg.CancelError = True
    cdg.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    cdg.Filter = "Secuencia (*.sca)|*.sca"
    cdg.ShowOpen
    If Trim$(cdg.Filename) = "" Then Exit Sub
    AbrirProyecto cdg.Filename
    AgregarMenuReciente cdg.Filename
salir:
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
    On Error GoTo salir
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
salir:
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
    
    ' REFRESCAR EL ESTADO ORIGINAL EN MEMORIA
    Dim i As Long
    If m_Count > 0 Then
        ReDim m_EventosOriginal(1 To m_Count) As TEvento
        For i = 1 To m_Count
            m_EventosOriginal(i) = m_Eventos(i)
        Next i
    Else
        Erase m_EventosOriginal
    End If
    
    ' Forzar la variable a False
    HayCambios = False
    
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
        If Not m_Eventos(i).Habilitado Then GoTo Siguiente
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
Public Sub CmdSecClick()
    If m_Count = 0 Then Exit Sub
    If Not VerificarRutas Then Exit Sub
    ForzarMinusculas
   
' 1. Cambiar a Inglés para tipear los símbolos
    CambiarTecladoForzado HKL_ENG_US
    DoEvents
    Sleep 100
    
' 2.
    EjecutarSecuencia
    
' 3. IMPORTANTE: Restaurar a Español ANTES de salir de la consola
    CambiarTecladoForzado LayoutOriginal
    DoEvents
    Sleep 100
    
    ' Si la opción de cerrar la app al finalizar
    If chkCerrarApp.Value Then
        If IsByParam Then
            KillProcessByPID ObtenerPID(Me.hwnd)
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub CambiarEstadoControles(Habilitado As Boolean)
    CmdSec.Enabled = Habilitado
    CmdNuevo.Enabled = Habilitado
    CmdInsertar.Enabled = Habilitado
    CmdEliminar.Enabled = Habilitado
    Frame1.Enabled = Habilitado
    Frame2.Enabled = Habilitado
    DgEventos.Enabled = Habilitado
    mnuArchivo.Enabled = Habilitado
End Sub

Private Sub cmdSec_Click()
    CambiarEstadoControles False
    Call CmdSecClick
    CambiarEstadoControles True
End Sub

Public Sub EjecutarSecuencia()
    Dim hwndVentana As Long, i As Long
    Dim hwndNuevaInstancia As Long
    
    On Error GoTo ErrHandler
    EscribirLog "EjecutarSecuencia_Inicio", "TotalEventos=" & m_Count

    LimpiarHistorialVentanas

    Screen.MousePointer = vbHourglass
    BlockInput m_Bloquear

    Dim timeoutGlobal As Long
    timeoutGlobal = CalcularTimeoutGlobal()
   
    i = 1
    Do While i <= m_Count
    
        If Not m_Eventos(i).Habilitado Then
            EscribirLog "EjecutarSecuencia_SkipEvento", "Evento=" & i
            GoTo ContinuarLoop
        End If
    
        m_FilaActual = i
        EscribirLog "EjecutarSecuencia_Evento", "FilaActual=" & m_FilaActual & ", TipoEvento=" & m_Eventos(i).Evento
    
        Select Case m_Eventos(i).Evento
            Case evEnter, evTab, evCerrarVentana, evClickIzquierdo, evClickDerecho, evDobleClick
                If hwndVentana <> 0 Then ActivarVentana hwndVentana
                EjecutarAccion hwndVentana
 
            Case evTiempoEspera
                EscribirLog "EjecutarSecuencia_TiempoEspera", "Segundos=" & m_Eventos(i).tiempoEspera
                Sleep m_Eventos(i).tiempoEspera * 1000
    
            Case Else
                ' Si hay ruta ejecutable, tomamos foto del estado ANTES de abrir la app
                If Len(m_Eventos(i).Ruta) > 0 Then
                    TomarSnapshotVentanas
                    m_LastPID = LanzarProceso(m_Eventos(i).Ruta, m_Eventos(i).Parametros)
                    EscribirLog "LanzarProceso", "PID=" & m_LastPID & ", FilaActual=" & m_FilaActual & ", Ruta=" & m_Eventos(m_FilaActual).Ruta & ", Parametros=" & m_Eventos(m_FilaActual).Parametros
                    
                    ' Esperar la ventana que NO existía en el snapshot
                    hwndNuevaInstancia = EsperarNuevaVentana(timeoutGlobal)
                    If hwndNuevaInstancia <> 0 Then
                        hwndVentana = hwndNuevaInstancia
                        RegistrarHwnd hwndVentana
                    End If
                End If
                  
                ' Evaluación de Timeout / Acción
                If hwndVentana = 0 Or IsWindow(hwndVentana) = 0 Then
                    If g_MostrarTimoutVentana = True Then
                        MsgBox "Timeout esperando ventana", vbExclamation
                        Exit Sub
                    Else
                        EscribirLog "Timeout", "En ventana " & timeoutGlobal & " ms"
                        If chkCerrarApp.Value Then End
                    End If
                Else
                    ' FORZAR FOCO en la ventana detectada
                    ActivarVentana hwndVentana
                    EjecutarAccion hwndVentana
                End If

        End Select
    
ContinuarLoop:
        i = i + 1
    Loop

salir:
    BlockInput False
    Screen.MousePointer = vbDefault
    
    CambiarTecladoForzado LayoutOriginal
    DoEvents
    Sleep 100
    
    EscribirLog "EjecutarSecuencia_Fin"
    Exit Sub

ErrHandler:
    EscribirLog "EjecutarSecuencia_Error", "Error=" & Err.Description
    Resume salir
End Sub

Private Function CalcularTimeoutGlobal() As Long
    Dim i As Long, total As Long
    
    total = 0
    For i = 1 To UBound(m_Eventos)
        total = total + m_Eventos(i).tiempoEspera * 1000  ' convertir a ms
    Next i
    
    ' margen extra de seguridad
    CalcularTimeoutGlobal = total + TIMEOUT_VENTANA
End Function

' ====================================================
'  EJECUTAR ACCIÓN
' ====================================================
Private Sub EjecutarAccion(ByVal hwndObjetivo As Long)
    If m_FilaActual < 1 Or m_FilaActual > m_Count Or hwndObjetivo = 0 Then Exit Sub
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

    ' Hwnd Parent
    Dim hParent As Long
    hParent = GetParent(hwndActual)
    txthwndp.Text = CStr(hParent)

    ' Clase
    Dim sClass As String * 256
    Dim ret As Long
    ret = GetClassName(hwndActual, sClass, 256)
    If ret > 0 Then txtNC.Text = Left$(sClass, ret) Else txtNC.Text = ""

    ' Posición relativa
    Dim rc As RECT
    GetWindowRect hwndActual, rc
    txtRelX.Text = pt.X - rc.Left
    txtRelY.Text = pt.Y - rc.Top
End Sub

' ====================================================
'  HELPERS
' ====================================================
Private Function CompararEvento(ev1 As TEvento, ev2 As TEvento) As Boolean
    ' Usar Trim$ para evitar falsos positivos por espacios o valores vacíos
    CompararEvento = _
        (ev1.Id = ev2.Id) And _
        (ev1.Habilitado = ev2.Habilitado) And _
        (Trim$(ev1.Ruta) = Trim$(ev2.Ruta)) And _
        (Trim$(ev1.Parametros) = Trim$(ev2.Parametros)) And _
        (Trim$(ev1.TextoVentana) = Trim$(ev2.TextoVentana)) And _
        (Trim$(ev1.TextoSubClase) = Trim$(ev2.TextoSubClase)) And _
        (ev1.Evento = ev2.Evento) And _
        (Trim$(ev1.Texto) = Trim$(ev2.Texto)) And _
        (ev1.ClickPosX = ev2.ClickPosX) And _
        (ev1.ClickPosY = ev2.ClickPosY) And _
        (ev1.ModoEnvioTexto = ev2.ModoEnvioTexto) And _
        (ev1.IntervaloTecla = ev2.IntervaloTecla) And _
        (ev1.tiempoEspera = ev2.tiempoEspera)
End Function

Private Function VerificarCambios() As Boolean
    Dim i As Long
    
    On Error GoTo ErrorHandler

    ' 1. Verificar si el array de referencia está inicializado
    If (Not Not m_EventosOriginal) = 0 Then
        ' Si el array original no se ha creado pero m_Count tiene items
        If m_Count > 0 Then HayCambios = True: Exit Function
    Else
        ' Comparar la cantidad de elementos usando el tamaño real (1 To m_Count)
        If m_Count <> UBound(m_EventosOriginal) Then
            HayCambios = True
            Exit Function
        End If
    End If

    ' 2. Comparar evento por evento
    For i = 1 To m_Count
        If Not CompararEvento(m_Eventos(i), m_EventosOriginal(i)) Then
            HayCambios = True
            Exit Function
        End If
    Next i
    
    ' 3. Comparar configuración global del proyecto (Cerrar al finalizar)
    Dim configIni As Long
    configIni = CLng(Val(LeeINI(m_Archivo, "Proyecto", "CerrarAlFinalizar", "0")))
    
    If Abs(chkCerrarApp.Value) <> Abs(configIni) Then
        HayCambios = True
        Exit Function
    End If
    
    ' Si llegó hasta acá, todo es idéntico
    HayCambios = False
    VerificarCambios = False
    Exit Function

ErrorHandler:
    ' Ante cualquier error de lectura/array, asumir que no hay cambios o loguearlo
    HayCambios = False
    VerificarCambios = False
End Function

Private Sub Form_Unload(Cancel As Integer)
    EscribirLog "Form_Unload", "Cancel=" & Cancel
    
    VerificarCambios
    
    ' Verificar si hubo cambios
    If HayCambios Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("¿Desea guardar los cambios antes de salir?", vbYesNoCancel + vbQuestion, "Confirmar salida")
        Select Case resp
            Case vbYes
                GuardarEnArchivo m_Archivo
            Case vbCancel
                Cancel = 1   ' abortar cierre
                Exit Sub
            Case vbNo
                ' salir sin guardar
        End Select
    End If

    ' Limpieza normal
    On Error Resume Next
    Timer1.Enabled = False
    VBHotKey1.StopHotkey
    Set DgEventos.DataSource = Nothing
    Set m_EnumWin = Nothing
    
    EscribirLog "RestaurarLayoutVentana: Unload() [" & LayoutOriginal & "]"
    CambiarTecladoForzado LayoutOriginal
    DoEvents
    Sleep 100
        
    'Terminar proceso completo
    End
End Sub


