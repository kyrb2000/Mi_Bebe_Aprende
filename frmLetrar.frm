VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AgentCtl.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmLetrar 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Letras"
   ClientHeight    =   9195
   ClientLeft      =   2130
   ClientTop       =   1260
   ClientWidth     =   12450
   Icon            =   "frmLetrar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   9960
      TabIndex        =   23
      Top             =   7920
      Width           =   2055
   End
   Begin VB.Timer timTiempo 
      Interval        =   1000
      Left            =   11760
      Top             =   7440
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   15901
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "&Escribir"
      TabPicture(0)   =   "frmLetrar.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Adivinar"
      TabPicture(1)   =   "frmLetrar.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Jugar"
      TabPicture(2)   =   "frmLetrar.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Palabras"
      TabPicture(3)   =   "frmLetrar.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&Imagenes"
      TabPicture(4)   =   "frmLetrar.frx":04B2
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Frame4"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Matematicas"
      TabPicture(5)   =   "frmLetrar.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame6 
         Caption         =   "Frame1"
         Height          =   8535
         Left            =   -74880
         TabIndex        =   54
         Top             =   360
         Width           =   12135
         Begin VB.OptionButton Option1 
            Caption         =   "Suma"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   9840
            TabIndex        =   62
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Resta"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   9840
            TabIndex        =   61
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Multiplicación"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   9840
            TabIndex        =   60
            Top             =   1080
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
            Caption         =   "División"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   9840
            TabIndex        =   59
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox txtRespuesta 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   99.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   360
            TabIndex        =   58
            Top             =   5880
            Width           =   9375
         End
         Begin VB.CommandButton cmdComenzar 
            Caption         =   "&Iniciar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   9840
            TabIndex        =   57
            Top             =   2400
            Width           =   2055
         End
         Begin VB.ComboBox cmbNivel 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            ItemData        =   "frmLetrar.frx":04EA
            Left            =   240
            List            =   "frmLetrar.frx":04FA
            TabIndex        =   56
            Text            =   "Nivel 1"
            Top             =   480
            Width           =   1695
         End
         Begin VB.CheckBox chkRespuesta 
            Caption         =   "Respuesta"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9840
            TabIndex        =   55
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label lblN0 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   99.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   2400
            TabIndex        =   65
            Top             =   600
            Width           =   5295
         End
         Begin VB.Label lblN1 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   99.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   2400
            TabIndex        =   64
            Top             =   2880
            Width           =   5295
         End
         Begin VB.Line Line1 
            BorderWidth     =   10
            X1              =   240
            X2              =   9840
            Y1              =   5640
            Y2              =   5640
         End
         Begin VB.Label lblSigno 
            Alignment       =   2  'Center
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   99.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2415
            Left            =   600
            TabIndex        =   63
            Top             =   2880
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Height          =   8535
         Left            =   -74880
         TabIndex        =   24
         Top             =   360
         Width           =   12135
         Begin VB.CommandButton cmdIniciar 
            Caption         =   "&Iniciar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   9840
            TabIndex        =   48
            Top             =   2640
            Width           =   2055
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3015
            Left            =   9840
            ScaleHeight     =   3015
            ScaleWidth      =   4215
            TabIndex        =   39
            Top             =   2640
            Visible         =   0   'False
            Width           =   4215
            Begin VB.ListBox lstFoundFiles4 
               Appearance      =   0  'Flat
               Height          =   2370
               Left            =   120
               TabIndex        =   40
               Top             =   480
               Width           =   1935
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "&Archivo Encontrado:"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   120
               Width           =   1815
            End
            Begin VB.Label Label5 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "0"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   1800
               TabIndex        =   41
               Top             =   120
               Width           =   1095
            End
         End
         Begin VB.TextBox txtTexto4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   80.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Top             =   240
            Width           =   11895
         End
         Begin VB.PictureBox Picture4 
            Height          =   3015
            Left            =   9960
            ScaleHeight     =   2955
            ScaleWidth      =   4275
            TabIndex        =   31
            Top             =   2760
            Visible         =   0   'False
            Width           =   4335
            Begin VB.ComboBox cmbLista4 
               Height          =   315
               IntegralHeight  =   0   'False
               ItemData        =   "frmLetrar.frx":0522
               Left            =   120
               List            =   "frmLetrar.frx":053E
               TabIndex        =   36
               Text            =   "Animales"
               Top             =   120
               Width           =   1935
            End
            Begin VB.TextBox txtSearchSpec4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2040
               TabIndex        =   35
               Text            =   "*.bmp;*.ico;*.wmf;*.rle;*.dib;*.jpg"
               Top             =   120
               Width           =   1575
            End
            Begin VB.FileListBox filList4 
               Appearance      =   0  'Flat
               Height          =   2370
               Left            =   120
               Pattern         =   "*.bmp;*.ico;*.wmf;*.rle;*.dib;*.jpg"
               TabIndex        =   34
               Top             =   480
               Width           =   1815
            End
            Begin VB.DirListBox dirList4 
               Appearance      =   0  'Flat
               Height          =   1665
               Left            =   2040
               TabIndex        =   33
               Top             =   960
               Width           =   1575
            End
            Begin VB.DriveListBox drvList4 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   2040
               TabIndex        =   32
               Top             =   600
               Width           =   2655
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Archivos"
               Height          =   375
               Left            =   120
               TabIndex        =   37
               Top             =   120
               Visible         =   0   'False
               Width           =   1815
            End
         End
         Begin VB.TextBox Text14 
            Height          =   375
            Left            =   8880
            TabIndex        =   30
            Top             =   5640
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdLimpiar4 
            Caption         =   "&Limpiar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   29
            Top             =   2640
            Width           =   1695
         End
         Begin VB.CommandButton cmdHablar4 
            Caption         =   "&Hablar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1800
            TabIndex        =   28
            Top             =   2640
            Width           =   1815
         End
         Begin VB.CommandButton cmdDetener4 
            Caption         =   "&Detener"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3600
            TabIndex        =   27
            Top             =   2640
            Width           =   2055
         End
         Begin VB.CommandButton cmdVer4 
            Caption         =   "&Ver"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5640
            TabIndex        =   26
            Top             =   2640
            Width           =   2055
         End
         Begin VB.CommandButton cmdAleatorio4 
            Caption         =   "&Aleatorio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   7680
            TabIndex        =   25
            Top             =   2640
            Width           =   2055
         End
         Begin AgentObjectsCtl.Agent Agent2 
            Left            =   14280
            Top             =   6960
            _cx             =   847
            _cy             =   847
         End
         Begin VB.Image ImgLetra4 
            BorderStyle     =   1  'Fixed Single
            Height          =   4935
            Left            =   120
            Stretch         =   -1  'True
            Top             =   3480
            Width           =   9615
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Tiempo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   9960
            TabIndex        =   44
            Top             =   6960
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label lblTiempo4 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1125
            Left            =   8760
            TabIndex        =   43
            Top             =   5760
            Visible         =   0   'False
            Width           =   1905
         End
      End
      Begin VB.Frame Frame3 
         Height          =   8535
         Left            =   -74880
         TabIndex        =   11
         Top             =   360
         Width           =   12135
         Begin VB.CheckBox chkAyuda 
            Caption         =   "Ayuda"
            Height          =   375
            Left            =   9480
            TabIndex        =   50
            Top             =   2400
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CommandButton cmdEmpezar2 
            Caption         =   "&Empezar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   9360
            TabIndex        =   21
            Top             =   2880
            Width           =   2415
         End
         Begin VB.TextBox txtTexto2 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   48.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   8895
         End
         Begin VB.Label lblLetra 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   99.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2475
            Index           =   5
            Left            =   9120
            TabIndex        =   22
            Top             =   4200
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label lblSeleccion2 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A E I O U"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   69.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   120
            TabIndex        =   20
            Top             =   2400
            Width           =   8895
         End
         Begin VB.Image ImgBarra2 
            BorderStyle     =   1  'Fixed Single
            Height          =   2055
            Left            =   9120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label lblLetra 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "U"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   99.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2475
            Index           =   4
            Left            =   7320
            TabIndex        =   19
            Top             =   4200
            Width           =   1695
         End
         Begin VB.Label lblLetra 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "O"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   99.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2475
            Index           =   3
            Left            =   5520
            TabIndex        =   18
            Top             =   4200
            Width           =   1695
         End
         Begin VB.Label lblLetra 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "I"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   99.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2475
            Index           =   2
            Left            =   3720
            TabIndex        =   17
            Top             =   4200
            Width           =   1695
         End
         Begin VB.Label lblLetra 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "E"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   99.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2475
            Index           =   1
            Left            =   1920
            TabIndex        =   16
            Top             =   4200
            Width           =   1695
         End
         Begin VB.Label lblLetra 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   99.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2475
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   4200
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Height          =   8535
         Left            =   -74880
         TabIndex        =   10
         Top             =   420
         Width           =   12135
         Begin VB.CheckBox chkNivel 
            Caption         =   "Ayuda"
            Height          =   375
            Left            =   9720
            TabIndex        =   52
            Top             =   2760
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CommandButton cmdEmpezar 
            Caption         =   "&Empezar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   9720
            TabIndex        =   13
            Top             =   3360
            Width           =   2295
         End
         Begin VB.Label lblResp2 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A E I O U"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   69.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   0
            TabIndex        =   51
            Top             =   6480
            Width           =   9735
         End
         Begin VB.Label lblSeleccion 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblSeleccion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   9720
            TabIndex        =   12
            Top             =   2160
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Image ImgRespuestas 
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Index           =   11
            Left            =   7200
            Stretch         =   -1  'True
            Top             =   4440
            Width           =   2415
         End
         Begin VB.Image ImgRespuestas 
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Index           =   10
            Left            =   4800
            Stretch         =   -1  'True
            Top             =   4440
            Width           =   2415
         End
         Begin VB.Image ImgRespuestas 
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Index           =   9
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   4440
            Width           =   2415
         End
         Begin VB.Image ImgBarra1 
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Left            =   9600
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2415
         End
         Begin VB.Image ImgRespuestas 
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Index           =   8
            Left            =   0
            Stretch         =   -1  'True
            Top             =   4440
            Width           =   2415
         End
         Begin VB.Image ImgRespuestas 
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Index           =   7
            Left            =   7200
            Stretch         =   -1  'True
            Top             =   2280
            Width           =   2415
         End
         Begin VB.Image ImgRespuestas 
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Index           =   6
            Left            =   4800
            Stretch         =   -1  'True
            Top             =   2280
            Width           =   2415
         End
         Begin VB.Image ImgRespuestas 
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Index           =   2
            Left            =   4800
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2415
         End
         Begin VB.Image ImgRespuestas 
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Index           =   3
            Left            =   7200
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2415
         End
         Begin VB.Image ImgRespuestas 
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Index           =   5
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   2280
            Width           =   2415
         End
         Begin VB.Image ImgRespuestas 
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Index           =   4
            Left            =   0
            Stretch         =   -1  'True
            Top             =   2280
            Width           =   2415
         End
         Begin VB.Image ImgRespuestas 
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Index           =   1
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2415
         End
         Begin VB.Image ImgRespuestas 
            BorderStyle     =   1  'Fixed Single
            Height          =   1935
            Index           =   0
            Left            =   0
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label lblQueEs 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   80.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   9600
            TabIndex        =   53
            Top             =   120
            Visible         =   0   'False
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Height          =   8415
         Left            =   -74880
         TabIndex        =   1
         Top             =   420
         Width           =   12135
         Begin VB.Frame Frame7 
            Caption         =   "Frame7"
            Height          =   3855
            Left            =   9840
            TabIndex        =   66
            Top             =   3480
            Width           =   5295
            Begin VB.TextBox Text1 
               Height          =   375
               Left            =   240
               TabIndex        =   78
               Top             =   3360
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.PictureBox Picture1 
               Height          =   3015
               Left            =   120
               ScaleHeight     =   2955
               ScaleWidth      =   4875
               TabIndex        =   71
               Top             =   480
               Width           =   4935
               Begin VB.DriveListBox drvList 
                  Appearance      =   0  'Flat
                  Height          =   315
                  Left            =   2040
                  TabIndex        =   76
                  Top             =   600
                  Width           =   2655
               End
               Begin VB.DirListBox dirList 
                  Appearance      =   0  'Flat
                  Height          =   1890
                  Left            =   2040
                  TabIndex        =   75
                  Top             =   960
                  Width           =   2655
               End
               Begin VB.FileListBox filList 
                  Appearance      =   0  'Flat
                  Height          =   2370
                  Left            =   120
                  Pattern         =   "*.bmp;*.ico;*.wmf;*.rle;*.dib;*.jpg"
                  TabIndex        =   74
                  Top             =   480
                  Width           =   1815
               End
               Begin VB.TextBox txtSearchSpec 
                  Appearance      =   0  'Flat
                  Height          =   285
                  Left            =   2040
                  TabIndex        =   73
                  Text            =   "*.bmp;*.ico;*.wmf;*.rle;*.dib;*.jpg"
                  Top             =   120
                  Width           =   1575
               End
               Begin VB.ComboBox cmbLista 
                  Height          =   315
                  IntegralHeight  =   0   'False
                  ItemData        =   "frmLetrar.frx":0599
                  Left            =   120
                  List            =   "frmLetrar.frx":05BB
                  TabIndex        =   72
                  Text            =   "Todos"
                  Top             =   120
                  Width           =   1935
               End
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Archivos"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   77
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1815
               End
            End
            Begin VB.PictureBox Picture2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   3015
               Left            =   0
               ScaleHeight     =   3015
               ScaleWidth      =   4575
               TabIndex        =   67
               Top             =   0
               Visible         =   0   'False
               Width           =   4575
               Begin VB.ListBox lstFoundFiles 
                  Appearance      =   0  'Flat
                  Height          =   2370
                  Left            =   120
                  TabIndex        =   68
                  Top             =   480
                  Width           =   1935
               End
               Begin VB.Label lblfound 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "&Archivo Encontrado:"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   120
                  TabIndex        =   70
                  Top             =   120
                  Width           =   1815
               End
               Begin VB.Label lblCount 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "0"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   1800
                  TabIndex        =   69
                  Top             =   120
                  Width           =   1095
               End
            End
         End
         Begin VB.CommandButton cmdIniciar 
            Caption         =   "&Iniciar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   2
            Left            =   9720
            TabIndex        =   49
            Top             =   2640
            Width           =   2055
         End
         Begin VB.TextBox txtTexto 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   80.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   240
            Width           =   11895
         End
         Begin VB.CommandButton cmdAleatorio 
            Caption         =   "&Aleatorio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   7680
            TabIndex        =   6
            Top             =   2640
            Width           =   2055
         End
         Begin VB.CommandButton cmdVer 
            Caption         =   "&Ver"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   5640
            TabIndex        =   5
            Top             =   2640
            Width           =   2055
         End
         Begin VB.CommandButton cmdDetener 
            Caption         =   "&Detener"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   3600
            TabIndex        =   4
            Top             =   2640
            Width           =   2055
         End
         Begin VB.CommandButton cmdHablar 
            Caption         =   "&Hablar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1800
            TabIndex        =   3
            Top             =   2640
            Width           =   1815
         End
         Begin VB.CommandButton cmdLimpiar 
            Caption         =   "&Limpiar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   2
            Top             =   2640
            Width           =   1695
         End
         Begin AgentObjectsCtl.Agent Agent1 
            Left            =   12360
            Top             =   5880
            _cx             =   847
            _cy             =   847
         End
         Begin VB.Image ImgLetra 
            BorderStyle     =   1  'Fixed Single
            Height          =   4815
            Left            =   120
            Stretch         =   -1  'True
            Top             =   3480
            Width           =   9615
         End
         Begin VB.Label lblTiempo2 
            Alignment       =   2  'Center
            Caption         =   "Tiempo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   10200
            TabIndex        =   9
            Top             =   6480
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label lblTiempo 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1125
            Left            =   9840
            TabIndex        =   8
            Top             =   6360
            Width           =   1905
         End
      End
      Begin VB.Frame Frame4 
         Height          =   8535
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   12135
         Begin VB.CommandButton cmdIniciar 
            Caption         =   "&Iniciar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   0
            Left            =   7560
            TabIndex        =   46
            Top             =   7560
            Width           =   2055
         End
         Begin VB.Label lblLetra5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   69.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   240
            TabIndex        =   47
            Top             =   5880
            Width           =   11655
         End
         Begin VB.Image ImgLetra5 
            BorderStyle     =   1  'Fixed Single
            Height          =   7335
            Left            =   120
            Stretch         =   -1  'True
            Top             =   120
            Width           =   11895
         End
      End
   End
End
Attribute VB_Name = "frmLetrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AnimeLextor As IAgentCtlCharacterEx
Dim Tiempo_Activo, Segundos, NumRegistros, IniciarImg As Integer
Const DATAPATH = "E-man.acs" 'genie.acs"
Dim SearchFlag As Integer    ' Used as flag for cancelling, etc.
Dim xDirectorioActual As String
Private Type ListaArchivos
    Nombre As String
    Letra As String
    Archivo As String
End Type
Dim XLetra() As ListaArchivos

Private Sub chkNivel_Click()
    If ImgBarra1.Visible = True Then
        ImgBarra1.Visible = False
        lblQueEs.Visible = True
        chkAyuda.Value = 0
    Else
       ImgBarra1.Visible = True
        lblQueEs.Visible = False
       chkAyuda.Value = 1
    End If
End Sub

Private Sub Form_Load()
    Randomize
    IniciarImg = 0
    Segundos = 10
    Segundos3 = 0
    lblTiempo = Segundos
    Tiempo_Activo = True
    xDirectorioActual = dirList.Path
    Agent1.Characters.Load "E-man", DATAPATH
    Set AnimeLextor = Agent1.Characters("E-man")
'    AnimeLextor.Show
     AnimeLextor.Left = 600
     AnimeLextor.Top = 450
     Hablar_Leer ("Dios los Bendiga Les Saluda E-Man")
        Call DirList_LostFocus
        Call Arreglo
    lblTiempo.Visible = False
    Call cmbLista_Click
    'AnimeLextor.LanguageID = &H409
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub cmbLista_Click()
    If Me.cmbLista.Text <> "Todos" Then
        Picture1.Visible = True
        Picture2.Visible = False
        Call DirList_LostFocus
        Call Arreglo
        'Call DirList_Change
    Else
      ' Continue with the search.
        Picture1.Visible = False
        Picture2.Visible = True
        Call cmdSearch_Click
    End If

End Sub

Private Sub cmdSearch_Click()
' Initialize for search, then call DirDiver to perform recursive search.
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim result As Integer
  ' Check what the user did last:
  'If cmdSearch.Caption = "&Resetear" Then  ' If just a reset,
  '  ResetSearch                         ' initialize and exit.
  '  txtSearchSpec.SetFocus
  '  Exit Sub
  'End If

  ' Update dirList.Path if it is different from the currently
  ' selected directory, otherwise perform the search.
  If dirList.Path <> dirList.List(dirList.ListIndex) Then
     dirList.Path = dirList.List(dirList.ListIndex)
     Exit Sub         ' Exit so user can take a look before searching.
  End If

  dirList.Path = xDirectorioActual
  'cmdExit.Caption = "&Cancelar"

  filList.Pattern = txtSearchSpec.Text
  FirstPath = dirList.Path
  DirCount = dirList.ListCount

  'Start recursive direcory search.
  NumFiles = 0                       ' Reset global foundfiles indicator.
  result = DirDiver(FirstPath, DirCount, "")
  filList.Path = dirList.Path
  'cmdSearch.Caption = "&Resetear"
  'cmdSearch.SetFocus
  'cmdExit.Caption = "&Salir"
   Call Arreglo2
End Sub

Sub Arreglo2()
    Dim NumPosiciones As Byte
    Dim P, I As Integer
    Dim xTexto As String
    If (lstFoundFiles.ListCount - 1) = -1 Then Exit Sub
    NumRegistros = lstFoundFiles.ListCount
    ReDim XLetra(NumRegistros)
    For I = 0 To lstFoundFiles.ListCount - 1
        For P = 0 To Len(lstFoundFiles.List(I))
            xTexto = Mid(lstFoundFiles.List(I), 1, Len(lstFoundFiles.List(I)) - P)
            If Right(Mid(lstFoundFiles.List(I), 1, Len(lstFoundFiles.List(I)) - P), 1) = "\" Then
               P = Len(xTexto)
               Exit For
            End If
        Next
        XLetra(I).Nombre = Mid(lstFoundFiles.List(I), (Len(xTexto) + 1), Len(lstFoundFiles.List(I)))
        XLetra(I).Nombre = Mid(XLetra(I).Nombre, 1, (Len(XLetra(I).Nombre) - 4))
        XLetra(I).Letra = Mid(XLetra(I).Nombre, 1, 1)
        If Right(filList.Path, 1) <> "\" Then
            XLetra(I).Archivo = lstFoundFiles.List(I)
        Else
            XLetra(I).Archivo = lstFoundFiles.List(I)
        End If
    Next
End Sub

Private Sub cmbLista_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdAleatorio_Click()
    On Error GoTo FueraRango
    Dim P As Integer
    Segundos3 = 0
    If NumRegistros <> 0 Then
       'Seleccionamos aletatoriamente la palabra
        P = Int((Rnd * NumRegistros))
        txtTexto.Text = XLetra(P).Letra 'XLetra(P).Nombre
        Select Case XLetra(P).Letra
        Case "V"
            Hablar_Leer (XLetra(P).Letra & "e de " & XLetra(P).Nombre)
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
            Hablar_Leer XLetra(P).Letra
        Case Else
            Hablar_Leer (XLetra(P).Letra & " de " & XLetra(P).Nombre)
        End Select
'        If XLetra(P).Letra = "V" Then
'            Hablar_Leer (XLetra(P).Letra & "e de " & XLetra(P).Nombre)
'        Else
'            Hablar_Leer (XLetra(P).Letra & " de " & XLetra(P).Nombre)
'        End If
        ImgLetra.Picture = LoadPicture(XLetra(P).Archivo)
        Text14.Text = XLetra(P).Nombre
    End If
    Exit Sub
FueraRango:
    NumRegistros = 0
    Call Arreglo
End Sub

Private Sub cmdSalir_Click(Index As Integer)
  'If cmdSalir.Caption = "&Salir" Then
    End
  'Else                ' If Cancel, just end Search.
  '  SearchFlag = False
  'End If
End Sub

Private Sub cmdDetener_Click()
     AnimeLextor.Stop
     txtTexto.SetFocus
End Sub

Private Sub cmdHablar_Click()
    Hablar_Leer (txtTexto.Text)
    txtTexto.SetFocus
    txtTexto.SelStart = 0
    txtTexto.SelLength = Len(txtTexto.Text)
    'AnimeLextor.Hide
End Sub

Sub Hablar_Leer(Texto As String)
    AnimeLextor.Show
    If Not Texto = "" Then
        AnimeLextor.Speak Texto
    End If
End Sub

Private Sub cmdLimpiar_Click()
    txtTexto.Text = ""
    txtTexto.SetFocus
End Sub

Private Sub Frame7_DblClick()
    If Frame7.Top <> 3720 Then
        Frame7.Top = 3720
        Frame7.Left = 2760
    Else
        Frame7.Top = 3480
        Frame7.Left = 9840
    End If
End Sub

Private Sub lblTiempo_Click()
    If lblTiempo.Visible = False Then
        lblTiempo.Visible = True
    Else
        lblTiempo.Visible = False
    End If
    txtTexto.SetFocus
End Sub

Private Sub lblTiempo2_Click()
    If lblTiempo.Visible = False Then
        lblTiempo.Visible = True
    Else
        lblTiempo.Visible = False
    End If
End Sub

Private Sub Picture2_Click()
    Me.Picture1.Visible = True
    Picture2.Visible = False
End Sub

Private Sub Picture5_Click()
    Picture4.Visible = True
End Sub

Private Sub timTiempo_Timer()
    On Error GoTo FueraRango
    If Not lblTiempo.Visible = False Then
        Segundos = Val(lblTiempo)
        If Tiempo_Activo = True Then
            Segundos = Segundos - 1
        Else
            Segundos = Segundos + 1
        End If
        If Segundos = -1 Then
            Segundos = 0
            Tiempo_Activo = False
            ImgLetra.Picture = LoadPicture("")
            'AnimeLextor.Hide
        End If
        If Segundos = 11 Then
            Segundos = 10
            Tiempo_Activo = True
            ImgLetra.Picture = LoadPicture("")
            'AnimeLextor.Hide
        End If
        lblTiempo = Segundos
    End If
    Segundos3 = Segundos3 + 1
    If Segundos3 >= 40 Then
       Segundos3 = 0
       'AnimeLextor.Hide
       AnimeLextor.Stop
    End If
    If cmdIniciar(0).Caption <> "&Iniciar" Then
       IniciarImg = IniciarImg + 1
       If IniciarImg >= 5 Then
          Dim P As Integer
          IniciarImg = 0
          If NumRegistros <> 0 Then
             'Seleccionamos aletatoriamente la palabra
             P = Int((Rnd * NumRegistros))
             If Len(XLetra(P).Nombre) >= 3 Then
                lblLetra5.Caption = XLetra(P).Nombre  'XLetra(P).Letra
             Else
                lblLetra5.Caption = ""
             End If
             Select Case XLetra(P).Letra
             Case "V"
'                 Hablar_Leer (XLetra(P).Letra & "e de " & XLetra(P).Nombre)
                  Hablar_Leer XLetra(P).Nombre
             Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
'                 Hablar_Leer XLetra(P).Letra
                  Hablar_Leer XLetra(P).Nombre
             Case Else
                  Hablar_Leer XLetra(P).Nombre
             End Select
                txtTexto4.Text = XLetra(P).Nombre   'XLetra(P).Letra
                txtTexto.Text = XLetra(P).Nombre   'XLetra(P).Letra
                ImgLetra5.Picture = LoadPicture(XLetra(P).Archivo)
                ImgLetra4.Picture = LoadPicture(XLetra(P).Archivo)
                ImgLetra.Picture = LoadPicture(XLetra(P).Archivo)
                SeekView.Image1.Picture = ImgLetra5.Picture
          End If
       End If
    End If
    Exit Sub
FueraRango:
    NumRegistros = 0
    Call Arreglo
End Sub

Private Sub txtTexto_KeyPress(KeyAscii As Integer)
    Segundos3 = 0
    On Error Resume Next
    Call DirList_Change
    FLetra (KeyAscii)
End Sub

Sub FLetra(KeyAscii As Integer)
   Dim P, Tiempo As Integer
   'Seleccionamos aletatoriamente la palabra
   Tiempo = 0
   If NumRegistros = 0 Then
        AnimeLextor.Speak Chr(KeyAscii)
        ImgLetra.Picture = LoadPicture("")
        Text1.Text = ""
        Exit Sub
   End If
   
   Do While Tiempo <> 10000
      P = Int((Rnd * NumRegistros))
      If XLetra(P).Letra = UCase(Chr(KeyAscii)) Then Exit Do
      Tiempo = Tiempo + 1
   Loop
   If Tiempo <> 10000 Then
      Select Case XLetra(P).Letra
      Case "V"
          Hablar_Leer (XLetra(P).Letra & "e de " & XLetra(P).Nombre)
      Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
          Hablar_Leer XLetra(P).Letra
      Case Else
          Hablar_Leer (XLetra(P).Letra & " de " & XLetra(P).Nombre)
      End Select
      
      If NumRegistros <> 0 Then
         ImgLetra.Picture = LoadPicture(XLetra(P).Archivo)
         Text1.Text = XLetra(P).Archivo
      End If
      Text14.Text = XLetra(P).Nombre
      Call Tiempo_Act
      Exit Sub
    Else
        AnimeLextor.Speak Chr(KeyAscii)
        ImgLetra.Picture = LoadPicture("")
        Text1.Text = ""
    End If
End Sub

Sub Tiempo_Act()
    If Tiempo_Activo = True Then
       lblTiempo = 10
    Else
       lblTiempo = 0
    End If
End Sub

Private Sub cmdVer_Click()
    txtTexto.SetFocus
    If Not ImgLetra.Picture = 0 Then
       If Text14.Text <> "" Then
          txtTexto.Text = Text14.Text
          Call cmdHablar_Click
       End If
       SeekView.Show
       SeekView.Image1.Picture = ImgLetra.Picture
       Load SeekView
    End If
End Sub

Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer
'  Recursively search directories from NewPath down...
'     NewPath is searched on this recursion.
'     BackUp is origin of this recursion.
'     DirCount is number of subdirectories in this directory.
Static FirstErr As Integer
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String
Dim retval As Integer
  SearchFlag = True             ' Set flag so user can interrupt.
  DirDiver = False              ' Set to TRUE if there is an error.
  retval = DoEvents()           ' check for events (i.e. user Cancels).
  If SearchFlag = False Then
    DirDiver = True
    Exit Function
  End If
  On Local Error GoTo DirDriverHandler
  DirsToPeek = dirList.ListCount            ' How many directories below this?
  Do While DirsToPeek > 0 And SearchFlag = True
    OldPath = dirList.Path                  ' Save old path for next recursion.
    dirList.Path = NewPath
    If dirList.ListCount > 0 Then
    ' Get to the node bottom.
      dirList.Path = dirList.List(DirsToPeek - 1)
      AbandonSearch = DirDiver((dirList.Path), DirCount%, OldPath)
    End If
    ' Go up 1 level in directories.
    DirsToPeek = DirsToPeek - 1
    If AbandonSearch = True Then Exit Function
  Loop
  ' Call function to enumerate files.
  If filList.ListCount Then
    If Len(dirList.Path) <= 3 Then
        ThePath = dirList.Path         ' If at root level, leave as is...
    Else
        ThePath = dirList.Path + "\"   ' otherwise put "\" before file name.
    End If
    For ind = 0 To filList.ListCount - 1        ' Add conforming files in
        entry = ThePath + filList.List(ind)     ' this directory to listbox.
        lstFoundFiles.AddItem entry
        lblCount.Caption = Str$(Val(lblCount.Caption) + 1)
    Next ind
  End If
  If BackUp <> "" Then         ' If there is a superior
      dirList.Path = BackUp    ' directory, move to it.
  End If
  Exit Function
DirDriverHandler:
  If Err = 7 Then         ' If Out of Memory, assume listbox just got full.
    DirDiver = True       ' Create Msg$ and set return value AbandonSearch.
    MsgBox "You've filled the listbox. Search being abandoned..."
    Exit Function         ' Note that EXIT procedure resets ERR to 0.
  Else                    ' Otherwise display error message and quit.
    MsgBox Error
    End
  End If
End Function

Private Sub DirList_Change()
    ' Update File listbox to sync with Dir listbox.
    filList.Path = dirList.Path
    filList.Pattern = txtSearchSpec.Text
    If txtTexto.Text <> "" Then txtTexto.SetFocus
End Sub

Private Sub DirList_LostFocus()
    Dim xDirectorio As String
    Dim ind As Integer
    If Me.cmbLista.Text <> "Todos" Then
       xDirectorio = cmbLista.Text
       filList.Path = xDirectorioActual 'dirList.Path
       dirList.Path = filList.Path & "\" & xDirectorio
       If txtTexto.Text <> "" Then txtTexto.SetFocus
    Else
        ReDim XLetra(filList.ListCount)
       For ind = 0 To cmbLista.ListCount - 1        ' Add conforming files in
           xDirectorio = cmbLista.List(ind)     ' this directory to listbox.
           filList.Path = xDirectorioActual 'dirList.Path
           If xDirectorio <> "Todos" Then
           dirList.Path = filList.Path & "\" & xDirectorio
           Call Arreglo
           End If
       Next ind
       xDirectorio = "Animales"
       If txtTexto.Text <> "" Then txtTexto.SetFocus
    
    End If

End Sub

Private Sub DrvList_Change()
    On Error GoTo DriveHandler
    dirList.Path = drvList.Drive
    txtTexto.SetFocus
    Exit Sub

DriveHandler:
    drvList.Drive = dirList.Path
    Exit Sub
End Sub

Private Sub filList_Click()
    On Error GoTo Error
      If Right(filList.Path, 1) <> "\" Then
        Text1.Text = filList.Path & "\" & filList.FileName
      Else
        Text1.Text = filList.Path & filList.FileName
      End If
    Dim dropfile As String, X, nl, msg
    dropfile = Text1.Text
    Select Case Right(LCase(Text1.Text), 3)
    Case "txt"
        Me.Label1.Caption = "Archivo de Texto"
    Case "bmp", "wmf", "rle", "ico", "dib", "jpg"
        Me.ImgLetra.Picture = LoadPicture(Text1.Text)
    Case "exe"
        Me.Label1.Caption = "Archivo Ejecutable"
    Case "hlp"
        Me.Label1.Caption = "Archivo de Ayuda"
    Case Else
        nl = Chr$(10) + Chr$(13)
        msg = "Es un Tipo de Archivo Desconocido:"
        msg = nl + msg + nl + nl + "     .ico, .bmp, .wmf, .rle, .jpg"
        MsgBox msg
    End Select
    Call Tiempo_Act
    txtTexto.Text = Mid(filList.FileName, 1, Len(filList.FileName) - 4)
    Text14.Text = Mid(filList.FileName, 1, Len(filList.FileName) - 4)
    Call cmdHablar_Click
    txtTexto.SetFocus
    Exit Sub
    
Error:
    txtTexto.SetFocus
End Sub

Private Sub Image1_DblClick()
    Dim dropfile As String, X, nl, msg
    dropfile = Text1.Text
    Select Case Right(LCase(Text1.Text), 3)
    Case "txt"
        X = Shell("Notepad " + dropfile, 1)
    Case "bmp", "dib"
        X = Shell("pbrush " + dropfile, 1)
    Case "ico"
        X = Shell("ICONWRKS " + dropfile, 1)
    Case "bmp", "wmf", "rle", "ico", "dib", "jpg"
        Me.ImgLetra.Picture = LoadPicture(Text1.Text)
'        X = Shell("pbrush " + dropfile, 1)
    Case "exe"
        X = Shell(dropfile, 1)
    Case "hlp"
        X = Shell("WinHelp " + dropfile, 1)
    Case Else
        nl = Chr$(10) + Chr$(13)
        msg = "Es un Tipo de Archivo Desconosido:"
        msg = nl + msg + nl + nl + "     .ico, .bmp, .wmf, .rle, .jpg"
        MsgBox msg
    End Select
End Sub

Private Sub ImgLetra_DblClick()
    Dim dropfile As String, X, nl, msg
    dropfile = Text1.Text
    Select Case Right(LCase(Text1.Text), 3)
    Case "txt"
        X = Shell("Notepad " + dropfile, 1)
    Case "bmp", "wmf", "rle", "ico", "dib", "jpg"
        Me.ImgLetra.Picture = LoadPicture(Text1.Text)
    Case "exe"
        X = Shell(dropfile, 1)
    Case "hlp"
        X = Shell("WinHelp " + dropfile, 1)
    Case Else
        If ImgLetra.Picture = 0 Then
           nl = Chr$(10) + Chr$(13)
           msg = "Es un Tipo de Archivo Desconosido:"
           msg = nl + msg + nl + nl + "     .ico, .bmp, .wmf, .rle, .jpg"
           MsgBox msg
             Exit Sub
        End If
    End Select
    If Text14.Text <> "" Then
       txtTexto.Text = Text14.Text
       Call cmdHablar_Click
    End If
    SeekView.Show
    If Picture2.Visible = False Then
       SeekView.Image1.Picture = ImgLetra.Picture
    Else
       SeekView.Image1.Picture = ImgLetra.Picture
    End If
    Load SeekView
End Sub

Private Sub lstFoundFiles_Click()
    Text1.Text = lstFoundFiles
    Dim dropfile As String, X, nl, msg
    dropfile = Text1.Text
    Select Case Right(LCase(Text1.Text), 3)
    Case "txt"
        Me.Label1.Caption = "Archivo de Texto"
    Case "bmp", "wmf", "rle", "ico", "dib", "jpg"
        Me.ImgLetra.Picture = LoadPicture(Text1.Text)
    Case "exe"
        Me.Label1.Caption = "Archivo Ejecutable"
    Case "hlp"
        Me.Label1.Caption = "Archivo de Ayuda"
    Case Else
        nl = Chr$(10) + Chr$(13)
        msg = "Es un Tipo de Archivo Desconocido:"
        msg = nl + msg + nl + nl + "     .ico, .bmp, .wmf, .rle , .jpg"
        MsgBox msg
    End Select
    X = lstFoundFiles.ListIndex
    txtTexto.Text = XLetra(X).Nombre
    Call cmdHablar_Click
End Sub

Private Sub ResetSearch()
' Reinitialize before starting a new search.
'    lstFoundFiles.Clear
'    lblCount.Caption = 0
'    SearchFlag = False                  ' Flag indicating search in progress.
'    Picture2.Visible = False
'    cmdSearch.Caption = "&Buscar"
'    cmdExit.Caption = "&Salir"
'    Picture1.Visible = True
'   dirList.Path = CurDir$: drvList.Drive = dirList.Path ' Reset DOS path.
End Sub

Private Sub txtSearchSpec_Change()
' Update file list box if user changes pattern.
    filList.Pattern = txtSearchSpec.Text
End Sub

Private Sub txtSearchSpec_GotFocus()
    txtSearchSpec.SelStart = 0      ' Highlight the current entry.
    txtSearchSpec.SelLength = Len(txtSearchSpec.Text)
End Sub

Sub Arreglo()
    Dim NumPosiciones As Byte
    Dim P, I As Integer
    If (filList.ListCount - 1) = -1 Then Exit Sub
    NumRegistros = filList.ListCount
    ReDim XLetra(NumRegistros)
    For I = 0 To filList.ListCount - 1
        XLetra(I).Nombre = Mid(filList.List(I), 1, Len(filList.List(I)) - 4)
        XLetra(I).Letra = Mid(filList.List(I), 1, 1)
        If Right(filList.Path, 1) <> "\" Then
            XLetra(I).Archivo = filList.Path & "\" & filList.List(I)
        Else
            XLetra(I).Archivo = filList.Path & "\" & filList.List(I)
        End If
    Next
End Sub

'*********************************************************************************
' Adivinar
'*********************************************************************************
Private Sub cmdEmpezar_Click()
        Call Cargar
        Call Inicio
End Sub

Sub Cargar()
    Dim I, Niveles As Integer
    Dim P As Integer
    NumRegistros = 0
    Niveles = 1
    If lstFoundFiles.ListCount > 10 And Me.Picture1.Visible = False Then
       NumRegistros = lstFoundFiles.ListCount
    Else
       NumRegistros = filList.ListCount
    End If
    'ReDim XLetra(filList.ListCount)
    ReDim MILLONARIO(NumRegistros)
    P = Int((Rnd * NumRegistros))
    MILLONARIO(P).Preguntas = XLetra(P).Nombre
    MILLONARIO(P).Respuestas = XLetra(P).Archivo
    ImgBarra1.Picture = LoadPicture(XLetra(P).Archivo)
    If NumRegistros < 12 Then
       NumRegistros = 12
    End If
End Sub
        
Private Sub Inicio()
    On Error GoTo Error
    Dim P, P2 As Integer
    Dim I As Integer
    Dim X, J As Integer
    Dim NumPosiciones As Byte
    Dim PosLetra As Byte
    Dim PosUsadas() As Integer
    Dim PosLibre As Boolean
    Dim Categoria As String

    'Seleccionamos aletatoriamente la palabra
    P = Int((Rnd * NumRegistros))
    'Configuración del entorno para un nuevo turno
    MILLONARIO(P).Preguntas = XLetra(P).Nombre
    MILLONARIO(P).Respuestas = XLetra(P).Archivo
    ImgBarra1.Picture = LoadPicture(XLetra(P).Archivo)
    lblSeleccion = MILLONARIO(P).Respuestas
    ImgBarra1.Picture = LoadPicture(XLetra(P).Archivo)
    NumPosiciones = 12 'NumRegistros

    'Colorear todas las Respuestas en Color Negro
'    For I = 0 To NumPosiciones
'        ImgRespuestas(I).Picture = LoadPicture(XLetra(I).Archivo)
'    Next I
    '---------------------------------------------
    '| Asignamos al array auxiliar la dimensión del
    '| número de letras y lo inicializamos a -1
    '---------------------------------------------
    ReDim PosUsadas(NumPosiciones)
    For I = 0 To NumPosiciones
        PosUsadas(I) = -1
    Next I
    'Colocamos las letras desordenadas
    I = 0
    J = P
    Do
        PosLetra = Int(Rnd * NumPosiciones)
        PosLibre = True
        For X = 0 To NumPosiciones
            If PosUsadas(X) = PosLetra Then
                PosLibre = False
            End If
        Next X
        If PosLibre Then
'            If lblSeleccion = MILLONARIO(J).Respuestas Then
                P2 = Int((Rnd * NumRegistros))
                PosUsadas(I) = PosLetra
                ImgRespuestas(PosLetra).Picture = LoadPicture(XLetra(P2).Archivo)
                MILLONARIO(PosLetra).Preguntas = XLetra(P2).Nombre
                MILLONARIO(PosLetra).Respuestas = XLetra(P2).Archivo
                I = I + 1
'            End If
            J = J + 1
            If J > NumRegistros Then J = 0
        End If
    Loop While I < NumPosiciones
    'Colocamos las Figura en la Posicion desordenadas
    PosLetra = Int(Rnd * NumPosiciones)
    PosUsadas(PosLetra) = PosLetra
    ImgRespuestas(PosLetra).Picture = LoadPicture(XLetra(P).Archivo)
    MILLONARIO(PosLetra).Preguntas = XLetra(P).Nombre
    MILLONARIO(PosLetra).Respuestas = XLetra(P).Archivo
    If Len(MILLONARIO(PosLetra).Preguntas) > 15 Then
       lblResp2.FontSize = 40
    Else
       lblResp2.FontSize = 70
    End If
    lblResp2 = MILLONARIO(PosLetra).Preguntas
    Hablar_Leer lblResp2
    Exit Sub
Error:

End Sub

Private Sub ImgRespuestas_Click(Index As Integer)
    On Error GoTo ErrorX
    Dim X As Integer
    'If ImgRespuestas(Index).Picture = ImgBarra1.Picture Then
    If lblSeleccion = MILLONARIO(Index).Respuestas Then
        Hablar_Leer "Muy bien acertó la Imagen es un " & MILLONARIO(Index).Preguntas
        'MsgBox "Muy bien acertó la Imagen", , "Muy Bien"
            SeekView.Show
       SeekView.Image1.Picture = ImgRespuestas(Index).Picture
    Load SeekView
        Segundos3 = 0
        Call Cargar
        Call Inicio
    Else
        Hablar_Leer MILLONARIO(Index).Preguntas
        If MILLONARIO(Index).Preguntas = lblResp2 Then
            Hablar_Leer "Muy bien acertó la Imagen es un " & MILLONARIO(Index).Preguntas
            'MsgBox "Muy bien acertó la Imagen", , "Muy Bien"
            SeekView.Show
            SeekView.Image1.Picture = ImgRespuestas(Index).Picture
            Load SeekView
            Segundos3 = 0
            Call Cargar
            Call Inicio
        End If
    End If
    Exit Sub
ErrorX:
End Sub

Private Sub lblQueEs_Click()
    Call ImgBarra1_Click
End Sub

Private Sub ImgBarra1_Click()
    Hablar_Leer lblResp2
End Sub

'-------------------------------------------------------------------------------------------
' Juego
'-------------------------------------------------------------------------------------------
Private Sub ImgBarra2_Click()
    Call chkAyuda_Click
    Hablar_Leer lblSeleccion2
End Sub

Private Sub chkAyuda_Click()
  If lblSeleccion2.Visible = False Then
     chkAyuda.Value = 1
     lblSeleccion2.Visible = True
  Else
     chkAyuda.Value = 0
     lblSeleccion2.Visible = False
  End If
  cmdEmpezar2.SetFocus
End Sub

Private Sub cmdEmpezar2_Click()
    Dim P As Integer
    Call Cargar2(P)
    Call Inicio2(P)
End Sub

Private Sub cmdEmpezar2_KeyPress(KeyAscii As Integer)
    Dim X As Integer
    If UCase(Chr(KeyAscii)) = lblLetra(5).Caption Then
        txtTexto2.Text = UCase(lblSeleccion2)
        Hablar_Leer "Muy Bien La Letra es " & lblLetra(5).Caption & " Y dece : " & txtTexto2.Text
'        MsgBox "Muy Bien", , "Muy Bien"
            SeekView.Show
            SeekView.Image1.Picture = ImgBarra2.Picture
            Load SeekView

        Call cmdEmpezar2_Click
    Else
        Hablar_Leer "No es la Letra "
    End If
End Sub

Sub Cargar2(P As Integer)
    On Error GoTo NuevoNumero
    Dim I, Niveles As Integer
    P = 0
    NumRegistros = 0
    Niveles = 1
    NumRegistros = filList.ListCount
    If NumRegistros = 2 Then
       NumRegistros = lstFoundFiles.ListCount
    End If
    If NumRegistros < 12 Then
       NumRegistros = 12
    End If
    ReDim ALEATORIO(NumRegistros)
    P = Int((Rnd * NumRegistros))
NumeroNew:
    ALEATORIO(P).Preguntas = XLetra(P).Nombre
    ALEATORIO(P).Respuestas = XLetra(P).Archivo
    ImgBarra2.Picture = LoadPicture(XLetra(P).Archivo)
    lblSeleccion2 = ALEATORIO(P).Respuestas
    txtTexto2.Text = UCase(XLetra(P).Nombre)
    If Len(txtTexto2.Text) > 9 Then
       txtTexto2.FontSize = 40
    Else
       txtTexto2.FontSize = 80
    End If
    If Len(txtTexto2.Text) <= 2 Then ' La Letra es muy pequeño
       Call Cargar2(P)
    Else
    End If
    Exit Sub
NuevoNumero:
    P = 1
    GoTo NumeroNew
End Sub
        
Private Sub Inicio2(P As Integer)
    
    Dim I As Integer
    Dim X, J As Integer
    Dim NumPosiciones As Byte
    Dim PosLetra As Byte
    Dim PosUsadas() As Integer
    Dim PosLibre As Boolean
    Dim Categoria As String
    NumRegistros = Len(txtTexto2.Text)
    'Configuración del entorno para un nuevo turno
    ALEATORIO(P).Preguntas = XLetra(P).Nombre
    ALEATORIO(P).Respuestas = Mid(txtTexto2.Text, 1, P)
    ImgBarra1.Picture = LoadPicture(XLetra(P).Archivo)
    lblSeleccion2 = XLetra(P).Nombre
    'ImgBarra1.Picture = LoadPicture(XLetra(P).Archivo)
    NumPosiciones = NumRegistros

    '---------------------------------------------
    '| Asignamos al array auxiliar la dimensión del
    '| número de letras y lo inicializamos a -1
    '---------------------------------------------
    NumPosiciones = 5
    ReDim PosUsadas(NumPosiciones)
    For I = 0 To NumPosiciones
        PosUsadas(I) = -1
    Next I
    'Colocamos las letras desordenadas
    I = 0
    J = P
    Do
        PosLetra = Int(Rnd * NumPosiciones)
        PosLibre = True
        For X = 0 To NumPosiciones
            If PosUsadas(X) = PosLetra Then
                PosLibre = False
            End If
        Next X
        If PosLibre Then
'            If lblSeleccion = MILLONARIO(J).Respuestas Then
                PosUsadas(I) = PosLetra
                ImgRespuestas(PosLetra).Picture = LoadPicture(XLetra(I).Archivo)
                lblLetra(PosLetra).Caption = UCase(Mid(txtTexto2.Text, I + 1, 1))
                ALEATORIO(PosLetra).Preguntas = XLetra(I).Nombre
                ALEATORIO(PosLetra).Respuestas = XLetra(I).Archivo
                I = I + 1
'            End If
            J = J + 1
            If J > NumRegistros Then J = 0
        End If
    Loop While I < NumPosiciones
    'Colocamos las Figura en la Posicion desordenadas
    PosLetra = Int(Rnd * NumPosiciones)
    PosUsadas(PosLetra) = PosLetra
    lblLetra(5).Caption = UCase(Mid(txtTexto2.Text, PosLetra + 1, 1))
    ALEATORIO(P).Respuestas = txtTexto2.Text
    txtTexto2.Text = ""
    If Len(lblLetra(5).Caption) = 0 Then
       lblLetra(5).Caption = lblLetra(2).Caption
       
    End If
    For X = 0 To Len(XLetra(P).Nombre)
        If lblLetra(5).Caption = UCase(Mid(XLetra(P).Nombre, X + 1, 1)) Then
           txtTexto2.Text = txtTexto2.Text + "_"
        Else
           txtTexto2.Text = txtTexto2.Text + UCase(Mid(XLetra(P).Nombre, X + 1, 1))
        End If
    Next
End Sub

Private Sub lblLetra_Click(Index As Integer)
    Dim X, XX As Integer
    If lblLetra(Index).Caption = lblLetra(5).Caption Then
        txtTexto2.Text = UCase(lblSeleccion2)
        Hablar_Leer "Muy Bien La Letra es " & lblLetra(5).Caption & " Y dece : " & txtTexto2.Text
            SeekView.Show
            SeekView.Image1.Picture = ImgBarra2.Picture
            Load SeekView
        'MsgBox "Muy Bien", , "Muy Bien"
        Call cmdEmpezar2_Click
    Else
        Hablar_Leer "No es la Letra "
    End If
End Sub

'-------------------------------------------------------------------------------------------
' Videos
'-------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Palabras
'-------------------------------------------------------------------------------------------------
Private Sub txtTexto4_KeyPress(KeyAscii As Integer)
    Segundos3 = 0
    On Error Resume Next
    Call DirList4_Change
    FLetra4 (KeyAscii)
End Sub

Sub FLetra4(KeyAscii As Integer)
   Dim P, Tiempo As Integer
   'Seleccionamos aletatoriamente la palabra
   Tiempo = 0
   If NumRegistros = 0 Then
        AnimeLextor.Speak Chr(KeyAscii)
        ImgLetra4.Picture = LoadPicture("")
        Exit Sub
   End If
   
   Do While Tiempo <> 10000
      P = Int((Rnd * NumRegistros))
      If XLetra(P).Letra = UCase(Chr(KeyAscii)) Then Exit Do
      Tiempo = Tiempo + 1
   Loop
   If Tiempo <> 10000 Then
      Select Case XLetra(P).Letra
      Case "V"
'          Hablar_Leer (XLetra(P).Letra & "e de " & XLetra(P).Nombre)
          Hablar_Leer XLetra(P).Nombre
      Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
'          Hablar_Leer XLetra(P).Letra
          Hablar_Leer XLetra(P).Nombre
      Case Else
'          Hablar_Leer (XLetra(P).Letra & " de " & XLetra(P).Nombre)
          Hablar_Leer XLetra(P).Nombre
      End Select
      txtTexto4.Text = ""
      txtTexto4.Text = Mid(XLetra(P).Nombre, 2, Len(XLetra(P).Nombre))
      If NumRegistros <> 0 Then
         ImgLetra4.Picture = LoadPicture(XLetra(P).Archivo)
      End If
      Call Tiempo_Act
      Exit Sub
    Else
        AnimeLextor.Speak Chr(KeyAscii)
        ImgLetra4.Picture = LoadPicture("")
      txtTexto4.Text = ""
    End If
End Sub

Private Sub DirList4_Change()
    ' Update File listbox to sync with Dir listbox.
    filList4.Path = dirList4.Path
    filList4.Pattern = txtSearchSpec4.Text
    If txtTexto4.Text <> "" Then txtTexto4.SetFocus
End Sub

Private Sub cmdLimpiar4_Click()
    txtTexto4.Text = ""
    txtTexto4.SetFocus
End Sub

Private Sub cmdHablar4_Click()
    Hablar_Leer (txtTexto4.Text)
    txtTexto4.SetFocus
    txtTexto4.SelStart = 0
    txtTexto4.SelLength = Len(txtTexto4.Text)
    'AnimeLextor.Hide
End Sub

Private Sub cmdDetener4_Click()
     AnimeLextor.Stop
     txtTexto4.SetFocus
End Sub

Private Sub cmdVer4_Click()
    txtTexto4.SetFocus
    If Not ImgLetra4.Picture = 0 Then
       SeekView.Show
       SeekView.Image1.Picture = ImgLetra4.Picture
       Load SeekView
    End If
End Sub

Private Sub cmdAleatorio4_Click()
    Dim P As Integer
    Segundos3 = 0
    If NumRegistros <> 0 Then
       'Seleccionamos aletatoriamente la palabra
        P = Int((Rnd * NumRegistros))
        txtTexto4.Text = XLetra(P).Nombre 'XLetra(P).Letra
        Select Case XLetra(P).Letra
        Case "V"
'            Hablar_Leer (XLetra(P).Letra & "e de " & XLetra(P).Nombre)
            Hablar_Leer XLetra(P).Nombre
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
'           Hablar_Leer XLetra(P).Letra
            Hablar_Leer XLetra(P).Nombre
        Case Else
            Hablar_Leer XLetra(P).Nombre
        End Select
        ImgLetra4.Picture = LoadPicture(XLetra(P).Archivo)
    End If
End Sub


Sub Arreglo4()
    Dim NumPosiciones As Byte
    Dim P, I As Integer
    Dim xTexto As String
    If (lstFoundFiles4.ListCount - 1) = -1 Then Exit Sub
    NumRegistros = lstFoundFiles4.ListCount
    ReDim XLetra(NumRegistros)
    For I = 0 To lstFoundFiles4.ListCount - 1
        For P = 0 To Len(lstFoundFiles4.List(I))
            xTexto = Mid(lstFoundFiles4.List(I), 1, Len(lstFoundFiles4.List(I)) - P)
            If Right(Mid(lstFoundFiles4.List(I), 1, Len(lstFoundFiles4.List(I)) - P), 1) = "\" Then
               P = Len(xTexto)
               Exit For
            End If
        Next
        XLetra(I).Nombre = Mid(lstFoundFiles4.List(I), (Len(xTexto) + 1), Len(lstFoundFiles4.List(I)))
        XLetra(I).Nombre = Mid(XLetra(I).Nombre, 1, (Len(XLetra(I).Nombre) - 4))
        XLetra(I).Letra = Mid(XLetra(I).Nombre, 1, 1)
        If Right(filList4.Path, 1) <> "\" Then
            XLetra(I).Archivo = lstFoundFiles4.List(I)
        Else
            XLetra(I).Archivo = lstFoundFiles4.List(I)
        End If
    Next
End Sub

Private Sub cmdSearch4_Click()
' Initialize for search, then call DirDiver to perform recursive search.
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim result As Integer
  ' Check what the user did last:
  'If cmdSearch.Caption = "&Resetear" Then  ' If just a reset,
  '  ResetSearch                         ' initialize and exit.
  '  txtSearchSpec.SetFocus
  '  Exit Sub
  'End If

  ' Update dirList.Path if it is different from the currently
  ' selected directory, otherwise perform the search.
  If dirList4.Path <> dirList4.List(dirList4.ListIndex) Then
     dirList4.Path = dirList4.List(dirList4.ListIndex)
     Exit Sub         ' Exit so user can take a look before searching.
  End If

  dirList4.Path = xDirectorioActual
  'cmdExit.Caption = "&Cancelar"

  filList4.Pattern = txtSearchSpec4.Text
  FirstPath = dirList4.Path
  DirCount = dirList4.ListCount

  'Start recursive direcory search.
  NumFiles = 0                       ' Reset global foundfiles indicator.
  result = DirDiver(FirstPath, DirCount, "")
  filList4.Path = dirList4.Path
  'cmdSearch.Caption = "&Resetear"
  'cmdSearch.SetFocus
  'cmdExit.Caption = "&Salir"
   Call Arreglo4
End Sub

Private Sub cmbLista4_Click()
    If Me.cmbLista4.Text <> "Todos" Then
        Picture4.Visible = True
        Picture5.Visible = False
        Call DirList4_LostFocus
        Call Arreglo4
        'Call DirList_Change
    Else
      ' Continue with the search.
        Picture4.Visible = False
        Picture5.Visible = True
        Call cmdSearch4_Click
    End If
End Sub

Private Sub DirList4_LostFocus()
    Dim xDirectorio As String
    Dim ind As Integer
    If Me.cmbLista4.Text <> "Todos" Then
       xDirectorio = cmbLista4.Text
       filList4.Path = xDirectorioActual 'dirList.Path
       dirList4.Path = filList4.Path & "\" & xDirectorio
       If txtTexto4.Text <> "" Then txtTexto4.SetFocus
    Else
        ReDim XLetra(filList4.ListCount)
       For ind = 0 To cmbLista4.ListCount - 1        ' Add conforming files in
           xDirectorio = cmbLista4.List(ind)     ' this directory to listbox.
           filList4.Path = xDirectorioActual 'dirList.Path
           If xDirectorio <> "Todos" Then
           dirList4.Path = filList4.Path & "\" & xDirectorio
           Call Arreglo
           End If
       Next ind
       xDirectorio = "Animales"
       If txtTexto4.Text <> "" Then txtTexto4.SetFocus
    
    End If

End Sub

Private Sub DrvList4_Change()
    On Error GoTo DriveHandler
    dirList4.Path = drvList4.Drive
    txtTexto4.SetFocus
    Exit Sub

DriveHandler:
    drvList4.Drive = dirList4.Path
    Exit Sub
End Sub

Private Sub filList4_Click()
    On Error GoTo Error
      If Right(filList4.Path, 1) <> "\" Then
        Text14.Text = filList4.Path & "\" & filList4.FileName
      Else
        Text14.Text = filList4.Path & filList4.FileName
      End If
    Dim dropfile As String, X, nl, msg
    dropfile = Text14.Text
    Select Case Right(LCase(Text14.Text), 3)
    Case "txt"
        Me.Label1.Caption = "Archivo de Texto"
    Case "bmp", "wmf", "rle", "ico", "dib", "jpg"
        Me.ImgLetra.Picture = LoadPicture(Text14.Text)
    Case "exe"
        Me.Label1.Caption = "Archivo Ejecutable"
    Case "hlp"
        Me.Label1.Caption = "Archivo de Ayuda"
    Case Else
        nl = Chr$(10) + Chr$(13)
        msg = "Es un Tipo de Archivo Desconocido:"
        msg = nl + msg + nl + nl + "     .ico, .bmp, .wmf, .rle, .jpg"
        MsgBox msg
    End Select
    Call Tiempo_Act
    txtTexto4.Text = Mid(filList4.FileName, 1, Len(filList4.FileName) - 4)
    Call cmdHablar4_Click
    txtTexto4.SetFocus
    Exit Sub
    
Error:
    txtTexto4.SetFocus
End Sub

Private Sub lstFoundFiles4_Click()
    Text14.Text = lstFoundFiles4
    Dim dropfile As String, X, nl, msg
    dropfile = Text14.Text
    Select Case Right(LCase(Text14.Text), 3)
    Case "txt"
        Me.Label1.Caption = "Archivo de Texto"
    Case "bmp", "wmf", "rle", "ico", "dib", "jpg"
        Me.ImgLetra.Picture = LoadPicture(Text14.Text)
    Case "exe"
        Me.Label1.Caption = "Archivo Ejecutable"
    Case "hlp"
        Me.Label1.Caption = "Archivo de Ayuda"
    Case Else
        nl = Chr$(10) + Chr$(13)
        msg = "Es un Tipo de Archivo Desconocido:"
        msg = nl + msg + nl + nl + "     .ico, .bmp, .wmf, .rle , .jpg"
        MsgBox msg
    End Select
    X = lstFoundFiles4.ListIndex
    txtTexto4.Text = XLetra(X).Nombre
    Call cmdHablar_Click
End Sub

'**************************************************************************
'   Iniciar Imagenes
'
'**************************************************************************
Private Sub cmdIniciar_Click(Index As Integer)
    Dim I As Integer
    If cmdIniciar(Index).Caption = "&Iniciar" Then
        For I = 0 To 2
            cmdIniciar(I).Caption = "&Detener"
        Next
        AnimeLextor.Left = 600
        AnimeLextor.Top = 50
    Else
        For I = 0 To 2
            cmdIniciar(I).Caption = "&Iniciar"
        Next
        AnimeLextor.Left = 600
        AnimeLextor.Top = 450
    End If
End Sub
Private Sub ImgLetra5_Click()
    SeekView.Show
    If Picture2.Visible = False Then
       SeekView.Image1.Picture = ImgLetra5.Picture
    Else
       SeekView.Image1.Picture = ImgLetra5.Picture
    End If
    Load SeekView
End Sub

'**************************************************************************
'   Iniciar Matematica
'
'**************************************************************************
Private Sub cmbNivel_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdComenzar_Click()
       Dim Valor
       Valor = 10
       Select Case cmbNivel.Text
       Case "Nivel 1"
          Valor = 10
       Case "Nivel 2"
          Valor = 100
       Case "Nivel 3"
          Valor = 1000
       Case "Nivel 4"
          Valor = 100000
       End Select
       lblN0 = Int((Rnd * Valor))
       lblN1 = Int((Rnd * Valor))
End Sub

Private Sub chkRespuesta_Click()
        If Option1(0).Value = True Then
            txtRespuesta = Val(lblN0) + Val(lblN1)
        End If
        If Option1(1).Value = True Then
            txtRespuesta = Val(lblN0) - Val(lblN1)
        End If
        If Option1(2).Value = True Then
            txtRespuesta = Val(lblN0) * Val(lblN1)
        End If
        If Option1(3).Value = True Then
            txtRespuesta = Val(lblN0) / Val(lblN1)
        End If
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
    Case 0:
       Me.lblSigno = "+"
    Case 1:
       Me.lblSigno = "-"
    Case 2:
       Me.lblSigno = "*"
    Case 3:
       Me.lblSigno = "÷"
    End Select
End Sub

Private Sub txtRespuesta_KeyPress(KeyAscii As Integer)
    Dim xResp, xMsg
    xResp = 0
    If KeyAscii = 13 Then
        If Option1(0).Value = True Then
            xResp = Val(lblN0) + Val(lblN1)
            xMsg = "Suma"
        End If
        If Option1(1).Value = True Then
            xResp = Val(lblN0) - Val(lblN1)
            xMsg = "Resta"
        End If
        If Option1(2).Value = True Then
            xResp = Val(lblN0) * Val(lblN1)
            xMsg = "Suma"
        End If
        If Option1(3).Value = True Then
            xResp = Val(lblN0) - Val(lblN1)
            xMsg = "Resta"
        End If
        If xResp = txtRespuesta Then
           'MsgBox "Respuesta Correcta en la " + xMsg, vbDefaultButton1, "Respuesta"
           Hablar_Leer "Respuesta Correcta en la " + xMsg
        Else
           Hablar_Leer "Respuesta InCorrecta en la " + xMsg
           'MsgBox "Respuesta inCorrecta ", vbDefaultButton1 + vbCritical, "Respuest"
        End If

    End If
End Sub
