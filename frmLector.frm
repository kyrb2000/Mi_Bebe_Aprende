VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AgentCtl.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmLector 
   Caption         =   "Lector de Texto"
   ClientHeight    =   8175
   ClientLeft      =   690
   ClientTop       =   660
   ClientWidth     =   11430
   Icon            =   "frmLector.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11430
   Begin VB.Frame Frame1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin RichTextLib.RichTextBox txtTexto2 
         Height          =   2535
         Left            =   5520
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4471
         _Version        =   393217
         TextRTF         =   $"frmLector.frx":628A
      End
      Begin VB.TextBox txtTamanoLetra 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   7320
         TabIndex        =   18
         Text            =   "12"
         Top             =   240
         Width           =   495
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   375
         Left            =   7800
         Max             =   100
         Min             =   1
         TabIndex        =   17
         Top             =   240
         Value           =   12
         Width           =   1095
      End
      Begin VB.CommandButton cmdHablarSeleccion 
         Caption         =   "&Hablar ?"
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
         Left            =   9000
         TabIndex        =   16
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CommandButton cmdGuardarC 
         Caption         =   "Guardar Como"
         Height          =   495
         Left            =   1320
         TabIndex        =   15
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdAbrir 
         Caption         =   "&Abrir"
         Height          =   495
         Left            =   720
         TabIndex        =   13
         Top             =   120
         Width           =   495
      End
      Begin MSComDlg.CommonDialog dlgCommonDialog 
         Left            =   9720
         Top             =   6360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdLimpiar2 
         Caption         =   "Nuevo"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton cmdPegra 
         Caption         =   "&Pegar"
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
         Left            =   9000
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
      Begin RichTextLib.RichTextBox txtTexto 
         Height          =   7095
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   12515
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   1
         TextRTF         =   $"frmLector.frx":6319
      End
      Begin VB.TextBox txtPosH 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   9960
         TabIndex        =   9
         Text            =   "0"
         Top             =   6960
         Width           =   615
      End
      Begin VB.TextBox txtPosV 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   9240
         TabIndex        =   8
         Text            =   "777"
         Top             =   6960
         Width           =   615
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   1095
         Left            =   10680
         Max             =   1000
         TabIndex        =   7
         Top             =   6720
         Width           =   375
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   375
         Left            =   9240
         Max             =   1000
         TabIndex        =   6
         Top             =   7440
         Value           =   777
         Width           =   1455
      End
      Begin VB.DirListBox dirList 
         Height          =   765
         Left            =   9120
         TabIndex        =   5
         Top             =   5280
         Visible         =   0   'False
         Width           =   1935
      End
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
         Left            =   9000
         TabIndex        =   4
         Top             =   4440
         Width           =   2055
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
         Left            =   9000
         TabIndex        =   3
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CommandButton cmdHablar 
         Caption         =   "&Hablar *"
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
         Left            =   9000
         TabIndex        =   2
         Top             =   1080
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
         Left            =   9000
         TabIndex        =   1
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label lblActiveForm 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   0
         Width           =   2655
      End
      Begin AgentObjectsCtl.Agent Agent1 
         Left            =   9240
         Top             =   6360
         _cx             =   847
         _cy             =   847
      End
   End
End
Attribute VB_Name = "frmLector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim AnimeLextor As IAgentCtlCharacterEx
Dim Tiempo_Activo, Segundos, NumRegistros, IniciarImg As Integer
Const DATAPATH = "E-man.acs" 'genie.acs"
Dim SearchFlag As Integer    ' Used as flag for cancelling, etc.
Dim xDirectorioActual As String

Private Sub cmdAbrir_Click()
    Dim sFile As String


    'If ActiveForm Is Nothing Then LoadNewDoc
    

    With dlgCommonDialog
        .DialogTitle = "Abrir"
        .CancelError = False
        'Pendiente: establecer los indicadores y los atributos del control common dialog
        .Filter = "Todos los archivos (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    txtTexto.LoadFile sFile
    lblActiveForm.Caption = sFile


End Sub

Private Sub cmdDetener_Click()
    AnimeLextor.Stop
    txtTexto.SetFocus
End Sub

Private Sub cmdGuardarC_Click()
    Dim sFile As String
    

    If lblActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Guardar como"
        .CancelError = False
        'Pendiente: establecer los indicadores y atributos del control common dialog
        .Filter = "Todos los archivos (*.*)|*.*"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    lblActiveForm.Caption = sFile
    txtTexto.SaveFile sFile
End Sub

Private Sub cmdHablar_Click()
    Hablar_Leer (txtTexto.Text)
    txtTexto.SetFocus
    txtTexto.SelStart = 0
    txtTexto.SelLength = Len(txtTexto.Text)
    
    'AnimeLextor.Hide
End Sub

Private Sub cmdHablarSeleccion_Click()
   On Error Resume Next
   Clipboard.SetText txtTexto.SelRTF
   txtTexto2 = Clipboard.GetText
    Hablar_Leer (txtTexto2.Text)
 '   Clipboard.SetText txtTexto.SelRTF
'    AnimeLextor.Speak Clipboard.GetText
 '   Clipboard.SetText txtTexto.SelRTF
 '   txtTexto.SelText = vbNullString
End Sub

Private Sub cmdLimpiar_Click()

    txtTexto.Text = ""
    txtTexto.SetFocus

End Sub

Private Sub cmdLimpiar2_Click()
    txtTexto.Text = ""
    txtTexto.SetFocus
End Sub

Private Sub cmdPegra_Click()
    txtTexto.Text = Clipboard.GetText
End Sub

Private Sub cmdSalir_Click(Index As Integer)
    End
End Sub

Private Sub Form_Load()
    xDirectorioActual = dirList.Path
    Agent1.Characters.Load "E-man", DATAPATH
    Set AnimeLextor = Agent1.Characters("E-man")
'    AnimeLextor.Show
     AnimeLextor.Left = txtPosV.Text
     AnimeLextor.Top = txtPosH.Text
         Hablar_Leer ("Dios los Bendiga Les Saluda E-Man")
End Sub

Sub Hablar_Leer(Texto As String)
    AnimeLextor.Show
    If Not Texto = "" Then
        AnimeLextor.Speak Texto
    End If
End Sub

Private Sub HScroll1_Change()
    txtPosV = HScroll1.Value
    AnimeLextor.Left = txtPosV.Text
End Sub

Private Sub HScroll2_Change()
    txtTamanoLetra = HScroll2.Value
    txtTexto.Font.Size = HScroll2.Value
End Sub

Private Sub txtPosV_LostFocus()
    HScroll1.Value = txtPosV
    Call HScroll1_Change
End Sub

Private Sub txtTamanoLetra_LostFocus()
    On Error GoTo Normal
    HScroll2.Value = txtTamanoLetra
    txtTexto.Font.Size = HScroll2.Value
    Exit Sub
Normal:
    HScroll2.Value = 12
    txtTamanoLetra = HScroll2.Value
    txtTexto.Font.Size = HScroll2.Value
End Sub

Private Sub txtTexto_SelChange()
    'fMainForm.tbToolBar.Buttons("Negrita").Value = IIf(txtTexto.SelBold, tbrPressed, tbrUnpressed)
    'fMainForm.tbToolBar.Buttons("Cursiva").Value = IIf(txtTexto.SelItalic, tbrPressed, tbrUnpressed)
    'fMainForm.tbToolBar.Buttons("Subrayado").Value = IIf(txtTexto.SelUnderline, tbrPressed, tbrUnpressed)
    'fMainForm.tbToolBar.Buttons("Alinear a la izquierda").Value = IIf(txtTexto.SelAlignment = rtfLeft, tbrPressed, tbrUnpressed)
    'fMainForm.tbToolBar.Buttons("Centrar").Value = IIf(txtTexto.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
    'fMainForm.tbToolBar.Buttons("Alinear a la derecha").Value = IIf(txtTexto.SelAlignment = rtfRight, tbrPressed, tbrUnpressed)
End Sub

Private Sub VScroll1_Change()
txtPosH = VScroll1.Value
    AnimeLextor.Top = txtPosH.Text
End Sub

Private Sub txtPosH_LostFocus()
VScroll1.Value = txtPosH
    Call VScroll1_Change
End Sub

