VERSION 5.00
Begin VB.Form frmMatematica 
   Caption         =   "Matematica"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   8535
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   10695
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
         Left            =   8280
         TabIndex        =   11
         Top             =   2040
         Width           =   2175
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
         ItemData        =   "frmMatematica.frx":0000
         Left            =   240
         List            =   "frmMatematica.frx":0010
         TabIndex        =   10
         Text            =   "Nivel 1"
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton cmdComenzar 
         Caption         =   "&Comenzar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8400
         TabIndex        =   9
         Top             =   4560
         Width           =   2055
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
         Left            =   720
         TabIndex        =   8
         Top             =   5880
         Width           =   9375
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
         Left            =   8280
         TabIndex        =   7
         Top             =   1560
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
         Left            =   8280
         TabIndex        =   6
         Top             =   1200
         Width           =   2175
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
         Left            =   8280
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
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
         Left            =   8280
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   1815
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
         TabIndex        =   3
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderWidth     =   10
         X1              =   720
         X2              =   10320
         Y1              =   5640
         Y2              =   5640
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
         Left            =   2640
         TabIndex        =   2
         Top             =   2880
         Width           =   5295
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
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmMatematica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
           MsgBox "Respuesta Correcta en la " + xMsg, vbDefaultButton1, "Respuesta"
           'Hablar_Leer "Respuesta Correcta en la " + xMsg
        Else
           MsgBox "Respuesta inCorrecta ", vbDefaultButton1 + vbCritical, "Respuest"
        End If

    End If
End Sub

