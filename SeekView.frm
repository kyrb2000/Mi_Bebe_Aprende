VERSION 5.00
Begin VB.Form SeekView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12060
   Icon            =   "SeekView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmpTiempo 
      Interval        =   500
      Left            =   360
      Top             =   240
   End
   Begin VB.Image Image1 
      Height          =   8535
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "SeekView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Segundos2 As Integer
Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Segundos2 = 0
    Image1.Width = Me.Width
    Image1.Height = Me.Height
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub


Private Sub tmpTiempo_Timer()
        Segundos2 = Segundos2 + 1
        If Segundos2 = 11 Then
           Unload Me
        End If
End Sub
