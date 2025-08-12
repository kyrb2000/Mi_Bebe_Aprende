VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmCalendario 
   Caption         =   "Calendario"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   5
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdHoy 
      Caption         =   "Hoy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   4575
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5895
      _Version        =   524288
      _ExtentX        =   10398
      _ExtentY        =   8070
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2013
      Month           =   12
      Day             =   28
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblx3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblx2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label lblx1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      Top             =   1080
      Width           =   3015
   End
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar1_Click()
    Text1 = Str(Calendar1.Day) + "/" + Str(Calendar1.Month) + "/" + Str(Calendar1.Year)
    'Calendar1.Day = Day(Date)
    'Calendar1.Month = Month(Date)
    'Calendar1.Year = Year(Date)
    lblx1 = Val(Mid(Calendar1.Year, 1, 1)) + Val(Mid(Calendar1.Year, 2, 1)) + Val(Mid(Calendar1.Year, 3, 1)) + Val(Mid(Calendar1.Year, 4, 1))
    lblx2 = lblx1 * 11
    lblx3 = lblx2 + Val(Calendar1.Day) + Val(Calendar1.Month) - 90
End Sub

Private Sub cmdHoy_Click()
    Text1 = Date
    Calendar1.Day = Day(Date)
    Calendar1.Month = Month(Date)
    Calendar1.Year = Year(Date)
    lblx1 = Val(Mid(Calendar1.Year, 1, 1)) + Val(Mid(Calendar1.Year, 2, 1)) + Val(Mid(Calendar1.Year, 3, 1)) + Val(Mid(Calendar1.Year, 4, 1))
    lblx2 = lblx1 * 11
    lblx3 = lblx2 + Val(Calendar1.Day) + Val(Calendar1.Month) - 60 - 30
End Sub
