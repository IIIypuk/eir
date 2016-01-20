VERSION 5.00
Begin VB.Form info_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "О программе"
   ClientHeight    =   3795
   ClientLeft      =   4965
   ClientTop       =   3495
   ClientWidth     =   4200
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "info_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   253
   ScaleMode       =   3  'Пиксель
   ScaleWidth      =   280
   Begin EIR.VistaButton Button_close 
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      Caption         =   "Закрыть"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label text_site 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Web-сайт: http://net-popov.ucoz.ru"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label text_email 
      BackStyle       =   0  'Прозрачно
      Caption         =   "email: man_x@mail.ru"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label text_develop 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Прозрачно
      Caption         =   "Разработчик: Александр Попов          aka IIIypuk"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label text_version 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Версия: v1.0 Alpha Build 2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Line line_02 
      BorderColor     =   &H00FFFFFF&
      X1              =   32
      X2              =   32
      Y1              =   56
      Y2              =   216
   End
   Begin VB.Line line_01 
      BorderColor     =   &H00FFFFFF&
      X1              =   16
      X2              =   264
      Y1              =   64
      Y2              =   64
   End
   Begin VB.Image fon_image 
      Height          =   5250
      Left            =   0
      Picture         =   "info_form.frx":0CCA
      Top             =   0
      Width           =   4200
   End
End
Attribute VB_Name = "info_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button_close_Click()
info_form.Hide
End Sub
