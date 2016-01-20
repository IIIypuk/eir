VERSION 5.00
Begin VB.Form main_form 
   Caption         =   "EIR Register"
   ClientHeight    =   3030
   ClientLeft      =   4845
   ClientTop       =   4350
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   202
   ScaleMode       =   3  'Пиксель
   ScaleWidth      =   280
   Begin keygen.VistaButton button_gen 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Показать"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox pole_key 
      Alignment       =   1  'Правая привязка
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox pole_serial 
      Alignment       =   1  'Правая привязка
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label text_key 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Рег. ключ:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Серийный номер:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   8
      X2              =   272
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Label text_nameprog 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Регистратор"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   5250
      Left            =   0
      Picture         =   "main.frx":0000
      Top             =   0
      Width           =   4200
   End
End
Attribute VB_Name = "main_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub button_gen_Click()
pole_key.Text = Fix(Val(pole_serial.Text) * 2 - 11111)
End Sub
