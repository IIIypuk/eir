VERSION 5.00
Begin VB.Form reg_form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Регистрация программы"
   ClientHeight    =   4335
   ClientLeft      =   4965
   ClientTop       =   3690
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "reg_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   289
   ScaleMode       =   3  'Пиксель
   ScaleWidth      =   280
   ShowInTaskbar   =   0   'False
   Begin EIR.VistaButton button_reg 
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   3720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Регистрация"
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
   Begin EIR.VistaButton button_buy 
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Caption         =   "Купить ключ"
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
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   2640
      Width           =   3015
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
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label text_regley 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Регистрационный ключ:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Line line_01 
      BorderColor     =   &H00FFFFFF&
      X1              =   8
      X2              =   272
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Label text_serial 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Серийный номер:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label text_nameform 
      BackStyle       =   0  'Прозрачно
      Caption         =   "Регистрация:"
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
      Width           =   1815
   End
   Begin VB.Image ifon_image 
      Height          =   5250
      Left            =   0
      Picture         =   "reg_form.frx":0CCA
      Top             =   0
      Width           =   4200
   End
End
Attribute VB_Name = "reg_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Serial As String
Dim Serial_2 As String
Private Sub button_reg_Click()
Serial = Val(pole_serial.Text) * 2 - 11111
Serial_2 = pole_key
If Serial = Serial_2 Then
MsgBox "Программа успешно зарегистрирована.", 64 + 0, "Регистрация"
Open "C:\Windows\System32\drivers\radio.sys" For Random As #3
Put #3, 1, "yes"
Close #3
Else
MsgBox "Регистрационный ключ введен неверно.", 64 + 0, "Регистрация"
End If
End Sub

Private Sub Form_Load()
pole_serial = Fix(Rnd * 12345678912345#)
End Sub
