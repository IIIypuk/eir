VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form player 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Easy Internet Radio v1.0 Alpha Build 2"
   ClientHeight    =   4320
   ClientLeft      =   4965
   ClientTop       =   3795
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "player.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   288
   ScaleMode       =   3  'Пиксель
   ScaleWidth      =   280
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox list_radio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      ItemData        =   "player.frx":0CCA
      Left            =   120
      List            =   "player.frx":0CCC
      TabIndex        =   0
      Top             =   960
      Width           =   3975
   End
   Begin WMPLibCtl.WindowsMediaPlayer player_radio 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   3975
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   7011
      _cy             =   1720
   End
   Begin VB.Image fon_image 
      Height          =   5250
      Left            =   0
      Picture         =   "player.frx":0CCE
      Top             =   0
      Width           =   4200
   End
   Begin VB.Menu up_player 
      Caption         =   "Плеер"
      Begin VB.Menu up_player_update 
         Caption         =   "Обновить список радио"
      End
      Begin VB.Menu up_player_close 
         Caption         =   "Закрыть"
      End
   End
   Begin VB.Menu up_help 
      Caption         =   "Помощь"
      Begin VB.Menu up_help_help 
         Caption         =   "Справка"
      End
      Begin VB.Menu up_help_site 
         Caption         =   "Посетить сайт"
      End
      Begin VB.Menu up_reg 
         Caption         =   "Регистрация"
      End
      Begin VB.Menu up_help_update2 
         Caption         =   "Обновить программу"
      End
      Begin VB.Menu up_help_info 
         Caption         =   "О программе"
      End
   End
End
Attribute VB_Name = "player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If App.PrevInstance = True Then
MsgBox "Запуск двух копий программ невозможен.", 64 + 0, "Easy Internet Radio"
End
End If
Dim RList As String
Open App.Path & "\Data\Radio list\radio_list.erl" For Input As #2
Do Until EOF(2)
    Line Input #2, RList
    list_radio.AddItem RList
Loop
Close #2
End Sub

Private Sub list_radio_Click()
Open App.Path & "\Data\Radio list\" & list_radio.Text & ".eurl" For Input As #1
    Do Until EOF(1)
    Line Input #1, URL
    player_radio.URL = URL
    Loop
Close #1
End Sub

Private Sub up_help_help_Click()
MsgBox "Пока недоделал.", 64 + 0, "Ошибка"
End Sub

Private Sub up_help_info_Click()
info_form.Show
End Sub

Private Sub up_help_site_Click()
MsgBox "Пока недоделал.", 64 + 0, "Ошибка"
End Sub

Private Sub up_help_update2_Click()
MsgBox "Информация: Перед выполнением обновления, программа проверит версию на сервере и если ваша версия программы старее то что на сервере, тогда произойдет обновление. После успешного обновления запустится новая версия.", 64 + 0, "Обновление программы"
End Sub

Private Sub up_player_close_Click()
End
End Sub

Private Sub up_player_update_Click()
MsgBox "Пока недоделал.", 64 + 0, "Ошибка"
End Sub

Private Sub up_reg_Click()
reg_form.Show
End Sub
