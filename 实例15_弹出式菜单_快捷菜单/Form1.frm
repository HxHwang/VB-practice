VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1590
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2850
   LinkTopic       =   "Form1"
   ScaleHeight     =   1590
   ScaleWidth      =   2850
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "��������Ҽ�������ʹ�ÿ�ݲ˵���"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Menu Filemenu 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu NewMenu 
         Caption         =   "�½�(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu OpenMenu 
         Caption         =   "��(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu SaveMenu 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu Menu1 
         Caption         =   "-"
      End
      Begin VB.Menu CloseMenu 
         Caption         =   "�ر�"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "�༭(&E)"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu Filemenu
End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu Filemenu
End If
End Sub
