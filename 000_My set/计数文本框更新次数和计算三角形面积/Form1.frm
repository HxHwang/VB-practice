VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�ı���Ӧ��ʾ��"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6825
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "��ʾ�������������"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Text1_Change()
Static a%
a = a + 1
Label1.Caption = a
End Sub

Private Sub Text2_Change()
Label2.Caption = Len(Text2.Text)
End Sub
