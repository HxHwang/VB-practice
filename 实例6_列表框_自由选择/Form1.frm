VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   1710
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1815
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":001F
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "�ڴ����ѡ��"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "����Ԥ��2006���籭�ھ���"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub List1_Click()
Select Case List1.ListIndex
Case 0
Label2.Caption = "������ʵ�������۹�"
Case 1
Label2.Caption = "�������ݣ�ʱ�˲���"
Case 2
Label2.Caption = "�ؼ�ʱ������"
Case 3
Label2.Caption = "�¹�ս���ƾɲ���"
Case 4
Label2.Caption = "����������������"
Case 5
Label2.Caption = "�����﷭����������"
Case 6
Label2.Caption = "�ִ������ܿ����ض���"
Case 7
Label2.Caption = "�ѳɴ�����"
Case 8
Label2.Caption = "��ϲ�㣬����ˣ�"
End Select

End Sub


