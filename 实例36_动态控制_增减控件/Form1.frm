VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��̬�����ؼ�������"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "ɾ���ؼ�"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���ӿؼ�"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    'ͨ��ʹ��WithEvents�ؼ�������һ���������Ϊ�µ����ť
    Private WithEvents NewButton As CommandButton
Attribute NewButton.VB_VarHelpID = -1
    '���ӿؼ�
    Private Sub Command1_Click()
    If NewButton Is Nothing Then
    '�����µİ�ťcmdNew
    Set NewButton = Controls.Add("VB.CommandButton", "cmdNew", Me)
    'ȷ��������ťcmdNew��λ��
    NewButton.Move Command1.Left + Command1.Width + 240, Command1.Top
    NewButton.Caption = "������ť"
    NewButton.Visible = True
    End If
    End Sub
    'ɾ���ؼ�(ע��ֻ��ɾ����̬���ӵĿؼ�)
    Private Sub Command2_Click()
    If NewButton Is Nothing Then
    Else
    Controls.Remove NewButton
    Set NewButton = Nothing
    End If
    End Sub
    '�����ؼ��ĵ����¼�
    Private Sub NewButton_Click()
     MsgBox "��ѡ�е��Ƕ�̬���ӵİ�ť!"
    End Sub

