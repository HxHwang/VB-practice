VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��������"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1812
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   372
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   852
   End
   Begin VB.Label Label2 
      Caption         =   "���ú��������1 �� N �ĺ�"
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "������ N ֵ"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'��������Ҫ��ϰ�Ӷ��庯��
'ע���ʽ�Ͳ���
'ע��һ�µ��ú����ĳ���ִ��˳��

Private Sub Command1_Click()

    Dim a As Integer
    Dim b As Integer
    a = Form1.Text1.Text   'a ��������� N ֵ�������ı�������
    
    b = Plus_N(a) '���ú������̣��������Ҫ����������ִ�к�����Ȼ�������������ִ��
    
    Form1.Text1.Text = b  'b ���Ǵ������ۼӺ�
End Sub
'���ﶨ��������������N ��������һ������M ���������ۼӵĺ�,Plus_N ���Լ��������
Private Function Plus_N(N As Integer)

    Dim I As Integer
    Dim Sum As Integer
    For I = 1 To N
        Sum = Sum + I
    Next I
    Plus_N = Sum
   
End Function




