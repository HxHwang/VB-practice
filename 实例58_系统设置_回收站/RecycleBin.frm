VERSION 5.00
Begin VB.Form RecycleBin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Դ����վ����"
   ClientHeight    =   2115
   ClientLeft      =   2715
   ClientTop       =   1770
   ClientWidth     =   4245
   Icon            =   "RecycleBin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "�����Դ����վ"
      Height          =   1050
      Left            =   2160
      Picture         =   "RecycleBin.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�����Դ����վ"
      Height          =   1050
      Left            =   600
      Picture         =   "RecycleBin.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��Դ����վ�а��� �� Byte"
      Height          =   195
      Left            =   570
      TabIndex        =   3
      Top             =   1680
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��Դ����վ�а��� �� ���ļ�"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   1410
      Width           =   2250
   End
End
Attribute VB_Name = "RecycleBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim rbinfo As SHQUERYRBINFO  ' ��Դ����վ����Ѷ
    Dim retval As Long           ' ����ֵ
    ' ��ʼ�� rbinfo �Ĵ�С
    rbinfo.cbSize = Len(rbinfo)
    ' ��ѯ��Դ����վ������
    retval = SHQueryRecycleBin("C:\", rbinfo)  ' the path doesn't have to be the root path
    ' ��ʾ��Դ����վ��Ŀǰ�ж������
    If (rbinfo.i64NumItems.LowPart And &H80000000) = &H80000000 Or rbinfo.i64NumItems.HighPart > 0 Then
        Label1 = "C���̻���վ���г��� 2,147,483,647 ���ļ�"
    Else
        Label1 = "C���̻���վ�а��� " & rbinfo.i64NumItems.LowPart & " ���ļ�"
    End If
    ' ��ʾ��Դ����վ�е������ռ�˶��� Bytes��
    If (rbinfo.i64Size.LowPart And &H80000000) = &H80000000 Or rbinfo.i64Size.HighPart > 0 Then
        Label2 = "C���̻���վ���ļ��������� 2,147,483,647 Byte"
    Else
        Label2 = "C���̻���վ���ļ����� " & rbinfo.i64Size.LowPart & " Byte"
    End If
End Sub

Private Sub Command2_Click()
    Dim retval As Long  ' return value
    ' ���������Դ����վ, ��ȷ��
    retval = SHEmptyRecycleBin(RecycleBin.hwnd, "", SHERB_NOCONFIRMATION)
    ' ���д���ѶϢ���֣���ظ���Դ����վ��ͼʾ
    ' ��ʵ��һ�㲻�Ǻ���Ҫ
    If retval <> 0 Then  ' error
        retval = SHUpdateRecycleBinIcon()
    End If
    Command1_Click
End Sub

