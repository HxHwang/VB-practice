VERSION 5.00
Begin VB.Form PingTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���߼���ʽ"
   ClientHeight    =   1080
   ClientLeft      =   3750
   ClientTop       =   3180
   ClientWidth     =   2355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   2355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "�������״̬"
      Height          =   525
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   1755
   End
End
Attribute VB_Name = "PingTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Delay(HowLong As Date)
    TempTime = DateAdd("s", HowLong, Now)
    While TempTime > Now
        DoEvents '�� windows ȥ����������
    Wend
End Sub

Private Sub Command1_Click()
    Dim FileFile As Integer
    Dim TestString As String
    
    '����һ�����ֵ� Test.txt��д��һ�� '0' ��
    TestString = "command.com /c echo 0 > " & "c:\Test.txt"
    Shell (TestString), vbHide
    
    '����һ�� Bat ��������� Bat ���У����ǻ��趨��
    '��� Ping һ���� Internet �ϵ� Server ���Σ������д�����ֵ� Test.txt
    '������, ������ Ping www.edu.cn Ϊ��
    FileFile = FreeFile
    Open ("c:\Test.bat") For Binary As FileFile
    TestString = "ping -n 2 www.edu.cn > " & "c:\Test.txt"
    Put #FileFile, , TestString
    Close FileFile
    
    '================
    '��ʼ����Ƿ�����
    '================
    'ִ�����ǽ����� Bat �� --> Ping
    TestString = "command.com /c " & "c:\Test.bat"
    Shell (TestString), vbHide
    '��� Ping �ɹ�, д�����ֵ� Test.txt ���ִ��������ٻ���� 200
    '�������� Ping �Ķ������ӳټ����ӣ����ԣ������ó�ʽ�ȴ� 5 ����
    Delay 5

    If FileLen("c:\Test.txt") > 201 Then
        Call MsgBox("���ĵ���Ŀǰ�Ѿ����ߵ� Internet��", vbInformation)
    Else
        Call MsgBox("���ĵ���Ŀǰ��δ���ߵ� Internet��.", vbInformation)
    End If
    
    'ɾ�������ڳ�ʽ�в����Ķ�������
    Kill "c:\Test.bat"
    Kill "c:\Test.txt"
End Sub

