VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3836
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fsoSys As New Scripting.FileSystemObject
Dim fsoRootFolder As Folder

Private Sub Form_Load()
    Dim fsoSubFolder As Folder
    Dim nodRootNode As Node
    Dim nodChild As Node
    Dim astr$
    
    Set nodRootNode = TreeView1.Nodes.Add(, , "Root", "c:\")
    Set fsoRootFolder = fsoSys.GetFolder("c:\")
    For Each fsoSubFolder In fsoRootFolder.SubFolders
        astr = fsoSubFolder.Path
        Set nodChild = TreeView1.Nodes.Add("Root", tvwChild, astr, fsoSubFolder.Name)
    Next
    
    Set fsoRootFolder = Nothing
    Command1.Caption = "����Ŀ¼"
    Command2.Caption = "ɾ��Ŀ¼"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fsoSys = Nothing
End Sub

Private Sub Command1_Click()
    Dim fsoFolder As Folder
    
    '���Ŀ¼�Ƿ����,���Ŀ¼������������Ŀ¼
    If fsoSys.FolderExists("c:\test") Then
        MsgBox ("Ŀ¼c:\test�Ѿ����ڣ��޷�����Ŀ¼")
    Else
        Set fsoFolder = fsoSys.CreateFolder("c:\test")
        Set fsoFolder = Nothing
    End If
End Sub

Private Sub Command2_Click()
    '���Ŀ¼�Ƿ����,�������ɾ��Ŀ¼
    If fsoSys.FolderExists("c:\test") Then
        fsoSys.DeleteFolder ("c:\test")
    Else
        MsgBox ("Ŀ¼c:\test������")
    End If
End Sub

