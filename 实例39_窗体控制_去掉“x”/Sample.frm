VERSION 5.00
Begin VB.Form frmSample 
   Caption         =   "ȥ���رհ�ť - ����"
   ClientHeight    =   1410
   ClientLeft      =   4005
   ClientTop       =   2205
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1410
   ScaleWidth      =   3315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "�˳�����"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'��һ�ַ���
    Dim hwndMenu As Long
    Dim c As Long
    hwndMenu = GetSystemMenu(Me.hwnd, 0)
    
    c = GetMenuItemCount(hwndMenu)
    
    DeleteMenu hwndMenu, c - 1, MF_BYPOSITION
    
    c = GetMenuItemCount(hwndMenu)
    DeleteMenu hwndMenu, c - 1, MF_BYPOSITION
    
'�ڶ��ַ���
    'Call DisableX(Me)
End Sub

Private Sub Command1_Click()
    End
End Sub

