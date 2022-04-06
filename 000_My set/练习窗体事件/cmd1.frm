VERSION 5.00
Begin VB.Form cmd1 
   Caption         =   "练习窗体事件"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6450
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "cmd1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
cmd1.Width = Width + 500

End Sub

Private Sub Form_DblClick()
End
End Sub

Private Sub Form_Load()
cmd1.Picture = LoadPicture("c:\a.jpg")
End Sub

