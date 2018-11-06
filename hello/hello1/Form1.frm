VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   6855
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton btn_loop 
      Caption         =   "循环"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton btn1 
      Caption         =   "按钮1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_loop_Click()
Dim i
Dim sum As Long: sum = 0

For i = 1 To 100
  sum = sum + i
Next
MsgBox (sum)
End Sub

Private Sub btn1_Click()
MsgBox ("按钮点击")
End Sub
