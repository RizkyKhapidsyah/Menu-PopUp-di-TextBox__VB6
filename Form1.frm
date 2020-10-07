VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menu Pop-Up di TextBox"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Menu MyMenu 
      Caption         =   "Menu"
      Begin VB.Menu menuFile 
         Caption         =   "File"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WM_RBUTTONDOWN = &H204
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub OpenContextMenu(FormName As Form, menuName As Menu)
  Call SendMessage(FormName.hwnd, WM_RBUTTONDOWN, 0, 0&)
  FormName.PopupMenu menuName
End Sub

Private Sub Form_Load()
  MyMenu.Visible = False  'Agar tdk kelihatan di bagian atas form
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Ganti 'MyMenu' dengan menu yang Anda inginkan tampil
  'secara pop up.
If Button = vbRightButton Then _
    Call OpenContextMenu(Me, Me.MyMenu)
End Sub


