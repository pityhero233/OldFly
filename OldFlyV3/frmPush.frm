VERSION 5.00
Begin VB.Form frmPush 
   Caption         =   "�Ϸ����´��һݰ� V3.0"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPush.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "ȷ��"
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Ҳ�֪��"
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "�����ǵػ���_____"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmPush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim times As Integer
Private Sub Command1_Click()
n = InputBox("������༶����:")

frmMain.Show
Unload Me

End Sub

Private Sub Command2_Click()
times = times + 1
If times >= 10 Then
n = 31
frmMain.Show
Unload Me
End If
End Sub

Private Sub Form_Load()
times = 0
End Sub
