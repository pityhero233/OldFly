VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "����..."
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "��֪����(&S)"
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "PS���������Զ����С���������ȡ�����乤���ң�������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "�������Զ��ѣ���Զ���ɣ��κ����ڲ�����ҵΪĿ�ĵ�ǰ���¿���������ġ�ʹ�ü�������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
frmMain.Show

End Sub
