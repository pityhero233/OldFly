VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Ϸ����´��һݰ� V3.0 ԡѪ�ع�"
   ClientHeight    =   4500
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   14505
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   4500
   ScaleWidth      =   14505
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H0080FFFF&
      Caption         =   "��ʼ / ֹͣ! ( &S )"
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "��ʷ����"
      Height          =   1095
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   12015
      Begin VB.TextBox txtHis 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "frmMain.frx":5350
         Top             =   240
         Width           =   11775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "����̨"
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1935
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�Զ��ų�"
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         Begin VB.CommandButton cmdStartup 
            Caption         =   "ֹͣ�Զ��ų�"
            Height          =   855
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Label lblPs 
         BackColor       =   &H00FF8080&
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   1695
      End
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   170
      Left            =   12480
      Top             =   3600
   End
   Begin VB.Timer tmrTime 
      Interval        =   100
      Left            =   13080
      Top             =   3600
   End
   Begin VB.TextBox txtTime 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4080
      Width           =   14415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "�Ϸ������������"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�������С�������ְ�"
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label lblCap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   11655
   End
   Begin VB.Menu Tname 
      Caption         =   "Ψһ�����û�����ʦ������������ʵ����ѧ(&6)..."
   End
   Begin VB.Menu about 
      Caption         =   "����(&H)..."
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A() As Integer
Dim currn As Integer
Dim ANotCursed() As Integer
Dim NameList(240) As String
Dim setA As Boolean

Private Sub about_Click()
frmHelp.Show
Unload Me
End Sub

Private Sub cmdStart_Click()
If tmrMain.Enabled = True Then
    'stop
    If setA = True Then
    If currn = 0 Then
    currn = GetRandom(1, n)
    End If
    tmrMain.Enabled = False
    If currn <> 0 Then
    While A(currn) = 0
        currn = A(GetRandom(1, n))
    Wend
    End If
    lblCap.Caption = NameList(currn)
    A(currn) = 0
    txtHis.Text = txtHis.Text & vbCrLf & Time & " : --- " & currn & " - " & NameList(currn) & " ---."
    Else
    
    If currn = 0 Then
    currn = GetRandom(1, n)
    End If
    tmrMain.Enabled = False
    lblCap.Caption = NameList(currn)
    txtHis.Text = txtHis.Text & vbCrLf & Time & " : --- " & currn & " - " & NameList(currn) & " ---."
    End If
Else
    'start
    tmrMain.Enabled = True
End If
End Sub

Private Sub cmdStartup_Click()
tmrMain.Enabled = False
lblCap.Caption = "---"
'MsgBox "�⹦��ûɶ�ã�Ҫ�������ˣ�", vbOKOnly
If 1 = 1 Then
If setA = True Then
setA = False
cmdStartup.Caption = "��ʼ�Զ��ų�"
For c = 1 To n
    A(c) = c
Next
currn = 1
txtHis.Text = txtHis.Text & vbCrLf & Time & " : " & "����Զ��ų���ʽ��"
Else
setA = True
cmdStartup.Caption = "ֹͣ�Զ��ų�"
txtHis.Text = txtHis.Text & vbCrLf & Time & " : " & "�����Զ��ų���ʽ��"
End If
End If
End Sub

Private Sub Form_Load()
'read namelist
If MsgBox("STU TEA", vbYesNoCancel) = vbYes Then
Open App.Path & "\namelist.txt" For Input As #1
Else
Open App.Path & "\2.txt" For Input As #1
End If
Dim tot As Integer
tot = 0
While Not EOF(1)
    tot = tot + 1
    Line Input #1, NameList(tot)
Wend
Close #1
'over.
n = tot
lblPs.Caption = "PS:���� " & n & " ��δ���鵽."
ReDim A(1 To n) As Integer
ReDim ANotCursed(1 To n) As Integer
For i = 1 To n
    A(i) = i
    ANotCursed(i) = i
Next
setA = True

End Sub

Private Sub tmrMain_Timer()
currn = GetRandom(1, n)
lblCap.Caption = NameList(currn)
If lblCap.ForeColor = vbRed Then
lblCap.ForeColor = vbBlack
Else
lblCap.ForeColor = vbRed
End If

End Sub

Private Sub tmrTime_Timer()
txtTime.Text = "��ǰʱ����:" & Date & " " & Time & "."
num = 0
If setA = True Then
For b = 1 To n
    If A(b) <> 0 Then num = num + 1
Next
If num = 0 Then
For e = 1 To n
    A(n) = e
Next
End If
lblPs.Caption = "PS:���� " & num & " ��δ���鵽."
Else
lblPs.Caption = "PS:�Զ��ų��ѹرա�"
End If
End Sub

Private Sub Tname_Click()
MsgBox "�ǵģ���û����!!!", vbInformation
End Sub

Private Sub txtHis_Change()
txtHis.SelStart = Len(txtHis.Text)

End Sub
