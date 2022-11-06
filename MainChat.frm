VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form MainChat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dmc Chat"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "�ֶ������� W04"
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
   ScaleHeight     =   6930
   ScaleWidth      =   12015
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton DModeC 
      Caption         =   ">"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.Timer WaitB 
      Interval        =   100
      Left            =   11040
      Top             =   0
   End
   Begin MSWinsockLib.Winsock MainSck 
      Left            =   11520
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   38345
   End
   Begin VB.Frame UserFram 
      Caption         =   "��ϵ��"
      BeginProperty Font 
         Name            =   "�ֶ������� W04"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   6400
      Left            =   0
      TabIndex        =   0
      Top             =   500
      Width           =   2655
   End
   Begin VB.Line FGLine 
      X1              =   0
      X2              =   12000
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label NameTit 
      BackStyle       =   0  'Transparent
      Caption         =   "�û���"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu Hcpop 
      Caption         =   "Hcp_A"
      Begin VB.Menu ChangePort 
         Caption         =   "�������ͨ�Ŷ˿�(&P)"
      End
      Begin VB.Menu Exit 
         Caption         =   "�ǳ�(&T)"
      End
   End
   Begin VB.Menu Hdpop 
      Caption         =   "Hdp_A"
      Begin VB.Menu DCB 
         Caption         =   "Dmc Chat 2022,Biccak"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "MainChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MainconIP, MainconPort As String, UserName As String


Private Sub DModeC_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 Then PopupMenu Hcpop, , x, Y
If Button = 2 Then PopupMenu Hdpop, , x, Y
End Sub

Private Sub Exit_Click()
Me.MainSck.Close
Login.Show
Unload Me
End Sub


Private Sub MainSck_Connect()
Me.MainSck.SendData UserName
End Sub

Private Sub WaitB_Timer()
Me.WaitB.Enabled = False
MainconIP = Me.MainSck.RemoteHost
MainconPort = Me.MainSck.RemotePort
Me.MainSck.RemoteHost = ""
UserName = Me.UserFram.Caption
Me.NameTit.Caption = "�û���" & UserName & vbCrLf & "��������ַ��" & MainconIP & "   " & MainconPort
Me.MainSck.Connect MainconIP, MainconPort
End Sub
