VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Login 
   BackColor       =   &H80000016&
   Caption         =   "DMC CHAT--��¼"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "�ֶ������� W04"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Welcome"
   ScaleHeight     =   4560
   ScaleWidth      =   8295
   StartUpPosition =   2  '��Ļ����
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer ShowXD 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7800
      Top             =   0
   End
   Begin VB.Frame InfoFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000015&
      Height          =   2175
      Left            =   1560
      TabIndex        =   10
      ToolTipText     =   "����Թر�.."
      Top             =   360
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Timer WaitUser 
         Enabled         =   0   'False
         Interval        =   1200
         Left            =   4680
         Top             =   1680
      End
      Begin VB.Label InfCas 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "���Ժ򣬳�ʼ����..."
         ForeColor       =   &H00404040&
         Height          =   1095
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   4935
      End
   End
   Begin VB.Timer WaitCancel 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   2280
   End
   Begin VB.Timer WaitTimer 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   0
      Top             =   1800
   End
   Begin VB.CommandButton CancelConn 
      BackColor       =   &H80000014&
      Caption         =   "ȡ������"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock LoginSock 
      Left            =   0
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton ForgotPass 
      Caption         =   "��������"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton LoginNow 
      BackColor       =   &H00F0FFC0&
      Caption         =   "��¼"
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton SignUP 
      Caption         =   "ע���˺�"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton SHPass 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5400
      TabIndex        =   3
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox Password 
      Alignment       =   2  'Center
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   20
      MousePointer    =   14  'Arrow and Question
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "����"
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox UserName 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   2760
      MaxLength       =   20
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      ToolTipText     =   "�˺��û���"
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Pass_Title 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�ֶ������� W04"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2295
      TabIndex        =   5
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label User_Title 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�û���"
      BeginProperty Font 
         Name            =   "�ֶ������� W04"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2115
      TabIndex        =   4
      Top             =   1680
      Width           =   540
   End
   Begin VB.Label DMC_Title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DMC CHAT"
      BeginProperty Font 
         Name            =   "�ֶ������� W04"
         Size            =   26.25
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
      Top             =   480
      Width           =   8295
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Dim MainPrivateC As String
Const RHIP As String = "60.176.169.249"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub CancelConn_Click()
Me.CancelConn.Visible = False
Me.LoginSock.Close
Me.UserName.Enabled = True
Me.Password.Enabled = True
Me.LoginNow.Enabled = True
Me.ForgotPass.Enabled = True
Me.SignUP.Enabled = True
End Sub

' Note: "NS" Just is "Me.LoginSock.SendData"
Function NS(Send As String)
On Error Resume Next
   Me.LoginSock.SendData (Send)

End Function



Private Sub Form_Load()
Me.LoginSock.Close
End Sub





Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, 150, LWA_ALPHA
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
End Sub

Private Sub InfCas_Change()
Me.WaitCancel.Enabled = False
End Sub


Private Sub InfCas_Click()
Me.InfoFrame.Visible = False
End Sub

Private Sub InfoFrame_Click()
Me.InfoFrame.Visible = False
Me.InfoFrame.Top = 5360
Me.CancelConn.Visible = False
End Sub

Private Sub LoginNow_Click()
Me.LoginSock.Close
Me.WaitCancel.Enabled = True
If Me.LoginNow.Caption = "ע���!" Then

If Me.UserName.Text <> "" Then
If Me.Password.Text <> "" Then
Dim Choose
Choose = MsgBox("ȷ��ע����Ϊ��" & Me.UserName.Text & " ��ID��", vbInformation + vbYesNo, "ע����ID��")
If Choose = vbYes Then
Me.UserName.Enabled = False
Me.Password.Enabled = False
Me.SignUP.Enabled = False
Me.ForgotPass.Enabled = False
Me.LoginNow.Enabled = False

Me.LoginSock.Close
Me.LoginSock.Connect RHIP, 20859
End If
Else
Me.InfCas.Caption = "��Ϊ����ID����һ������..e"
Me.InfoFrame.Visible = True
Me.InfoFrame.Top = 5360
Me.ShowXD.Enabled = True
End If
Else
Me.InfCas.Caption = "�û����Ǳ���ģ��⽫�Ƕ�λ��ע��ID��Ψһ;��..."
Me.InfoFrame.Visible = True
Me.InfoFrame.Top = 5360
Me.ShowXD.Enabled = True
End If
End If



If Me.LoginNow.Caption = "��¼" Then
If Me.UserName.Text <> "" Then
If Me.Password.Text <> "" Then
Me.LoginSock.Close
Me.LoginSock.Connect RHIP, 20859
Me.UserName.Enabled = False
Me.Password.Enabled = False
Me.SignUP.Enabled = False
Me.ForgotPass.Enabled = False
Me.LoginNow.Enabled = False


Else
Me.InfCas.Caption = "����������.."
Me.InfoFrame.Visible = True
Me.InfoFrame.Top = 5360
Me.ShowXD.Enabled = True
End If
Else
Me.InfCas.Caption = "�������û���.."
Me.InfoFrame.Visible = True
Me.InfoFrame.Top = 5360
Me.ShowXD.Enabled = True
End If
End If


End Sub

Private Sub LoginSock_Close()
Me.LoginSock.Close
End Sub

Private Sub LoginSock_Connect()

Dim TempMD5
Me.CancelConn.Visible = False
If Me.LoginNow.Caption = "��¼" Then

TempMD5 = DigestStrToHexStr(Me.Password.Text)

NS ("%G" & Me.UserName.Text & Space(20 - Len(Me.UserName.Text)) & TempMD5)
End If
If Me.LoginNow.Caption = "ע���!" Then


TempMD5 = DigestStrToHexStr(Me.Password.Text)
Me.LoginSock.SendData "#SU" & TempMD5

Me.DMC_Title.Caption = "��ȴ�ע����ɣ�"
Me.WaitTimer.Enabled = True
End If
End Sub


Private Sub LoginSock_DataArrival(ByVal bytesTotal As Long)
Me.WaitCancel.Enabled = False
Me.CancelConn.Visible = False
Dim GetThings As String
'----------------------------------------------------------------
Me.LoginSock.GetData GetThings '������
'----------------------------------------------------------------
If GetThings = "Copy.Done" Then

Me.UserName.Enabled = True
Me.Password.Enabled = True
Me.LoginNow.Enabled = True
Me.ForgotPass.Enabled = True
Me.SignUP.Enabled = True
NS ("Over")
Sleep 200
Me.InfCas.Caption = "ע��ɹ���"
Me.InfoFrame.Visible = True
SignUP_Click
End If
If Mid(GetThings, 1, 6) = "DoneTo" Then
MainPrivateC = Mid(GetThings, 7, 12)
Me.UserName.Enabled = True
Me.Password.Enabled = True
Me.LoginNow.Enabled = True
Me.ForgotPass.Enabled = True
Me.SignUP.Enabled = True
NS ("Over")
Me.InfCas.Caption = "��¼�ɹ���������ת�����棡"
Me.InfoFrame.Visible = True
Me.WaitUser.Enabled = True
'=========================================================================================
End If
If Mid(GetThings, 1, 4) = "Ewro" Then

Me.UserName.Enabled = True
Me.Password.Enabled = True
Me.LoginNow.Enabled = True
Me.ForgotPass.Enabled = True
Me.SignUP.Enabled = True
NS ("Over")
Sleep 200
Me.InfCas.Caption = "�����Ѿ�����" & Mid(GetThings, 5, 6) & "���ˣ��������ᵼ������ID����ʱ����.."
Me.InfoFrame.Visible = True
Me.InfoFrame.Top = 5360
Me.ShowXD.Enabled = True
End If
If Mid(GetThings, 1, 4) = "Lock" Then
Me.UserName.Enabled = True
Me.Password.Enabled = True
Me.LoginNow.Enabled = True
Me.ForgotPass.Enabled = True
Me.SignUP.Enabled = True
NS ("Over")
Sleep 300
Dim CSSC
If Len(Mid(GetThings, 19, 22)) = "3" Then
CSSC = MsgBox("�����Ѿ�����" & Mid(GetThings, 19, 22) & "�Σ�����ǰʹ�õ�ID�� " & Format(Mid(GetThings, 5, 10), "yyyy��mm��dd��") & " ǰ�������޷���¼������������������ϵ��������", vbExclamation + vbRetryCancel, "�ܾ�����..")
Else
CSSC = vbCancel
End If
If CSSC = vbRetry Then
LoginNow_Click
End If
End If
If GetThings = "CannotSU" Then
Me.UserName.Enabled = True
Me.Password.Enabled = True
Me.LoginNow.Enabled = True
Me.ForgotPass.Enabled = True
Me.SignUP.Enabled = True
NS ("Over")
Sleep 200
MsgBox ("���������ڲ�����ע����Ϊ�������ر���Ҫ������ϵ����������Ա��"), vbExclamation, "�ܾ�����.."
End If
If GetThings = "CannotLn" Then
Me.UserName.Enabled = True
Me.Password.Enabled = True
Me.LoginNow.Enabled = True
Me.ForgotPass.Enabled = True
Me.SignUP.Enabled = True
NS ("Over")
Sleep 200
MsgBox ("���������ڲ�������¼��Ϊ�������ر���Ҫ������ϵ����������Ա��"), vbExclamation, "�ܾ�����.."
End If
End Sub

Private Sub ShowXD_Timer()
If Me.InfoFrame.Top > 360 Then
Me.InfoFrame.Top = Me.InfoFrame.Top - 1000
Else
Me.ShowXD.Enabled = False
End If
End Sub

Private Sub SHPass_Click()
If Me.Password.PasswordChar = "*" Then
Me.Password.PasswordChar = ""
Me.SHPass.Caption = "-"
Else
Me.Password.PasswordChar = "*"
Me.SHPass.Caption = "+"
End If
End Sub

Private Sub SignUP_Click()
If Me.LoginNow.Caption <> "ע���!" Then
Me.ForgotPass.Visible = False
Me.LoginNow.Caption = "ע���!"
Me.SignUP.Caption = "���ص�¼"
Me.DMC_Title.Caption = "ע��һ���û���"
Me.User_Title.Caption = "�µ��û�"
Me.Pass_Title.Caption = "�趨����"
Else
Me.ForgotPass.Visible = True
Me.LoginNow.Caption = "��¼"
Me.SignUP.Caption = "ע���˺�"
Me.DMC_Title.Caption = "DMC CHAT"
Me.User_Title.Caption = "�û���"
Me.Pass_Title.Caption = "����"
End If
End Sub

Private Sub WaitCancel_Timer()
Me.WaitCancel.Enabled = False
Me.CancelConn.Visible = True
End Sub

Private Sub WaitTimer_Timer()
Me.WaitTimer.Enabled = False
NS ("#KC" & Me.UserName)
Me.DMC_Title.Caption = "DMC CHAT"
End Sub



Private Sub WaitUser_Timer()
Me.WaitUser.Enabled = False
MainChat.Show
MainChat.MainSck.RemoteHost = RHIP
MainChat.MainSck.RemotePort = MainPrivateC
MainChat.UserFram.Caption = Me.UserName.Text
Sleep (100)
Unload Me
End Sub
