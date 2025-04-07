VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9225
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9225
   Begin VB.CommandButton Cmdreg 
      Caption         =   "Create Account"
      Height          =   495
      Left            =   7200
      TabIndex        =   11
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdshowpassword 
      Caption         =   "&Show Password"
      Height          =   495
      Left            =   1920
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   10
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "Login"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Frame FrameUserLogin 
      Caption         =   "UserLogin"
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      Begin VB.ComboBox cmbRole 
         Height          =   345
         ItemData        =   "Login.frx":0000
         Left            =   3000
         List            =   "Login.frx":000D
         TabIndex        =   4
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox txtUsername 
         Height          =   615
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   3
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtPassword 
         Height          =   615
         Left            =   3000
         MaxLength       =   25
         TabIndex        =   2
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   615
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Role"
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   2880
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdforgotpassword 
      Caption         =   "Forgot Password"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000003&
      Caption         =   "New User?"
      Height          =   255
      Left            =   7320
      TabIndex        =   12
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public UserLoggedIn As Boolean


Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
txtPassword.PasswordChar = "*"
End Sub
Private Sub cmdLogin_Click()
dbconnect
Dim username As String
Dim password As String
Dim UserRole As String

username = txtUsername.Text
password = txtPassword.Text
If username = "" Or password = "" Then
MsgBox "Please enter both username and password.", vbExclamation
Exit Sub
End If

If CheckCredentials(username, password, UserRole) Then
If cmbRole.Text = UserRole Then
UserLoggedIn = True
Roles = UserRole
MsgBox "Logged in as: " & UserRole, vbInformation
Unload Me
MDIForm1.Mfile.Enabled = True
MDIForm1.Fmvr.Enabled = True
MDIForm1.Fmac.Enabled = True

Select Case UserRole
Case "Admin"
MDIForm1.Mfile.Enabled = True
MDIForm1.Fmvr.Enabled = True
MDIForm1.Fmac.Enabled = True
MDIForm1.Fmmed.Enabled = True
MDIForm1.Fmop.Enabled = True
MDIForm1.Fmpc.Enabled = True
MDIForm1.Fmpr.Enabled = True
MDIForm1.Fmsr.Enabled = True
MDIForm1.fmst.Enabled = True
MDIForm1.Fmsup.Enabled = True
MDIForm1.FmUser.Enabled = True

MDIForm1.rptreg.Enabled = True
MDIForm1.rptO.Enabled = True
MDIForm1.rptsr.Enabled = True
MDIForm1.rptpr.Enabled = True
MDIForm1.rptInv.Enabled = True
MDIForm1.rptpay.Enabled = True
MDIForm1.rptsup.Enabled = True


Case "Pharmacist"
MDIForm1.Fmmed.Enabled = True
MDIForm1.Fmpr.Enabled = True
MDIForm1.fmst.Enabled = True
MDIForm1.Fmop.Enabled = True
MDIForm1.Fmsr.Enabled = True

MDIForm1.Fmup.Enabled = False
MDIForm1.Fmsup.Enabled = False
MDIForm1.FmUser.Enabled = False
MDIForm1.Fmpc.Enabled = False

MDIForm1.rptreg.Enabled = True
MDIForm1.rptO.Enabled = True
MDIForm1.rptsr.Enabled = True
MDIForm1.rptpr.Enabled = True
MDIForm1.rptInv.Enabled = True
MDIForm1.rptsup.Enabled = False
MDIForm1.rptpay.Enabled = False


Case "Cashier"
MDIForm1.Fmpc.Enabled = True
MDIForm1.Fmpr.Enabled = True

MDIForm1.Fmup.Enabled = False
MDIForm1.Fmop.Enabled = False
MDIForm1.Fmmed.Enabled = False
MDIForm1.fmst.Enabled = False
MDIForm1.Fmsr.Enabled = False
MDIForm1.Fmsup.Enabled = False
MDIForm1.FmUser.Enabled = False
MDIForm1.rptpr.Enabled = True
MDIForm1.rptpay.Enabled = True

Case Else
MDIForm1.Fmmed.Enabled = False
MDIForm1.Fmop.Enabled = False
MDIForm1.Fmpc.Enabled = False
MDIForm1.Fmpr.Enabled = False
MDIForm1.Fmsr.Enabled = False
MDIForm1.fmst.Enabled = False
MDIForm1.Fmsup.Enabled = False
MDIForm1.FmUser.Enabled = False
End Select
Else
MsgBox "Invalid Role.", vbExclamation
End If
Else
MsgBox "Invalid username or password.", vbExclamation
End If

Cleanup:
dbdisconnect
Exit Sub
Resume Cleanup
End Sub


Private Function CheckCredentials(username As String, password As String, ByRef UserRole As String) As Boolean
Dim sql As String
Dim cmd As ADODB.Command
Dim recordset As ADODB.recordset

sql = "SELECT Role FROM UserRegistration WHERE StrComp(Username, ?, 0) = 0 AND StrComp(Password, ?, 0) = 0"

Set cmd = New ADODB.Command
With cmd
.ActiveConnection = dbconnection
.CommandText = sql
.CommandType = adCmdText
.Parameters.Append .CreateParameter("Username", adVarChar, , 50, username)
.Parameters.Append .CreateParameter("Password", adVarChar, , 50, password)
End With

Set recordset = New ADODB.recordset
recordset.Open cmd

CheckCredentials = Not recordset.EOF
If CheckCredentials Then UserRole = recordset.Fields("Role").Value

If recordset.State = adStateOpen Then recordset.Close
End Function
Private Sub cmdforgotpassword_Click()
frmForgotPassword.Show
Unload frmLogin
End Sub

Private Sub cmdshowpassword_Click()
If txtPassword.PasswordChar = "*" Then
txtPassword.PasswordChar = ""
cmdshowpassword.Caption = "&hide"
ElseIf cmdshowpassword.Caption = "&hide" Then
txtPassword.PasswordChar = "*"
cmdshowpassword.Caption = "&show"
End If
End Sub
Private Sub Cmdreg_Click()
Unload frmLogin
frmUserReg.Show
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub cmbRole_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


