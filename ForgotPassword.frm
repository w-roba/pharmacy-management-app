VERSION 5.00
Begin VB.Form frmForgotPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forgot Password"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7095
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
   ScaleHeight     =   4860
   ScaleWidth      =   7095
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdresetpassword 
      Caption         =   "Reset Password"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Frame FrameResetPassword 
      Caption         =   "Password Reset"
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txtNewPassword 
         Height          =   615
         Left            =   3000
         MaxLength       =   25
         TabIndex        =   4
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txtUsername 
         Height          =   615
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   3
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtConfirmNewPassword 
         Height          =   615
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   2
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "New Password"
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
         Height          =   615
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Confirm New Password"
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   2640
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   3840
      Width           =   1815
   End
End
Attribute VB_Name = "frmForgotPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload frmForgotPassword
frmLogin.Show
End Sub

Private Sub Form_Load()
dbconnect
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
txtNewPassword.PasswordChar = "*"
txtConfirmNewPassword.PasswordChar = "*"
Command1.Enabled = False

End Sub
Private Sub cmdResetPassword_Click()

Dim rs As ADODB.recordset
Dim sql As String
Dim username As String
Dim newPassword As String
Dim confirmPassword As String

username = txtUsername.Text
newPassword = txtNewPassword.Text
confirmPassword = txtConfirmNewPassword.Text

If username = "" Or newPassword = "" Or confirmPassword = "" Then
MsgBox "Please fill in all required fields.", vbExclamation
Exit Sub
End If

If newPassword <> confirmPassword Then
MsgBox "New Password and Confirm Password do not match.", vbExclamation
Exit Sub
End If
dbconnect
sql = "SELECT * FROM Userregistration WHERE Username = '" & username & "'"

Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
sql = "UPDATE Userregistration SET [Password] = '" & newPassword & "', [ConfirmPassword] = '" & confirmPassword & "' WHERE [Username] = '" & username & "'"
dbconnection.Execute sql
MsgBox "Password reset successfully!", vbInformation
cmdreset.enable = False
Else
MsgBox "Username not found.", vbExclamation
End If
rs.Close
Set rs = Nothing
dbdisconnect
Exit Sub
dbdisconnect
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub

