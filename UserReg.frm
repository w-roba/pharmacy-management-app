VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User registration"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12900
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
   ScaleHeight     =   6195
   ScaleWidth      =   12900
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   495
      Left            =   2040
      TabIndex        =   16
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton cmddeleteuser 
      Caption         =   "Delete User"
      Height          =   495
      Left            =   3840
      TabIndex        =   13
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdadduser 
      Caption         =   "Add User"
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Frame FrameUserRegistration 
      Caption         =   "UserRegistration"
      Height          =   4455
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   6975
      Begin VB.TextBox txtUserID 
         Height          =   615
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   15
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox cmbRole 
         Height          =   345
         Left            =   3000
         TabIndex        =   11
         Top             =   3840
         Width           =   2655
      End
      Begin VB.TextBox txtUsername 
         Height          =   615
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtPassword 
         Height          =   615
         Left            =   3000
         MaxLength       =   25
         TabIndex        =   5
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtConfirmPassword 
         Height          =   615
         Left            =   3000
         MaxLength       =   25
         TabIndex        =   4
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "UserID"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
         Height          =   615
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Confirm Password"
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Role"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3840
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   4680
      Width           =   1575
   End
   Begin MSComctlLib.ListView UserList 
      Height          =   4215
      Left            =   7320
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000003&
      Caption         =   "Have an Account?"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   5640
      Width           =   1575
   End
End
Attribute VB_Name = "frmUserReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
dbconnect
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
FrameUserRegistration.Enabled = False
cmdRegister.Enabled = False
cmddeleteuser.Enabled = False
txtPassword.PasswordChar = "*"
txtConfirmPassword.PasswordChar = "*"

With UserList
.View = lvwReport
.ColumnHeaders.Add , , "UserID", 1500
.ColumnHeaders.Add , , "Username", 1500
.ColumnHeaders.Add , , "Role", 1500
End With
cmdadduser.Enabled = True

LoadUsers
LoadRoles
End Sub
Private Sub LoadRoles()
Dim rs As ADODB.recordset
Dim sql As String
dbconnect

sql = "SELECT RoleName FROM Roles"
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

cmbRole.clear

Do While Not rs.EOF
cmbRole.AddItem rs.Fields("RoleName").Value
rs.MoveNext
Loop

rs.Close
Set rs = Nothing
dbdisconnect
End Sub
Private Sub cmdAddUser_Click()
FrameUserRegistration.Enabled = True
cmdRegister.Enabled = True
cmdadduser.Enabled = False

End Sub
Private Sub cmdRegister_Click()
Dim conn As ADODB.Connection
Dim rs As ADODB.recordset
Dim sql As String
Dim userID As Integer

If txtUserID.Text = "" Then
MsgBox "Please fill in the UserID.", vbExclamation
Exit Sub
End If

If txtUsername.Text = "" Or txtPassword.Text = "" Or txtConfirmPassword.Text = "" Then
MsgBox "Please fill in all required fields.", vbExclamation
Exit Sub
End If

If txtPassword.Text <> txtConfirmPassword.Text Then
MsgBox "Passwords do not match.", vbExclamation
Exit Sub
End If

If cmbRole.Text = "" Then
MsgBox "Please select a role.", vbExclamation
Exit Sub
End If

userID = CInt(txtUserID.Text)

dbconnect
If IsDuplicate(txtUsername.Text, userID) Then
MsgBox "This username or UserID is already taken. Please choose another one.", vbExclamation
Exit Sub
End If
sql = "SELECT * FROM Userregistration WHERE 1=0"
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenDynamic, adLockOptimistic

rs.AddNew
rs!userID = userID
rs!username = txtUsername.Text
rs!password = HashPassword(txtPassword.Text)
rs!confirmPassword = HashPassword(txtConfirmPassword.Text)
rs!role = cmbRole.Text
rs.Update

MsgBox "Registration successful!.", vbInformation

AddUser userID, txtUsername.Text, cmbRole.Text

txtUserID.Text = ""
txtUsername.Text = ""
txtPassword.Text = ""
txtConfirmPassword.Text = ""
cmbRole.Text = ""
cmdadduser.Enabled = True
cmdRegister.Enabled = False
FrameUserRegistration.Enabled = False

rs.Close
Set rs = Nothing
dbdisconnect
Exit Sub
dbdisconnect
End Sub

Private Function IsDuplicate(username As String, userID As Integer) As Boolean
Dim rs As ADODB.recordset
Set rs = New ADODB.recordset

Dim sql As String
sql = "SELECT Username, UserID FROM Userregistration WHERE Username = '" & username & "' OR UserID = " & userID

rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If rs.EOF Then
IsDuplicate = False
Else
IsDuplicate = True
End If

rs.Close
Set rs = Nothing
End Function
Function HashPassword(password As String) As String
HashPassword = password
End Function
Private Sub LoadUsers()
Dim rs As ADODB.recordset
Dim sql As String
dbconnect

sql = "SELECT UserID, Username, Role FROM Userregistration"
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenForwardOnly, adLockReadOnly
UserList.ListItems.clear
Do While Not rs.EOF
AddUser rs.Fields("UserID").Value, rs.Fields("Username").Value, rs.Fields("Role").Value
rs.MoveNext
Loop
rs.Close
dbdisconnect
End Sub

Private Sub AddUser(userID As Integer, username As String, role As String)
Dim Item As ListItem
Set Item = UserList.ListItems.Add
Item.Text = userID
Item.SubItems(1) = username
Item.SubItems(2) = role
End Sub
Private Sub UserList_ItemClick(ByVal Item As MSComctlLib.ListItem)
txtUserID.Text = Item.Text
txtUsername.Text = Item.SubItems(1)
cmbRole.Text = Item.SubItems(2)
cmddeleteuser.Enabled = True
FrameUserRegistration.Enabled = True
End Sub
Private Sub cmdDeleteUser_Click()
Dim sql As String
Dim selectedItem As MSComctlLib.ListItem
Dim userID As Integer

If UserList.selectedItem Is Nothing Then
MsgBox "Please select a user to delete.", vbExclamation
Exit Sub
End If

Set selectedItem = UserList.selectedItem
userID = CInt(selectedItem.Text)
If MsgBox("Are you sure you want to delete user '" & selectedItem.SubItems(1) & "'?", vbYesNo + vbQuestion, "Confirm Deletion") = vbNo Then
Exit Sub
End If
dbconnect
sql = "DELETE FROM Userregistration WHERE UserID = " & userID

dbconnection.Execute sql
MsgBox "User  '" & selectedItem.SubItems(1) & "' has been deleted successfully.", vbInformation

dbdisconnect

UserList.ListItems.Remove selectedItem.Index
txtUserID.Text = ""
txtUsername.Text = ""
txtPassword.Text = ""
txtConfirmPassword.Text = ""
cmbRole.Text = ""
cmddeleteuser.Enabled = False
FrameUserRegistration.Enabled = False
Exit Sub

dbdisconnect
End Sub
Private Sub cmbRole_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub txtUserID_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyNumbers(KeyAscii, txtUserID)
End Sub
Private Sub cmdPrintReport_Click()
dbconnect

If recordset Is Nothing Then Set recordset = New ADODB.recordset

With recordset
If .State = 1 Then .Close
.Open "SELECT * FROM Userregistration", dbconnection, adOpenDynamic, adLockOptimistic
End With
With DataReportUserreg
Set .DataSource = recordset
.Refresh
.Show
End With

Exit Sub
dbdisconnect
End Sub
Private Sub Command1_Click()
Unload frmUserReg
frmLogin.Show
End Sub

