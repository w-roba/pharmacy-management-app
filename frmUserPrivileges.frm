VERSION 5.00
Begin VB.Form frmUserPrivileges 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Privileges"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10920
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
   ScaleHeight     =   6030
   ScaleWidth      =   10920
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      Height          =   735
      Left            =   4080
      TabIndex        =   18
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Frame FrameRolePermissions 
      Caption         =   "RolePermissions"
      Height          =   4575
      Left            =   6240
      TabIndex        =   11
      Top             =   360
      Width           =   4215
      Begin VB.CommandButton cmdassignpermissions 
         Caption         =   "Assign Permissions"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   3720
         Width           =   1335
      End
      Begin VB.ListBox LstAssignedPermissions 
         Height          =   2085
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtRolePermissions 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   14
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveRolePermissions 
         Caption         =   "Remove Role Permissions"
         Height          =   495
         Left            =   2640
         TabIndex        =   13
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Role Permissions"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame FramePermissions 
      Caption         =   "Permissions"
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   6015
      Begin VB.ListBox LstPermissions 
         Height          =   1860
         Left            =   3840
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmddeletepermissions 
         Caption         =   "Delete Permissions"
         Height          =   495
         Left            =   2160
         TabIndex        =   9
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtPermissions 
         Height          =   495
         Left            =   1800
         MaxLength       =   25
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Cmdaddpermissions 
         Caption         =   "Add Permissions"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Permissions"
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame FrameRoles 
      Caption         =   "Roles"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.ListBox LstRoles 
         Height          =   1860
         Left            =   3840
         TabIndex        =   16
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmddeleterole 
         Caption         =   "Delete Role"
         Height          =   495
         Left            =   2160
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdaddrole 
         Caption         =   "Add Role"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtRoleName 
         Height          =   495
         Left            =   1920
         MaxLength       =   15
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Role"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmUserPrivileges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dbconnect
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
LoadRoles
LoadPermissions
End Sub

Private Function GetRoleID(roleName As String) As Long
Dim sql As String
Dim rs As ADODB.recordset
sql = "SELECT RoleID FROM Roles WHERE RoleName = '" & roleName & "'"
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
GetRoleID = rs.Fields("RoleID").Value
Else
GetRoleID = -1
End If

rs.Close
Set rs = Nothing
End Function
Private Sub cmdAddRole_Click()
Dim sql As String
Dim roleName As String
Dim newRoleID As Double

roleName = txtRoleName.Text

If roleName = "" Then
MsgBox "Please enter a role name.", vbExclamation
Exit Sub
End If
sql = "SELECT RoleID FROM Roles WHERE RoleName = '" & roleName & "'"
Dim rs As ADODB.recordset
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
MsgBox "This role already exists.", vbExclamation
rs.Close
Set rs = Nothing
Exit Sub
End If
rs.Close
Set rs = Nothing

sql = "SELECT MAX(RoleID) AS MaxRoleID FROM Roles"
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
If IsNull(rs.Fields("MaxRoleID").Value) Then
  newRoleID = 1
Else
  newRoleID = rs.Fields("MaxRoleID").Value + 1
End If
Else
newRoleID = 1
End If
rs.Close
Set rs = Nothing

sql = "INSERT INTO Roles (RoleID, RoleName) VALUES (" & newRoleID & ", '" & roleName & "')"
dbconnection.Execute sql

MsgBox "Role added successfully!", vbInformation
txtRoleName.Text = ""
LoadRoles

Exit Sub
End Sub
Private Sub cmdDeleteRole_Click()
dbconnect
Dim sql As String
Dim selectedRole As String

If LstRoles.ListIndex = -1 Then
MsgBox "Please select a role to delete.", vbExclamation
Exit Sub
End If

selectedRole = LstRoles.Text

If MsgBox("Are you sure you want to delete the role '" & selectedRole & "'?", vbYesNo + vbQuestion, "Confirm Deletion") = vbNo Then
Exit Sub
End If

If dbconnection Is Nothing Then
dbconnect
End If

If dbconnection.State = adStateOpen Then
sql = "DELETE FROM Roles WHERE RoleName = '" & selectedRole & "'"
      dbconnection.Execute sql

MsgBox "Role '" & selectedRole & "' deleted successfully!", vbInformation
LoadRoles
Else
MsgBox "Database connection is not open.", vbCritical
End If
LoadRoles

Exit Sub
End Sub
Private Sub LoadRoles()
Dim rs As ADODB.recordset
Dim sql As String
dbconnect

sql = "SELECT RoleID, RoleName FROM Roles"
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

LstRoles.clear

Do While Not rs.EOF
LstRoles.AddItem rs.Fields("RoleID").Value & " - " & rs.Fields("RoleName").Value
rs.MoveNext
Loop

rs.Close
Set rs = Nothing
End Sub
Private Sub LstRoles_Click()
Dim roleID As Long
Dim selectedRole As String

If LstRoles.ListIndex <> -1 Then
selectedRole = LstRoles.Text
roleID = CLng(Split(selectedRole, " - ")(0))
LoadRolePermissions roleID
End If
Exit Sub
End Sub
Private Sub LoadPermissions()
Dim rs As ADODB.recordset
Dim sql As String
dbconnect

sql = "SELECT permissionsid, permissionname FROM permissions"
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

LstPermissions.clear
Do While Not rs.EOF
Dim permissionName As String
Dim permissionID As Variant

permissionName = IIf(IsNull(rs.Fields("permissionname").Value), "", rs.Fields("permissionname").Value)
permissionID = IIf(IsNull(rs.Fields("permissionsid").Value), -1, rs.Fields("permissionsid").Value)

If permissionName <> "" Then
  LstPermissions.AddItem permissionName
  LstPermissions.ItemData(LstPermissions.NewIndex) = permissionID
End If

rs.MoveNext
Loop
End Sub

Private Sub cmdAddPermissions_Click()
dbconnect
Dim sql As String
Dim permissionName As String
Dim newPermissionID As Double

permissionName = txtPermissions.Text

If permissionName = "" Then
MsgBox "Please enter a permission name.", vbExclamation
Exit Sub
End If
sql = "SELECT PermissionsID FROM Permissions WHERE PermissionName = '" & permissionName & "'"
Dim rs As ADODB.recordset
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
MsgBox "This permission already exists.", vbExclamation
rs.Close
Set rs = Nothing
Exit Sub
End If
rs.Close
Set rs = Nothing
sql = "SELECT MAX(PermissionsID) AS MaxPermissionsID FROM Permissions"
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
If IsNull(rs.Fields("MaxPermissionsID").Value) Then
  newPermissionsID = 1
Else
  newPermissionID = rs.Fields("MaxPermissionsID").Value + 1
End If
Else
newPermissionsID = 1
End If
rs.Close
Set rs = Nothing
sql = "INSERT INTO Permissions (PermissionsID, PermissionName) VALUES (" & newPermissionsID & ", '" & permissionName & "')"
dbconnection.Execute sql

MsgBox "Permission added successfully!", vbInformation
LoadPermissions
txtPermissions.Text = ""

Exit Sub
End Sub
Private Sub cmdDeletepermissions_Click()
Dim sql As String
Dim selectedPermission As String
Dim rs As ADODB.recordset
dbconnect
If LstPermissions.ListIndex = -1 Then
MsgBox "Please select a permission to delete.", vbExclamation
Exit Sub
End If

selectedPermission = LstPermissions.Text
If MsgBox("Are you sure you want to delete the permission '" & selectedPermission & "'?", vbYesNo + vbQuestion, "Confirm Deletion") = vbNo Then
Exit Sub
End If

sql = "SELECT PermissionsID FROM Permissions WHERE PermissionName = '" & selectedPermission & "'"
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If rs.EOF Then
MsgBox "Permission '" & selectedPermission & "' does not exist.", vbExclamation
rs.Close
Set rs = Nothing
Exit Sub
End If

rs.Close
Set rs = Nothing
sql = "DELETE FROM Permissions WHERE PermissionName = '" & selectedPermission & "'"
dbconnection.Execute sql

MsgBox "Permission '" & selectedPermission & "' deleted successfully!", vbInformation
LoadPermissions
Exit Sub
End Sub
Private Sub cmdAssignPermissions_Click()
Dim sql As String
Dim selectedRole As String
Dim selectedPermission As String
Dim roleID As Long
Dim permissionsID As Long
Dim newRolePermissionsID As Long
If LstRoles.ListIndex = -1 Or LstPermissions.ListIndex = -1 Then
MsgBox "Please select a role and a permission to assign.", vbExclamation
Exit Sub
End If

selectedRole = LstRoles.Text
selectedPermission = LstPermissions.Text
sql = "SELECT RoleID FROM Roles WHERE RoleName = '" & selectedRole & "'"
Dim rs As ADODB.recordset
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
roleID = rs.Fields("RoleID").Value
Else
MsgBox "Role not found.", vbExclamation
rs.Close
Set rs = Nothing
Exit Sub
End If
rs.Close
Set rs = Nothing
sql = "SELECT PermissionsID FROM Permissions WHERE PermissionName = '" & selectedPermission & "'"
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
permissionsID = rs.Fields("PermissionsID").Value
Else
MsgBox "Permission not found.", vbExclamation
rs.Close
Set rs = Nothing
Exit Sub
End If
rs.Close
Set rs = Nothing
sql = "SELECT * FROM RolePermissions WHERE RoleID = " & roleID & " AND PermissionsID = " & permissionsID
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
MsgBox "This permission is already assigned to the selected role.", vbExclamation
rs.Close
Set rs = Nothing
Exit Sub
End If
rs.Close
Set rs = Nothing
sql = "SELECT MAX(RolePermissionsID) AS MaxRolePermissionsID FROM RolePermissions"
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
If IsNull(rs.Fields("MaxRolePermissionsID").Value) Then
  newRolePermissionsID = 1
Else
  newRolePermissionsID = rs.Fields("MaxRolePermissionsID").Value + 1
End If
Else
newRolePermissionsID = 1
End If
rs.Close
Set rs = Nothing
sql = "INSERT INTO RolePermissions (RolePermissionsID, RoleID, PermissionsID) VALUES (" & newRolePermissionsID & ", " & roleID & ", " & permissionsID & ")"
dbconnection.Execute sql

MsgBox "Permission '" & selectedPermission & "' assigned to role '" & selectedRole & "' successfully!", vbInformation
LoadRolePermissions roleID

Exit Sub
End Sub

Private Sub LoadRolePermissions(roleID As Long)
Dim rs As ADODB.recordset
Dim sql As String
dbconnect
sql = "SELECT p.PermissionName FROM RolePermissions rp " & _
"INNER JOIN Permissions p ON rp.PermissionsID = p.PermissionsID " & _
"WHERE rp.RoleID = " & roleID

Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

LstAssignedPermissions.clear
While Not rs.EOF
LstAssignedPermissions.AddItem rs.Fields("PermissionName").Value
rs.MoveNext
Wend
rs.Close
Set rs = Nothing
End Sub
Private Sub cmdRemoveRolePermissions_Click()
Dim sql As String
Dim selectedPermission As String
Dim roleID As Long
Dim permissionID As Long
Dim rs As ADODB.recordset
If LstRoles.ListIndex = -1 Then
MsgBox "Please select a role.", vbExclamation
Exit Sub
End If
If LstAssignedPermissions.ListIndex = -1 Then
MsgBox "Please select a permission to remove.", vbExclamation
Exit Sub
End If
selectedPermission = LstAssignedPermissions.Text
roleID = CLng(Split(LstRoles.Text, " - ")(0))
sql = "SELECT PermissionsID FROM Permissions WHERE PermissionName = '" & selectedPermission & "'"
Set rs = New ADODB.recordset
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
permissionID = rs.Fields("PermissionsID").Value
Else
MsgBox "Permission '" & selectedPermission & "' does not exist.", vbExclamation
rs.Close
Set rs = Nothing
Exit Sub
End If
rs.Close
Set rs = Nothing
sql = "DELETE FROM RolePermissions WHERE RoleID = " & roleID & " AND PermissionsID = " & permissionID
dbconnection.Execute sql
MsgBox "Permission '" & selectedPermission & "' removed from role successfully!", vbInformation
LoadRolePermissions roleID
Exit Sub
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub



