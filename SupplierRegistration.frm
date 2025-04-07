VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSupplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier Registration"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14640
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
   ScaleHeight     =   7050
   ScaleWidth      =   14640
   Begin VB.CommandButton cmdPrintReport 
      Caption         =   "Generate Report"
      Height          =   495
      Left            =   9360
      TabIndex        =   17
      Top             =   5760
      Width           =   1335
   End
   Begin MSComctlLib.ListView SupplierList 
      Height          =   5175
      Left            =   9120
      TabIndex        =   16
      Top             =   360
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   9128
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6000
      TabIndex        =   15
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   7680
      TabIndex        =   14
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4320
      TabIndex        =   13
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdaddsupplier 
      Caption         =   "Add Supplier"
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Frame FrameSupplierDetails 
      Caption         =   "Supplier Details"
      Height          =   5295
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   8055
      Begin VB.TextBox txtEmail 
         Height          =   615
         Left            =   3000
         MaxLength       =   25
         TabIndex        =   10
         Top             =   4200
         Width           =   2655
      End
      Begin VB.TextBox txtPhone 
         Height          =   615
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   9
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox txtAddress 
         Height          =   615
         Left            =   3000
         MaxLength       =   25
         TabIndex        =   8
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox txtSupplierName 
         Height          =   615
         Left            =   3000
         MaxLength       =   25
         TabIndex        =   7
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtSupplierID 
         Height          =   615
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Email"
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Phone No."
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Address"
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Supplier Name"
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Supplier ID"
         Height          =   615
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
dbconnect
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
With SupplierList
.View = lvwReport
.ColumnHeaders.Add , , "Supplier ID", 1000
.ColumnHeaders.Add , , "Supplier Name"
.ColumnHeaders.Add , , "Address"
.ColumnHeaders.Add , , "Phone"
.ColumnHeaders.Add , , "Email"
End With

LoadRecordsIntoListView

txtSupplierID.Enabled = False
txtSupplierName.Enabled = False
txtAddress.Enabled = False
txtPhone.Enabled = False
txtEmail.Enabled = False
cmdsave.Enabled = False
cmddelete.Enabled = False
cmdcancel.Enabled = False
cmdexit.Enabled = True
FrameSupplierDetails.Enabled = False
cmdaddsupplier.Enabled = True
End Sub

Private Sub LoadRecordsIntoListView()
Set recordset = New ADODB.recordset
recordset.Open "SELECT * FROM SupplierTable", dbconnection, adOpenStatic, adLockReadOnly

SupplierList.ListItems.clear

Do While Not recordset.EOF
Set Item = SupplierList.ListItems.Add(, , recordset.Fields("SupplierID").Value)
Item.SubItems(1) = recordset.Fields("SupplierName").Value
Item.SubItems(2) = recordset.Fields("Address").Value
Item.SubItems(3) = recordset.Fields("Phone").Value
Item.SubItems(4) = recordset.Fields("Email").Value
recordset.MoveNext
Loop
recordset.Close
Set recordset = Nothing
End Sub



Private Sub SupplierList_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Not Item Is Nothing Then
txtSupplierID.Text = Item.Text
txtSupplierName.Text = Item.SubItems(1)
txtAddress.Text = Item.SubItems(2)
txtPhone.Text = Item.SubItems(3)
txtEmail.Text = Item.SubItems(4)

txtSupplierID.Enabled = True
txtSupplierName.Enabled = True
txtAddress.Enabled = True
txtPhone.Enabled = True
txtEmail.Enabled = True

cmdaddsupplier.Enabled = False
cmdsave.Enabled = False
cmddelete.Enabled = True
cmdcancel.Enabled = True
cmdexit.Enabled = True
FrameSupplierDetails.Enabled = False
End If
End Sub

Private Sub cmdaddsupplier_Click()
txtSupplierID.Enabled = True
txtSupplierName.Enabled = True
txtAddress.Enabled = True
txtPhone.Enabled = True
txtEmail.Enabled = True
cmdsave.Enabled = True
cmddelete.Enabled = False
cmdexit.Enabled = True
FrameSupplierDetails.Enabled = True
cmdaddsupplier.Enabled = False

txtSupplierID.Text = ""
txtSupplierName = ""
txtAddress.Text = ""
txtPhone.Text = ""
txtEmail.Text = ""

txtSupplierID.SetFocus
Exit Sub
End Sub
Private Sub txtSupplierID_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyNumbers(KeyAscii, txtSupplierID)
End Sub
Private Sub txtSupplierName_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyText(KeyAscii)
End Sub
Private Sub txtAddress_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyAddress(KeyAscii, txtAddress.Text)
End Sub
Private Sub txtPhone_KeyPress(KeyAscii As Integer)
If Len(txtPhone.Text) = 0 And KeyAscii = Asc(" ") Then
KeyAscii = 0
End If
KeyAscii = AllowOnlyNumbersAndSinglePlus(KeyAscii, txtPhone.Text)
End Sub
Private Sub txtEmail_KeyPress(KeyAscii As Integer)
Dim currentEmail As String
currentEmail = txtEmail.Text
KeyAscii = AllowOnlyEmail(KeyAscii, currentEmail)
End Sub
Private Sub cmdSave_Click()
cmdaddsupplier.Enabled = False
cmddelete.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
cmdexit.Enabled = True
FrameSupplierDetails.Enabled = True

If Len(Trim(txtSupplierID.Text)) = 0 Then
MsgBox "Supplier ID is required. Please enter it."
txtSupplierID.SetFocus
Exit Sub
End If
Set recordsetCheck = New ADODB.recordset
recordsetCheck.Open "SELECT 1 FROM SupplierTable WHERE SupplierID = " & CDbl(txtSupplierID.Text), dbconnection, adOpenStatic, adLockReadOnly

If Not recordsetCheck.BOF And Not recordsetCheck.EOF Then
MsgBox "Supplier ID already exists. Please enter a unique Supplier ID."
txtSupplierID.SetFocus
recordsetCheck.Close
Set recordsetCheck = Nothing
Exit Sub
End If

If Len(Trim(txtSupplierName.Text)) = 0 Then
MsgBox "The Supplier Name is required. Please enter it."
txtSupplierName.SetFocus
Exit Sub
End If

If Len(Trim(txtAddress.Text)) = 0 Then
MsgBox "The Address is required. Please enter it."
txtAddress.SetFocus
Exit Sub
End If

If Len(Trim(txtPhone.Text)) = 0 Then
MsgBox "The Phone Number is required. Please enter it."
txtPhone.SetFocus
Exit Sub
End If

If Len(Trim(txtEmail.Text)) = 0 Then
MsgBox "The Email is required. Please enter it."
txtEmail.SetFocus
Exit Sub
End If
Dim email As String
email = txtEmail.Text

If InStr(email, "@") = 0 Then
MsgBox "Please enter a valid email address that contains an '@' symbol.", vbExclamation
Exit Sub
End If

Dim atPos As Integer
atPos = InStr(email, "@")

If InStr(atPos + 1, email, ".") = 0 Then
MsgBox "Please enter a valid email address that contains at least one '.' after the '@' symbol.", vbExclamation
Exit Sub
End If

MsgBox "Email saved: " & email, vbInformation


Set recordset = New ADODB.recordset
recordset.Open "SELECT * FROM SupplierTable WHERE SupplierID = " & CDbl(txtSupplierID.Text), dbconnection, adOpenKeyset, adLockOptimistic

recordset.AddNew
With recordset
.Fields("SupplierID").Value = CDbl(txtSupplierID.Text)
.Fields("SupplierName").Value = txtSupplierName.Text
.Fields("Address").Value = txtAddress.Text
.Fields("Phone").Value = CDbl(txtPhone.Text)
.Fields("Email").Value = txtEmail.Text
.Update
End With
MsgBox "The Supplier has been added Successfully"

LoadRecordsIntoListView

cmdaddsupplier.Enabled = True
cmdsave.Enabled = False
Exit Sub
End Sub
Private Sub cmdDelete_Click()
If SupplierList.selectedItem Is Nothing Then
MsgBox "Please select a supplier to delete.", vbExclamation
Exit Sub
End If

If MsgBox("Are you sure you want to delete the selected supplier?", vbYesNo + vbQuestion, "Confirm Deletion") = vbNo Then
Exit Sub
End If

Dim SupplierID As Double
SupplierID = CDbl(SupplierList.selectedItem.Text)

dbconnect
dbconnection.Execute "DELETE FROM SupplierTable WHERE SupplierID = " & SupplierID

SupplierList.ListItems.Remove SupplierList.selectedItem.Index

ClearSupplierDetails
MsgBox "Supplier deleted successfully!"

dbdisconnect
Exit Sub
dbdisconnect
End Sub
Private Sub ClearSupplierDetails()
txtSupplierID.Text = ""
txtSupplierName.Text = ""
txtAddress.Text = ""
txtPhone.Text = ""
txtEmail.Text = ""

txtSupplierID.Enabled = True
txtSupplierName.Enabled = False
txtAddress.Enabled = False
txtPhone.Enabled = False
txtEmail.Enabled = False

cmdaddsupplier.Enabled = True
cmdsave.Enabled = False
cmddelete.Enabled = False
cmdexit.Enabled = True
cmdcancel.Enabled = False
FrameSupplierDetails.Enabled = False
End Sub
Private Sub cmdcancel_Click()
txtSupplierID.Text = ""
txtSupplierName = ""
txtAddress.Text = ""
txtPhone.Text = ""
txtEmail.Text = ""
End Sub
Private Sub cmdexit_Click()
If MsgBox("Are you sure you want to exit?", vbYesNo, "Confirm Exit") = vbYes Then
dbdisconnect
Unload Me
End If
End Sub

Private Sub cmdPrintReport_Click()
dbconnect

If recordset Is Nothing Then Set recordset = New ADODB.recordset

With recordset
If .State = 1 Then .Close
.Open "SELECT * FROM SupplierTable", dbconnection, adOpenDynamic, adLockOptimistic
End With
With DataReportSuppliers
Set .DataSource = recordset
.Refresh
.Show
End With

Exit Sub
dbdisconnect
End Sub



