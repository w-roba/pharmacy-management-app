VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMedicine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medicine Registration"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14970
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
   ScaleHeight     =   8055
   ScaleWidth      =   14970
   Begin VB.ComboBox Cmbcat 
      Height          =   345
      ItemData        =   "MedicineRegistration.frx":0000
      Left            =   13080
      List            =   "MedicineRegistration.frx":0022
      TabIndex        =   24
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrintCategory 
      Caption         =   "Filter Category"
      Height          =   615
      Left            =   13080
      TabIndex        =   23
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrintInventory 
      Caption         =   "View Inventory"
      Height          =   495
      Left            =   13080
      TabIndex        =   22
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrintReport 
      Caption         =   "Generate Report"
      Height          =   495
      Left            =   13080
      TabIndex        =   21
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6840
      TabIndex        =   20
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   3000
      TabIndex        =   19
      Top             =   7320
      Width           =   1455
   End
   Begin MSComctlLib.ListView MedicineList 
      Height          =   6975
      Left            =   7440
      TabIndex        =   18
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   12303
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdaddnewitem 
      Caption         =   "Add New Item"
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox txtMedicineID 
      Height          =   615
      Left            =   4080
      MaxLength       =   15
      TabIndex        =   7
      Top             =   600
      Width           =   2655
   End
   Begin VB.ComboBox cmbMedicineType 
      Height          =   345
      ItemData        =   "MedicineRegistration.frx":00A6
      Left            =   4080
      List            =   "MedicineRegistration.frx":00BF
      TabIndex        =   6
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8760
      TabIndex        =   5
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   7320
      Width           =   1455
   End
   Begin VB.TextBox txtCostPerUnit 
      Height          =   615
      Left            =   4080
      MaxLength       =   5
      TabIndex        =   3
      Top             =   5040
      Width           =   2655
   End
   Begin VB.TextBox txtReorderLevel 
      Height          =   615
      Left            =   4080
      MaxLength       =   5
      TabIndex        =   2
      Top             =   6000
      Width           =   2655
   End
   Begin VB.TextBox txtStockLevel 
      Height          =   615
      Left            =   4080
      MaxLength       =   5
      TabIndex        =   1
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox txtMedicineName 
      Height          =   615
      Left            =   4080
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Frame FrameMedicineDetails 
      Caption         =   "Medicine Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   1200
      TabIndex        =   9
      Top             =   0
      Width           =   6255
      Begin VB.ComboBox cmbMedicineCategory 
         Height          =   345
         ItemData        =   "MedicineRegistration.frx":0105
         Left            =   2880
         List            =   "MedicineRegistration.frx":0127
         TabIndex        =   17
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "Medicine Category"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Reorder Level"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Cost Per Unit"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Stock Level"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Medicine Type"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Medicine Name"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Medicine ID"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmMedicine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dbconnect
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
With MedicineList
.View = lvwReport
.ColumnHeaders.Add , , "Medicine ID", 800
.ColumnHeaders.Add , , "Medicine Name", 1000
.ColumnHeaders.Add , , "Category", 1000
.ColumnHeaders.Add , , "Type", 1000
.ColumnHeaders.Add , , "Stock Level", 1000
.ColumnHeaders.Add , , "Cost Per Unit", 1000
.ColumnHeaders.Add , , "Reorder Level", 1000
End With

LoadRecordsIntoListView

txtMedicineID.Enabled = False
txtMedicineName.Enabled = False
cmbMedicineCategory.Enabled = False
cmbMedicineType.Enabled = False
txtStockLevel.Enabled = False
txtCostPerUnit.Enabled = False
txtReorderLevel.Enabled = False
cmdaddnewitem.Enabled = True
cmdedit.Enabled = False
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdexit.Enabled = True
FrameMedicineDetails.Enabled = False
End Sub

Private Sub LoadRecordsIntoListView()
Set recordset = New ADODB.recordset
recordset.Open "SELECT * FROM MedicineTable", dbconnection, adOpenStatic, adLockReadOnly

MedicineList.ListItems.clear

Do While Not recordset.EOF
Dim Item As ListItem
Set Item = MedicineList.ListItems.Add(, , recordset.Fields("MedicineID").Value)
Item.SubItems(1) = recordset.Fields("MedicineName").Value
Item.SubItems(2) = recordset.Fields("MedicineCategory").Value
Item.SubItems(3) = recordset.Fields("MedicineType").Value
Item.SubItems(4) = recordset.Fields("StockLevel").Value
Item.SubItems(5) = recordset.Fields("CostPerUnit").Value
Item.SubItems(6) = recordset.Fields("ReorderLevel").Value
recordset.MoveNext
Loop

recordset.Close
Set recordset = Nothing
End Sub

Private Sub MedicineList_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Not Item Is Nothing Then
txtMedicineID.Text = Item.Text
txtMedicineName.Text = Item.SubItems(1)
cmbMedicineCategory.Text = Item.SubItems(2)
cmbMedicineType.Text = Item.SubItems(3)
txtStockLevel.Text = Item.SubItems(4)
txtCostPerUnit.Text = Item.SubItems(5)
txtReorderLevel.Text = Item.SubItems(6)
txtMedicineID.Enabled = False
txtMedicineName.Enabled = True
cmbMedicineCategory.Enabled = True
cmbMedicineType.Enabled = True
txtStockLevel.Enabled = True
txtCostPerUnit.Enabled = True
txtReorderLevel.Enabled = True
cmdaddnewitem.Enabled = False
cmdsave.Enabled = False
cmdedit.Enabled = True
cmdcancel.Enabled = True
End If
End Sub
Private Sub cmdAddNewItem_Click()
txtMedicineID.Enabled = True
txtMedicineName.Enabled = True
cmbMedicineCategory.Enabled = True
cmbMedicineType.Enabled = True
txtStockLevel.Enabled = True
txtCostPerUnit.Enabled = True
txtReorderLevel.Enabled = True
cmdedit.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
cmdexit.Enabled = True

FrameMedicineDetails.Enabled = True

cmdaddnewitem.Enabled = False

txtMedicineID.Text = ""
txtMedicineName.Text = ""
cmbMedicineCategory.Text = ""
cmbMedicineType.Text = ""
txtStockLevel.Text = ""
txtCostPerUnit.Text = ""
txtReorderLevel.Text = ""

txtMedicineID.SetFocus
Exit Sub
End Sub
Private Sub txtMedicineID_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyNumbers(KeyAscii, txtMedicineID)
End Sub
Private Sub txtMedicineName_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyText(KeyAscii)
End Sub
Private Sub txtStockLevel_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyNumbers(KeyAscii, txtStockLevel)
End Sub
Private Sub txtCostPerUnit_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyNumbersForCostPerUnit(KeyAscii, txtCostPerUnit)
End Sub
Private Sub txtReorderLevel_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyNumbers(KeyAscii, txtReorderLevel)
End Sub
Private Sub cmbMedicineCategory_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub cmbCat_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub cmbMedicineType_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdSave_Click()
cmdaddnewitem.Enabled = True
cmdedit.Enabled = False
cmdcancel.Enabled = True
cmdsave.Enabled = True
cmdexit.Enabled = True
FrameMedicineDetails.Enabled = True

If Len(Trim(txtMedicineID.Text)) = 0 Then
MsgBox "Medicine ID is required. Please enter it.", vbExclamation
txtMedicineID.SetFocus
Exit Sub
End If

Set recordsetCheck = New ADODB.recordset
recordsetCheck.Open "SELECT 1 FROM MedicineTable WHERE MedicineID = " & CDbl(txtMedicineID.Text), dbconnection, adOpenStatic, adLockReadOnly

If Not recordsetCheck.BOF And Not recordsetCheck.EOF Then
MsgBox "Medicine ID already exists. Please enter a unique Medicine ID.", vbExclamation
txtMedicineID.SetFocus
Exit Sub
End If

If Len(Trim(txtMedicineName.Text)) = 0 Then
MsgBox "Medicine Name is required. Please enter it.", vbExclamation
txtMedicineName.SetFocus
Exit Sub
End If

If Len(Trim(cmbMedicineCategory.Text)) = 0 Then
MsgBox "Medicine Category is required. Please enter it.", vbExclamation
cmbMedicineCategory.SetFocus
Exit Sub
End If

If Len(Trim(cmbMedicineType.Text)) = 0 Then
MsgBox "MedicineType is required. Please select it.", vbExclamation
cmbMedicineType.SetFocus
Exit Sub
End If

If Len(Trim(txtStockLevel.Text)) = 0 Then
MsgBox "Stock Level is required. Please enter it.", vbExclamation
txtStockLevel.SetFocus
Exit Sub
End If

If Len(Trim(txtCostPerUnit.Text)) = 0 Then
MsgBox "Cost Per Unit is required. Please enter it.", vbExclamation
txtCostPerUnit.SetFocus
Exit Sub
End If

If Len(Trim(txtReorderLevel.Text)) = 0 Then
MsgBox "Reorder Level is required. Please enter it.", vbExclamation
txtReorderLevel.SetFocus
Exit Sub
End If

Set recordset = New ADODB.recordset
recordset.Open "SELECT * FROM MedicineTable WHERE MedicineID = " & CDbl(txtMedicineID.Text), dbconnection, adOpenKeyset, adLockOptimistic

recordset.AddNew

With recordset
.Fields("MedicineID").Value = CDbl(txtMedicineID.Text)
.Fields("MedicineName").Value = txtMedicineName.Text
.Fields("MedicineCategory").Value = cmbMedicineCategory.Text
.Fields("MedicineType").Value = cmbMedicineType.Text
.Fields("StockLevel").Value = CDbl(txtStockLevel.Text)
.Fields("CostPerUnit").Value = CCur(txtCostPerUnit.Text)
.Fields("ReorderLevel").Value = CDbl(txtReorderLevel.Text)
.Update
End With
MsgBox "The Record has been added"

LoadRecordsIntoListView

cmdaddnewitem.Enabled = True
cmdsave.Enabled = False
Exit Sub
End Sub
Private Sub cmdEdit_Click()
If MedicineList.selectedItem Is Nothing Then
MsgBox "Please select an item to edit.", vbExclamation
Exit Sub
End If

If txtMedicineName.Text = "" Or cmbMedicineCategory.Text = "" Or cmbMedicineType.Text = "" Or txtStockLevel.Text = "" Or txtCostPerUnit.Text = "" Or txtReorderLevel.Text = "" Then
MsgBox "Please fill in all fields before saving.", vbExclamation
Exit Sub
End If

Set selectedItem = MedicineList.selectedItem

selectedItem.Text = txtMedicineID.Text
selectedItem.SubItems(1) = txtMedicineName.Text
selectedItem.SubItems(2) = cmbMedicineCategory.Text
selectedItem.SubItems(3) = cmbMedicineType.Text
selectedItem.SubItems(4) = txtStockLevel.Text
selectedItem.SubItems(5) = txtCostPerUnit.Text
selectedItem.SubItems(6) = txtReorderLevel.Text

sql = "UPDATE MedicineTable SET " & _
"MedicineName = '" & txtMedicineName.Text & "', " & _
"MedicineCategory = '" & cmbMedicineCategory.Text & "', " & _
"MedicineType = '" & cmbMedicineType.Text & "', " & _
"StockLevel = " & txtStockLevel.Text & ", " & _
"CostPerUnit = " & txtCostPerUnit.Text & ", " & _
"ReorderLevel = " & txtReorderLevel.Text & " " & _
"WHERE MedicineID = " & txtMedicineID.Text

dbconnection.Execute sql

MsgBox "Changes saved successfully!"
Exit Sub
End Sub
Private Sub cmdcancel_Click()
txtMedicineID.Text = ""
txtMedicineName.Text = ""
cmbMedicineCategory.Text = ""
cmbMedicineType.Text = ""
txtStockLevel.Text = ""
txtCostPerUnit.Text = ""
txtReorderLevel.Text = ""
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
.Open "SELECT * FROM MedicineTable", dbconnection, adOpenDynamic, adLockOptimistic
End With
With DataReportMedicine
Set .DataSource = recordset
.Refresh
.Show
End With

Exit Sub
dbdisconnect
End Sub
Private Sub cmdPrintInventory_Click()
dbconnect

If recordset Is Nothing Then Set recordset = New ADODB.recordset

With recordset
If .State = 1 Then .Close
.Open "SELECT * FROM MedicineTable", dbconnection, adOpenDynamic, adLockOptimistic
End With
With DataReportInventory
Set .DataSource = recordset
.Refresh
.Show
End With

Exit Sub
dbdisconnect
End Sub
Private Sub cmdPrintCategory_Click()
If Cmbcat.Text = "" Or Cmbcat.Text = "" Then
MsgBox "Please select a category before printing.", vbExclamation, "Selection Required"
Exit Sub
End If

dbconnect
Dim totalMedicines As Integer
Dim totalStock As Double
totalMedicines = 0
totalStock = 0

If recordset Is Nothing Then Set recordset = New ADODB.recordset
With recordset
If .State = 1 Then .Close
Dim selectedCat As String
selectedCat = Cmbcat.Text
Dim Query As String
If selectedCat = "ALL" Then
Query = "SELECT * FROM MedicineTable"
Else
Query = "SELECT * FROM MedicineTable WHERE MedicineCategory = '" & selectedCat & "'"
End If
.Open Query, dbconnection, adOpenDynamic, adLockOptimistic
If Not .EOF Then
.MoveFirst
Do While Not .EOF
totalMedicines = totalMedicines + 1
If Not IsNull(.Fields("StockLevel").Value) Then
totalStock = totalStock + .Fields("StockLevel").Value
End If
.MoveNext
Loop
End If
.MoveFirst
End With

MsgBox "" & selectedCat & vbCrLf & "Total Medicine: " & totalMedicines & vbCrLf & "Total Stock: " & totalStock

With DataReportCategory
Set .DataSource = recordset
With .Sections("Section5").Controls
Dim i As Integer
For i = 1 To .Count
If TypeOf .Item(i) Is RptLabel Then
If .Item(i).Name = "lblTotalMedicines" Then
.Item(i).Caption = "" & selectedCat & " - Total Medicines: " & totalMedicines & " - Total Stock: " & totalStock
End If
End If
Next i
End With
.Refresh
.Show
End With

dbdisconnect
End Sub


Private Sub cmdPrintCostanalysis_Click()
dbconnect

If recordset Is Nothing Then Set recordset = New ADODB.recordset

With recordset
If .State = 1 Then .Close
.Open "SELECT * FROM MedicineTable", dbconnection, adOpenDynamic, adLockOptimistic
End With
With DataReportcost
Set .DataSource = recordset
.Refresh
.Show
End With

Exit Sub
dbdisconnect
End Sub


