VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOrderPlacement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OrderPlacement"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14520
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
   ScaleHeight     =   8370
   ScaleWidth      =   14520
   Begin VB.CommandButton cmdFilterByDate 
      Caption         =   "Filter Orders by Date"
      Height          =   615
      Left            =   11880
      TabIndex        =   28
      Top             =   5040
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   495
      Left            =   12840
      TabIndex        =   25
      Top             =   4200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      Format          =   131465217
      CurrentDate     =   45691
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   495
      Left            =   12840
      TabIndex        =   24
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      _Version        =   393216
      Format          =   131465217
      CurrentDate     =   45691
   End
   Begin VB.CommandButton cmdPrintsup 
      Caption         =   "Filter by Supplier "
      Height          =   615
      Left            =   11880
      TabIndex        =   23
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox Cmbsup 
      Height          =   345
      ItemData        =   "OrderPlacement.frx":0000
      Left            =   11880
      List            =   "OrderPlacement.frx":000D
      TabIndex        =   22
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrintReport 
      Caption         =   "Generate Report"
      Height          =   615
      Left            =   11880
      TabIndex        =   21
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdadditem 
      Caption         =   "Add Item"
      Height          =   615
      Left            =   240
      TabIndex        =   20
      Top             =   7680
      Width           =   1695
   End
   Begin MSComctlLib.ListView OrderList 
      Height          =   6375
      Left            =   6240
      TabIndex        =   19
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   11245
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
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.ComboBox cmbMedicine 
      Height          =   345
      Left            =   2760
      TabIndex        =   18
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Frame FrameOrderDetails 
      Caption         =   "OrderDetails"
      Height          =   7335
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   5775
      Begin VB.ComboBox cmbSupplier 
         Height          =   345
         Left            =   2520
         TabIndex        =   17
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtOrderNumber 
         Height          =   735
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   9
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtCostPerUnit 
         Height          =   735
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   8
         Top             =   5280
         Width           =   2655
      End
      Begin VB.TextBox txtTotalCost 
         Height          =   735
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   7
         Top             =   6360
         Width           =   2655
      End
      Begin VB.TextBox txtQuantity 
         Height          =   735
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   6
         Top             =   4200
         Width           =   2655
      End
      Begin VB.TextBox txtOrderDate 
         Height          =   735
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "Total Cost"
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   6480
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Cost Per Unit"
         Height          =   615
         Left            =   360
         TabIndex        =   15
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Order Number"
         Height          =   615
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Supplier"
         Height          =   615
         Left            =   360
         TabIndex        =   13
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Order Date"
         Height          =   615
         Left            =   360
         TabIndex        =   12
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Medicine"
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Quantity"
         Height          =   615
         Left            =   360
         TabIndex        =   10
         Top             =   4320
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdsaveitem 
      Caption         =   "Save Item"
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdplaceorder 
      Caption         =   "Place Order"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdremoveitem 
      Caption         =   "Remove Item"
      Height          =   615
      Left            =   6360
      TabIndex        =   1
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   6360
      TabIndex        =   0
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "To"
      Height          =   255
      Left            =   11880
      TabIndex        =   27
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "From"
      Height          =   255
      Left            =   11880
      TabIndex        =   26
      Top             =   3600
      Width           =   735
   End
End
Attribute VB_Name = "frmOrderPlacement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
dtpStartDate.Value = Format(Now, "yyyy-mm-dd")
dtpEndDate.Value = Format(Now, "yyyy-mm-dd")
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
With OrderList
.View = lvwReport
.ColumnHeaders.Add , , "Medicine"
.ColumnHeaders.Add , , "Quantity"
.ColumnHeaders.Add , , "Total Cost"
End With

txtOrderNumber.Enabled = False
cmbSupplier.Enabled = False
txtOrderDate.Enabled = False
cmbMedicine.Enabled = False
txtQuantity.Enabled = False
txtCostPerUnit.Enabled = False
txtTotalCost.Enabled = False
cmdplaceorder.Enabled = False
cmdsaveitem.Enabled = False
cmdremoveitem.Enabled = False
cmdexit.Enabled = False
FrameOrderDetails.Enabled = False

dbconnect
LoadSuppliers
LoadMedicines
txtOrderDate.Text = Format(Now, "yyyy-mm-dd")
End Sub
Private Sub cmdadditem_Click()
txtOrderNumber.Enabled = True
cmbSupplier.Enabled = True
txtOrderDate.Enabled = True
cmbMedicine.Enabled = True
txtQuantity.Enabled = True
txtCostPerUnit.Enabled = True
txtTotalCost.Enabled = True
cmdplaceorder.Enabled = False
cmdremoveitem.Enabled = False
cmdsaveitem.Enabled = True
cmdexit.Enabled = True

FrameOrderDetails.Enabled = True

cmdadditem.Enabled = False

txtOrderNumber.SetFocus
Exit Sub
End Sub
Private Sub LoadSuppliers()
Dim rs As ADODB.recordset
Set rs = New ADODB.recordset
rs.Open "SELECT SupplierID, SupplierName FROM SupplierTable", dbconnection, adOpenStatic, adLockReadOnly

cmbSupplier.clear
Do While Not rs.EOF
cmbSupplier.AddItem rs.Fields("SupplierName").Value
cmbSupplier.ItemData(cmbSupplier.NewIndex) = rs.Fields("SupplierID").Value
rs.MoveNext
Loop
rs.Close
Set rs = Nothing
End Sub
Private Sub LoadMedicines()
Dim rs As ADODB.recordset
Set rs = New ADODB.recordset
rs.Open "SELECT MedicineID, MedicineName FROM MedicineTable", dbconnection, adOpenStatic, adLockReadOnly
cmbMedicine.clear
Do While Not rs.EOF
cmbMedicine.AddItem rs.Fields("MedicineName").Value
cmbMedicine.ItemData(cmbMedicine.NewIndex) = rs.Fields("MedicineID").Value
rs.MoveNext
Loop
rs.Close
Set rs = Nothing
End Sub
Private Sub cmdSaveItem_Click()
If txtOrderNumber.Text = "" Then
MsgBox "Please fill in the Order Number.", vbExclamation
Exit Sub
End If

If cmbSupplier.ListIndex = -1 Then
MsgBox "Please select a supplier.", vbExclamation
Exit Sub
End If

If cmbMedicine.ListIndex = -1 Then
MsgBox "Please select a medicine.", vbExclamation
Exit Sub
End If

If txtQuantity.Text = "" Then
MsgBox "Please enter the quantity.", vbExclamation
Exit Sub
End If

If txtOrderDate.Text = "" Then
MsgBox "Please enter the Order date.", vbExclamation
Exit Sub
End If

Dim Item As ListItem
Set Item = OrderList.ListItems.Add(, , cmbMedicine.Text)
Item.SubItems(1) = txtQuantity.Text
Item.SubItems(2) = txtTotalCost.Text

ClearMedicineFields
cmdadditem.Enabled = False
cmdplaceorder.Enabled = True
cmdremoveitem.Enabled = True

End Sub

Private Sub ClearMedicineFields()
cmbMedicine.ListIndex = -1
txtQuantity.Text = ""
txtCostPerUnit.Text = ""
txtTotalCost.Text = ""
End Sub

Private Sub cmbMedicine_Click()
UpdateCostAndTotal
End Sub

Private Sub txtQuantity_Change()
UpdateCostAndTotal
End Sub

Private Sub UpdateCostAndTotal()
If cmbMedicine.ListIndex <> -1 And txtQuantity.Text <> "" Then
Dim rs As ADODB.recordset
Set rs = New ADODB.recordset

Dim selectedMedicineID As Long
selectedMedicineID = cmbMedicine.ItemData(cmbMedicine.ListIndex)

rs.Open "SELECT CostPerUnit FROM MedicineTable WHERE MedicineID = " & selectedMedicineID, dbconnection, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
Dim costPerUnit As Currency
costPerUnit = rs.Fields("CostPerUnit").Value

txtCostPerUnit.Text = CStr(costPerUnit)
txtTotalCost.Text = CStr(costPerUnit * Val(txtQuantity.Text))
End If

rs.Close
Set rs = Nothing
Else
txtCostPerUnit.Text = ""
txtTotalCost.Text = ""
End If
End Sub
Private Sub cmdplaceorder_Click()
If OrderList.ListItems.Count = 0 Then
MsgBox "No items to place an order."
Exit Sub
End If
Dim OrderNumber As Double
OrderNumber = CDbl(txtOrderNumber.Text)
Dim sql As String
sql = "SELECT COUNT(*) AS OrderCount FROM OrderTable WHERE OrderNumber = " & OrderNumber
Dim rs As ADODB.recordset
Set rs = dbconnection.Execute(sql)

If Not rs.EOF Then
If rs.Fields("OrderCount").Value > 0 Then
MsgBox "This Order Number already exists in the database.", vbExclamation
Exit Sub
End If
End If

Dim SupplierID As Integer
SupplierID = cmbSupplier.ItemData(cmbSupplier.ListIndex)
Dim SupplierName As String
SupplierName = cmbSupplier.Text
Dim OrderDate As String
OrderDate = Format(Now, "yyyy-mm-dd")
Dim totalCost As Double
totalCost = CalculateTotalCost()
dbconnection.Execute "INSERT INTO OrderTable (OrderNumber, SupplierID, SupplierName, OrderDate, TotalCost) VALUES (" & OrderNumber & ", " & SupplierID & ", '" & Replace(SupplierName, "'", "''") & "', #" & OrderDate & "#, " & totalCost & ")"

Dim i As Integer
For i = 1 To OrderList.ListItems.Count
Dim MedicineID As Long
MedicineID = GetMedicineID(OrderList.ListItems(i).Text)

Dim Quantity As Long
Quantity = CLng(OrderList.ListItems(i).SubItems(1))

Dim ItemTotalCost As Double
ItemTotalCost = CDbl(OrderList.ListItems(i).SubItems(2))
Dim MedicineName As String
sql = "SELECT MedicineName FROM MedicineTable WHERE MedicineID = " & MedicineID
Set rs = dbconnection.Execute(sql)

If Not rs.EOF Then
MedicineName = rs.Fields("MedicineName").Value
Else
MsgBox "Medicine ID " & MedicineID & " not found for " & OrderList.ListItems(i).Text, vbExclamation
Exit Sub
End If
sql = "INSERT INTO OrderItemsTable (OrderNumber, MedicineID, MedicineName, Quantity, TotalCost) VALUES (" & OrderNumber & ", " & MedicineID & ", '" & Replace(MedicineName, "'", "''") & "', " & Quantity & ", " & ItemTotalCost & ")"
dbconnection.Execute sql
Next i

MsgBox "Order placed successfully!"
ClearForm
End Sub
Private Function GetMedicineID(MedicineName As String) As Long
Dim rs As ADODB.recordset
Set rs = New ADODB.recordset
Dim sql As String
sql = "SELECT MedicineID FROM MedicineTable WHERE MedicineName = '" & Replace(MedicineName, "'", "''") & "'"
rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
GetMedicineID = rs.Fields("MedicineID").Value
Else
GetMedicineID = 0
End If

rs.Close
Set rs = Nothing
End Function
Private Function CalculateTotalCost() As Double
Dim total As Double
Dim i As Integer
total = 0
For i = 1 To OrderList.ListItems.Count
total = total + Val(OrderList.ListItems(i).SubItems(2))
Next i
CalculateTotalCost = total
End Function

Private Sub ClearForm()
txtOrderNumber.Text = ""
cmbSupplier.ListIndex = -1
cmbMedicine.ListIndex = -1
txtQuantity.Text = ""
txtCostPerUnit.Text = ""
txtTotalCost.Text = ""
OrderList.ListItems.clear

txtQuantity.Enabled = False
txtCostPerUnit.Enabled = False
txtTotalCost.Enabled = False
cmdadditem.Enabled = True
cmdplaceorder.Enabled = False
cmdsaveitem.Enabled = False
cmdremoveitem.Enabled = False
End Sub

Private Sub cmdRemoveItem_Click()
If OrderList.selectedItem Is Nothing Then
MsgBox "Please select an item to remove.", vbExclamation
Exit Sub
End If
OrderList.ListItems.Remove OrderList.selectedItem.Index
End Sub
Private Sub txtOrderNumber_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyNumbers(KeyAscii, txtOrderNumber)
End Sub
Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyNumbers(KeyAscii, txtQuantity)
End Sub
Private Sub txtCostPerUnit_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyNumbersForCostPerUnit(KeyAscii, txtCostPerUnit)
End Sub
Private Sub txtTotalCost_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyNumbersForTotalCost(KeyAscii, txtTotalCost)
End Sub
Private Sub txtOrderDate_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyDate(KeyAscii, txtOrderDate)
End Sub
Private Sub cmbMedicine_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub cmbSupplier_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub cmdPrintReport_Click()
dbconnect

Dim totalOrderCost As Double
totalOrderCost = 0

If recordset Is Nothing Then Set recordset = New ADODB.recordset
With recordset
If .State = 1 Then .Close
.Open "SELECT * FROM OrderTable", dbconnection, adOpenDynamic, adLockOptimistic
If Not .EOF Then
.MoveFirst
Do While Not .EOF
If Not IsNull(.Fields("totalCost").Value) Then
totalOrderCost = totalOrderCost + .Fields("totalCost").Value
End If
.MoveNext
Loop
End If
.MoveFirst
End With
MsgBox "Total Order Cost: " & totalOrderCost
With DataReportOrders
Set .DataSource = recordset
With .Sections("Section5").Controls
Dim i As Integer
For i = 1 To .Count
If TypeOf .Item(i) Is RptLabel Then
If .Item(i).Name = "lbltotalcost" Then
.Item(i).Caption = "Total cost in Ksh: " & totalOrderCost
End If
End If
Next i
End With
.Refresh
.Show
End With
If recordset Is Nothing Then Set recordset = New ADODB.recordset
With recordset
If .State = 1 Then .Close
.Open "SELECT * FROM OrderItemsTable", dbconnection, adOpenDynamic, adLockOptimistic
End With

With DataReportOrderedItems
Set .DataSource = recordset
With .Sections("Section5").Controls
For i = 1 To .Count
If TypeOf .Item(i) Is RptLabel Then
If .Item(i).Name = "lbltotalcost" Then
.Item(i).Caption = "Total cost: " & totalOrderCost
End If
End If
Next i
End With
.Refresh
.Show
End With
dbdisconnect
End Sub



Private Sub cmdPrintsup_Click()
If Cmbsup.Text = "" Or Cmbsup.Text = "" Then
MsgBox "Please select a Supplier before printing.", vbExclamation, "Selection Required"
Exit Sub
End If
dbconnect
If recordset Is Nothing Then Set recordset = New ADODB.recordset
With recordset
If .State = 1 Then .Close
Dim selectedSupplier As String
selectedSupplier = Cmbsup.Text
Dim sqlQuery As String
If selectedSupplier = "ALL" Then
sqlQuery = "SELECT * FROM OrderTable"
Else
sqlQuery = "SELECT * FROM OrderTable WHERE SupplierName = '" & selectedSupplier & "'"
End If
.Open sqlQuery, dbconnection, adOpenDynamic, adLockOptimistic
End With
With DataReportsupplierfiltered
Set .DataSource = recordset
.Refresh
.Show
End With
Exit Sub
End Sub

Private Sub cmdFilterByDate_Click()
If dtpStartDate.Value > dtpEndDate.Value Then
MsgBox "Start date cannot be later than end date.", vbExclamation, "Invalid Date Range"
Exit Sub
End If

dbconnect
Dim totalMedicines As Integer
Dim totalQuantity As Integer
Dim totalCost As Double
totalMedicines = 0
totalQuantity = 0
totalCost = 0

If recordset Is Nothing Then Set recordset = New ADODB.recordset
With recordset
If .State = 1 Then .Close
Dim startDate As String
Dim endDate As String
startDate = Format(dtpStartDate.Value, "yyyy-mm-dd")
endDate = Format(dtpEndDate.Value, "yyyy-mm-dd")

Dim sqlQuery As String
sqlQuery = "SELECT o.OrderNumber, o.OrderDate, oi.MedicineName, oi.TotalCost, oi.Quantity " & _
"FROM OrderTable AS o " & _
"INNER JOIN OrderItemsTable AS oi ON o.OrderNumber = oi.OrderNumber " & _
"WHERE o.OrderDate BETWEEN #" & startDate & "# AND #" & endDate & "#"

.Open sqlQuery, dbconnection, adOpenDynamic, adLockOptimistic

If Not .EOF Then
.MoveFirst
Do While Not .EOF
totalMedicines = totalMedicines + 1
If Not IsNull(.Fields("Quantity").Value) Then
totalQuantity = totalQuantity + .Fields("Quantity").Value
End If
If Not IsNull(.Fields("TotalCost").Value) Then
totalCost = totalCost + .Fields("TotalCost").Value
End If
.MoveNext
Loop
End If
.MoveFirst
End With

MsgBox "Total Medicines: " & totalMedicines & vbCrLf & "Total Quantity: " & totalQuantity & vbCrLf & "Total Cost: " & totalCost

With DataReportOrderDT
Set .DataSource = recordset
With .Sections("Section5").Controls
Dim i As Integer
For i = 1 To .Count
If TypeOf .Item(i) Is RptLabel Then
If .Item(i).Name = "lblTotalMedicines" Then
.Item(i).Caption = "Total Medicine: " & totalMedicines
End If
If .Item(i).Name = "lblTotalQuantity" Then
.Item(i).Caption = "Total Quantity: " & totalQuantity
End If
If .Item(i).Name = "lblTotalCost" Then
.Item(i).Caption = "Total Cost: " & totalCost
End If
End If
Next i
End With
.Refresh
.Show
End With

dbdisconnect
End Sub



