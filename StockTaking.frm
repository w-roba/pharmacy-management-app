VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStockTaking 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "StockTaking"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13290
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
   ScaleHeight     =   6315
   ScaleWidth      =   13290
   Begin VB.TextBox txtNewStock 
      Height          =   615
      Left            =   3360
      MaxLength       =   25
      TabIndex        =   14
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox txtCountedStock 
      Height          =   615
      Left            =   3360
      MaxLength       =   25
      TabIndex        =   13
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Frame FrameStockTaking 
      Caption         =   "Stock Details"
      Height          =   5295
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtMedicineID 
         Height          =   615
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   7
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtMedicineName 
         Height          =   615
         Left            =   3000
         MaxLength       =   25
         TabIndex        =   6
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtStockLevel 
         Height          =   615
         Left            =   3000
         MaxLength       =   25
         TabIndex        =   5
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "New Stock"
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Counted Stock"
         Height          =   615
         Left            =   240
         TabIndex        =   11
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "MedicineID"
         Height          =   615
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Medicine Name"
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Stock Level"
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdupdatestock 
      Caption         =   "Update Stock"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   5640
      Width           =   1335
   End
   Begin MSComctlLib.ListView StockList 
      Height          =   5175
      Left            =   7920
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   9128
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
End
Attribute VB_Name = "frmStockTaking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dbconnect
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
With StockList
.View = lvwReport
.FullRowSelect = True
.GridLines = True

.ColumnHeaders.clear

.ColumnHeaders.Add , , "Medicine ID", 1500
.ColumnHeaders.Add , , "Medicine Name", 1500
.ColumnHeaders.Add , , "Stock Level", 1500
.ColumnHeaders.Add , , "Counted Quantity", 1500
.ColumnHeaders.Add , , "New Stock", 1500
End With
LoadStock
txtMedicineID.Enabled = False
txtMedicineName.Enabled = False
txtStockLevel.Enabled = False
FrameStockTaking.Enabled = False
cmdupdatestock.Enabled = False
cmdclear.Enabled = False
cmdexit.Enabled = True


End Sub
Private Sub LoadStock()
StockList.ListItems.clear

Dim rs As ADODB.recordset
Set rs = New ADODB.recordset
rs.Open "SELECT MedicineID, MedicineName, StockLevel FROM MedicineTable", dbconnection, adOpenStatic, adLockReadOnly

Do While Not rs.EOF
With StockList.ListItems.Add(, , rs.Fields("MedicineID").Value)
.SubItems(1) = rs.Fields("MedicineName").Value
.SubItems(2) = rs.Fields("StockLevel").Value
End With
rs.MoveNext
Loop
rs.Close
End Sub
Private Sub StockList_ItemClick(ByVal Item As MSComctlLib.ListItem)

txtMedicineID.Text = Item.Text
txtMedicineName.Text = Item.SubItems(1)
txtStockLevel.Text = Item.SubItems(2)
txtNewStock.Text = ""
FrameStockTaking.Enabled = True
cmdupdatestock.Enabled = True
cmdclear.Enabled = True
cmdexit.Enabled = True

End Sub
Private Sub cmdupdatestock_Click()
Dim selectedItem As ListItem
If StockList.selectedItem Is Nothing Then
MsgBox "Please select a stock taking entry to update.", vbExclamation
Exit Sub
End If

Set selectedItem = StockList.selectedItem

Dim countedQuantity As Integer
If Not IsNumeric(txtCountedStock.Text) Or Val(txtCountedStock.Text) < 0 Then
MsgBox "Please enter a valid counted quantity.", vbExclamation
Exit Sub
End If

countedQuantity = Val(txtCountedStock.Text)

Dim newStock As Integer
If Not IsNumeric(txtNewStock.Text) Or Val(txtNewStock.Text) < 0 Then
MsgBox "Please enter a valid new stock quantity.", vbExclamation
Exit Sub
End If

newStock = Val(txtNewStock.Text)

Dim MedicineID As Double
MedicineID = Val(selectedItem.Text)
Dim currentStockLevel As Integer
Dim rs As ADODB.recordset
Set rs = New ADODB.recordset
rs.Open "SELECT StockLevel FROM MedicineTable WHERE MedicineID = " & MedicineID, dbconnection, adOpenStatic, adLockReadOnly

If Not rs.EOF Then
currentStockLevel = rs.Fields("StockLevel").Value
Else
MsgBox "Medicine not found.", vbExclamation
rs.Close
Set rs = Nothing
Exit Sub
End If

rs.Close
Set rs = Nothing

Dim newStockLevel As Integer

If countedQuantity > 0 Then
If countedQuantity > currentStockLevel Then
Dim response As VbMsgBoxResult
response = MsgBox("Warning: The counted quantity exceeds the current stock level. The stock level will be updated accordingly." & vbCrLf & _
"Do you want to continue?", vbExclamation + vbYesNo)
If response = vbNo Then
txtCountedStock.SetFocus
Exit Sub
End If
End If
newStockLevel = countedQuantity
Else
newStockLevel = currentStockLevel + newStock
End If

If newStock > 0 And countedQuantity = 0 Then
newStockLevel = currentStockLevel + newStock
End If

dbconnection.Execute "UPDATE MedicineTable SET StockLevel = " & newStockLevel & " WHERE MedicineID = " & MedicineID

selectedItem.SubItems(3) = countedQuantity
selectedItem.SubItems(4) = newStock

MsgBox "Stock updated successfully. Final Stock Level: " & newStockLevel, vbInformation
txtMedicineID.Text = ""
txtMedicineName.Text = ""
txtCountedStock.Text = ""
txtStockLevel.Text = ""
txtNewStock.Text = ""
cmdupdatestock.Enabled = False
cmdclear.Enabled = False
cmdexit.Enabled = True
FrameStockTaking.Enabled = False

LoadStock

StockList.selectedItem = Nothing
End Sub
Private Sub CheckReorderLevels()
Dim rs As recordset
Dim sql As String
Dim message As String
message = "The following medicines need to be reordered:" & vbCrLf

sql = "SELECT MedicineName FROM MedicineTable WHERE StockLevel < ReorderLevel"

Set rs = dbconnection.Execute(sql)

If Not rs.EOF Then
Do While Not rs.EOF
message = message & rs.Fields("MedicineName").Value & vbCrLf
rs.MoveNext
Loop
MsgBox message, vbExclamation, "Reorder Notification"
Else
MsgBox "All medicines are above reorder levels.", vbInformation, "Stock Check"
End If

rs.Close
End Sub
Private Sub cmdclear_Click()
txtMedicineID.Text = ""
txtMedicineName.Text = ""
txtCountedStock.Text = ""
txtStockLevel.Text = ""
txtNewStock.Text = ""
StockList.selectedItem = Nothing
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub txtCountedStock_KeyPress(KeyAscii As Integer)
KeyAscii = AllowNumbers(KeyAscii, txtCountedStock)
End Sub
Private Sub txtNewStock_KeyPress(KeyAscii As Integer)
KeyAscii = AllowNumbers(KeyAscii, txtNewStock)
End Sub



