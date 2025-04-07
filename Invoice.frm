VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvoice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14745
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
   ScaleHeight     =   8025
   ScaleWidth      =   14745
   Begin VB.CommandButton cmdprintsales 
      Caption         =   "Generate Sales Report"
      Height          =   495
      Left            =   10920
      TabIndex        =   27
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrintReport 
      Caption         =   "Generate Report"
      Height          =   495
      Left            =   9480
      TabIndex        =   26
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdretrieve 
      Caption         =   "Fetch Details"
      Height          =   495
      Left            =   7920
      TabIndex        =   18
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   4800
      TabIndex        =   16
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6360
      TabIndex        =   15
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   495
      Left            =   1680
      TabIndex        =   14
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Frame FrameInvoiceDetails 
      Caption         =   "Invoice  Details"
      Height          =   6975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7695
      Begin VB.TextBox txtInvoiceDate 
         Height          =   495
         Left            =   3000
         TabIndex        =   25
         Top             =   5640
         Width           =   2655
      End
      Begin VB.TextBox txtPrescriptionID 
         Height          =   495
         Left            =   3000
         TabIndex        =   24
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtPatientName 
         Height          =   495
         Left            =   3000
         TabIndex        =   23
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox txtCostPerUnit 
         Height          =   495
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   22
         Top             =   4200
         Width           =   2655
      End
      Begin VB.TextBox txtTotalCost 
         Height          =   495
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   21
         Top             =   4920
         Width           =   2655
      End
      Begin VB.TextBox txtQuantity 
         Height          =   495
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   6
         Top             =   3480
         Width           =   2655
      End
      Begin VB.ComboBox cmbMedicineName 
         Height          =   345
         Left            =   3000
         TabIndex        =   5
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ComboBox cmbDosage 
         Height          =   345
         Left            =   3000
         TabIndex        =   4
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox txtInvoiceID 
         Height          =   495
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   3
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label9 
         Caption         =   "Total Cost"
         Height          =   615
         Left            =   360
         TabIndex        =   20
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Cost Per Unit"
         Height          =   615
         Left            =   360
         TabIndex        =   19
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Patient Name"
         Height          =   495
         Left            =   360
         TabIndex        =   13
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Quantity"
         Height          =   615
         Left            =   360
         TabIndex        =   12
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Dosage"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Medicine Name"
         Height          =   615
         Left            =   360
         TabIndex        =   10
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "InvoiceID"
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "PrescriptionID"
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Invoice Date"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   5760
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdaddInvoice 
      Caption         =   "Add Invoice"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   7200
      Width           =   1335
   End
   Begin MSComctlLib.ListView InvoiceList 
      Height          =   6855
      Left            =   7920
      TabIndex        =   17
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   12091
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
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dbconnect
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
FrameInvoiceDetails.Enabled = False
cmdaddInvoice.Enabled = True
cmdsave.Enabled = False
cmdSubmit.Enabled = False
cmdclear.Enabled = False
cmdexit.Enabled = True
cmdretrieve.Enabled = False
InvoiceList.ColumnHeaders.clear
InvoiceList.View = lvwReport
With InvoiceList.ColumnHeaders
.Add , , "Invoice ID", 1000
.Add , , "Prescription ID", 1000
.Add , , "Medicine Name", 1500
.Add , , "Patient Name", 1500
.Add , , "Dosage", 1000
.Add , , "Quantity", 1000
.Add , , "Cost per Unit", 1000
.Add , , "Total Cost", 1000
.Add , , "Invoice Date", 1500

End With
ClearFields
txtInvoiceDate.Text = Format(Now, "yyyy-mm-dd")
Exit Sub
End Sub
Private Sub cmdAddInvoice_Click()
FrameInvoiceDetails.Enabled = True
txtPrescriptionID.Enabled = True
txtInvoiceID.Enabled = False
cmbMedicineName.Enabled = False
txtPatientName.Enabled = False
cmbDosage.Enabled = False
txtQuantity.Enabled = False
txtCostPerUnit.Enabled = False
txtTotalCost.Enabled = False
txtInvoiceDate.Enabled = False

cmdaddInvoice.Enabled = False
cmdsave.Enabled = False
cmdSubmit.Enabled = False
cmdclear.Enabled = True
cmdexit.Enabled = True
cmdretrieve.Enabled = True





ClearFields
txtPrescriptionID.SetFocus
End Sub

Private Sub cmdclear_Click()
ClearFields
End Sub

Private Sub cmdRetrieve_Click()
FrameInvoiceDetails.Enabled = True

If Trim(txtPrescriptionID.Text) = "" Then
MsgBox "Please enter a Prescription ID.", vbExclamation
Exit Sub
End If
Dim prescriptionID As Double
prescriptionID = CDbl(txtPrescriptionID.Text)
Dim recordset As ADODB.recordset
Set recordset = New ADODB.recordset
Dim sql As String
sql = "SELECT PatientName, MedicineName, Quantity, Dosage FROM PrescriptionTable WHERE PrescriptionID = " & prescriptionID

recordset.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If Not recordset.EOF Then
txtPatientName.Text = recordset.Fields("PatientName").Value
cmbMedicineName.Text = recordset.Fields("MedicineName").Value
txtQuantity.Text = recordset.Fields("Quantity").Value
cmbDosage.Text = recordset.Fields("Dosage").Value
txtPrescriptionID.Enabled = False
txtInvoiceID.Enabled = True
cmbMedicineName.Enabled = False
txtPatientName.Enabled = False
cmbDosage.Enabled = False
txtQuantity.Enabled = False
txtCostPerUnit.Enabled = False
txtTotalCost.Enabled = False
txtInvoiceDate.Enabled = True


Dim MedicineName As String
MedicineName = cmbMedicineName.Text

Dim costRecordset As ADODB.recordset
Set costRecordset = New ADODB.recordset

Dim costSql As String
costSql = "SELECT CostPerUnit FROM MedicineTable WHERE MedicineName = '" & Replace(MedicineName, "'", "''") & "'"

costRecordset.Open costSql, dbconnection, adOpenStatic, adLockReadOnly

If Not costRecordset.EOF Then
Dim costPerUnit As Double
costPerUnit = costRecordset.Fields("CostPerUnit").Value

Dim Quantity As Double
Quantity = CDbl(txtQuantity.Text)
Dim totalCost As Double
totalCost = costPerUnit * Quantity

txtCostPerUnit.Text = costPerUnit
txtTotalCost.Text = totalCost
Else
MsgBox "Cost information not found for the selected medicine.", vbExclamation
txtCostPerUnit.Text = ""
txtTotalCost.Text = ""
End If

costRecordset.Close
Set costRecordset = Nothing
Else
MsgBox "No record found for the given Prescription ID.", vbInformation
End If

recordset.Close
Set recordset = Nothing
cmdSubmit.Enabled = True

End Sub
Private Sub cmdSubmit_Click()
dbconnect

If Trim(txtInvoiceID.Text) = "" Then
MsgBox "Please enter an Invoice ID.", vbExclamation
txtInvoiceID.Enabled = True
Exit Sub
End If

Dim InvoiceID As Double
InvoiceID = CDbl(txtInvoiceID.Text)

If CheckDuplicateID("InvoiceID", InvoiceID) Then
MsgBox "Invoice ID already exists in the database. Please enter a unique Invoice ID.", vbExclamation
txtInvoiceID.SetFocus
Exit Sub
End If

Dim prescriptionID As Double
prescriptionID = CDbl(txtPrescriptionID.Text)

If CheckDuplicateID("PrescriptionID", prescriptionID) Then
MsgBox "Prescription ID already exists in the database. Please enter a unique Prescription ID.", vbExclamation
Exit Sub
End If

Dim i As Integer
For i = 1 To InvoiceList.ListItems.Count
If InvoiceList.ListItems(i).Text = txtInvoiceID.Text Then
MsgBox "Invoice ID already exists in the list. Please enter a unique Invoice ID.", vbExclamation
txtInvoiceID.SetFocus
Exit Sub
End If
Next i

For i = 1 To InvoiceList.ListItems.Count
If InvoiceList.ListItems(i).SubItems(1) = txtPrescriptionID.Text Then
MsgBox "Prescription ID already exists in the list. Please enter a unique Prescription ID.", vbExclamation
Exit Sub
End If
Next i

If Len(Trim(txtInvoiceDate.Text)) = 0 Then
MsgBox "The Date is required. Please enter it.", vbExclamation
Exit Sub
End If

Dim Item As ListItem
Set Item = InvoiceList.ListItems.Add(, , txtInvoiceID.Text)
Item.SubItems(1) = txtPrescriptionID.Text
Item.SubItems(2) = cmbMedicineName.Text
Item.SubItems(3) = txtPatientName.Text
Item.SubItems(4) = cmbDosage.Text
Item.SubItems(5) = txtQuantity.Text
Item.SubItems(6) = txtCostPerUnit.Text
Item.SubItems(7) = txtTotalCost.Text
Item.SubItems(8) = txtInvoiceDate.Text

cmdsave.Enabled = True
cmdaddInvoice.Enabled = True
cmdSubmit.Enabled = False
cmdclear.Enabled = True
cmdexit.Enabled = True
cmdretrieve.Enabled = False

ClearFields
End Sub

Private Function CheckDuplicateID(fieldName As String, idValue As Double) As Boolean
Dim recordsetCheck As ADODB.recordset
Set recordsetCheck = New ADODB.recordset
recordsetCheck.Open "SELECT 1 FROM InvoiceTable WHERE " & fieldName & " = " & idValue, dbconnection, adOpenStatic, adLockReadOnly
CheckDuplicateID = Not recordsetCheck.BOF And Not recordsetCheck.EOF

recordsetCheck.Close
Set recordsetCheck = Nothing
End Function
Private Sub cmdSave_Click()
    If InvoiceList.ListItems.Count = 0 Then
        MsgBox "No invoices to save.", vbExclamation
        Exit Sub
    End If

    Dim i As Integer
    Dim sql As String
    Dim rs As ADODB.recordset
    Set rs = New ADODB.recordset

    dbconnect

    For i = 1 To InvoiceList.ListItems.Count
        With InvoiceList.ListItems(i)
            sql = "INSERT INTO InvoiceTable (InvoiceID, PrescriptionID, MedicineName, PatientName, Dosage, Quantity, CostPerUnit, TotalCost, InvoiceDate) VALUES (" & _
                  .Text & ", " & _
                  .SubItems(1) & ", '" & Replace(.SubItems(2), "'", "''") & "', '" & _
                  Replace(.SubItems(3), "'", "''") & "', '" & _
                  Replace(.SubItems(4), "'", "''") & "', " & _
                  .SubItems(5) & ", " & _
                  .SubItems(6) & ", " & _
                  .SubItems(7) & ", #" & Format(.SubItems(8), "yyyy-mm-dd") & "#)"

            dbconnection.Execute sql

            Dim MedicineName As String
            MedicineName = Replace(.SubItems(2), "'", "''")
            Dim Quantity As Long
            Quantity = .SubItems(5)

            ' Fetch MedicineID based on MedicineName
            rs.Open "SELECT MedicineID FROM MedicineTable WHERE MedicineName = '" & MedicineName & "'", dbconnection, adOpenDynamic, adLockOptimistic
            If Not rs.EOF Then
                Dim MedicineID As Long
                MedicineID = rs.Fields("MedicineID").Value
                
                ' Debug information
                Debug.Print "MedicineID: " & MedicineID
                Debug.Print "Quantity: " & Quantity

                sql = "UPDATE MedicineTable SET StockLevel = StockLevel - " & Quantity & " WHERE MedicineID = " & MedicineID
                Debug.Print "Executing SQL: " & sql
                dbconnection.Execute sql
            End If
            rs.Close
        End With
    Next i

    MsgBox "All invoices saved successfully!", vbInformation
    cmdsave.Enabled = False
    InvoiceList.ListItems.clear

    dbdisconnect
End Sub



Private Sub ClearFields()
cmdaddInvoice.Enabled = True

txtInvoiceID.Text = ""
txtPrescriptionID.Text = ""
cmbMedicineName.Text = ""
txtPatientName.Text = ""
cmbDosage.Text = ""
txtQuantity.Text = ""
txtCostPerUnit.Text = ""
txtTotalCost.Text = ""
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub txtPrescriptionID_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyNumbers(KeyAscii, txtPrescriptionID)
End Sub
Private Sub txtInvoiceID_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyNumbers(KeyAscii, txtInvoiceID)
End Sub
Private Sub txtInvoiceDate_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyDate(KeyAscii, txtInvoiceDate)
End Sub
Private Sub cmdPrintReport_Click()
dbconnect
Dim totalInvoiceCost As Double
totalInvoiceCost = 0

If recordset Is Nothing Then Set recordset = New ADODB.recordset
With recordset
If .State = 1 Then .Close
.Open "SELECT * FROM InvoiceTable", dbconnection, adOpenDynamic, adLockOptimistic
If Not .EOF Then
.MoveFirst
Do While Not .EOF
If Not IsNull(.Fields("totalCost").Value) Then
totalInvoiceCost = totalInvoiceCost + .Fields("totalCost").Value
End If
.MoveNext
Loop
End If
.MoveFirst
End With
MsgBox "Total Invoice Cost: " & totalInvoiceCost
With DataReportInvoice
Set .DataSource = recordset
With .Sections("Section5").Controls
Dim i As Integer
For i = 1 To .Count
If TypeOf .Item(i) Is RptLabel Then
If .Item(i).Name = "lbltotalcost" Then
.Item(i).Caption = "Total cost in Ksh: " & totalInvoiceCost
End If
End If
Next i
End With
.Refresh
.Show
End With

dbdisconnect
End Sub

Private Sub cmdPrintSales_Click()
dbconnect
If recordset Is Nothing Then Set recordset = New ADODB.recordset

With recordset
If .State = 1 Then .Close
.Open "SELECT * FROM InvoiceTable", dbconnection, adOpenDynamic, adLockOptimistic
End With
With DataReportInvoice
Set .DataSource = recordset
.Refresh
.Show
End With

Exit Sub
dbdisconnect
End Sub

