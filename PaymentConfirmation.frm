VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPaymentConfirmation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PaymentConfirmation"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   15105
   Begin VB.CommandButton cmdGenerateReportByDate 
      Caption         =   "Filter by Date Received"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      TabIndex        =   23
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrintReport 
      Caption         =   "Generate Report"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   22
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   16
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdaddpayment 
      Caption         =   "Add Payment"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   15
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Frame FramePaymentDetails 
      Caption         =   "Payment  Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7695
      Begin VB.TextBox txtPaymentID 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   21
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox cmbPatientName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3000
         TabIndex        =   19
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox txtPaymentDate 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   18
         Top             =   4440
         Width           =   2655
      End
      Begin VB.TextBox txtInvoiceID 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   8
         Top             =   1200
         Width           =   2655
      End
      Begin VB.ComboBox cmbMedicineName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3000
         TabIndex        =   7
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtTotalCost 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   6
         Top             =   3480
         Width           =   2655
      End
      Begin VB.ComboBox cmbPaymentStatus 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "PaymentConfirmation.frx":0000
         Left            =   3000
         List            =   "PaymentConfirmation.frx":000D
         TabIndex        =   5
         Top             =   5400
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "PaymentID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Payment Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Top             =   4560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "InvoiceID"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   13
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Medicine Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   12
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Patient Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Total Cost"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   10
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Payment Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   5400
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdprocesspayment 
      Caption         =   "Process Payment"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdretrieve 
      Caption         =   "Fetch Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   0
      Top             =   6720
      Width           =   1335
   End
   Begin MSComctlLib.ListView PaymentList 
      Height          =   6255
      Left            =   7920
      TabIndex        =   17
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   11033
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
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   495
      Left            =   13560
      TabIndex        =   24
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   126746625
      CurrentDate     =   45691
   End
   Begin MSComCtl2.DTPicker dtpStartDate 
      Height          =   495
      Left            =   13560
      TabIndex        =   25
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   126746625
      CurrentDate     =   45691
   End
   Begin VB.Label Label6 
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   27
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   26
      Top             =   2640
      Width           =   735
   End
End
Attribute VB_Name = "FrmPaymentConfirmation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub Form_Load()
dbconnect
dtpStartDate.Value = Format(Now, "yyyy-mm-dd")
dtpEndDate.Value = Format(Now, "yyyy-mm-dd")
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
FramePaymentDetails.Enabled = False
cmdaddpayment.Enabled = True
cmdprocesspayment.Enabled = False
cmdSubmit.Enabled = False
cmdclear.Enabled = False
cmdexit.Enabled = True
cmdretrieve.Enabled = False
PaymentList.ColumnHeaders.clear
PaymentList.View = lvwReport
With PaymentList.ColumnHeaders
.Add , , "Payment ID", 1300
.Add , , "Invoice ID", 1000
.Add , , "Medicine Name", 1500
.Add , , "Patient Name", 1500
.Add , , "Total Cost", 1000
.Add , , "Payment Date", 1500
.Add , , "Payment Status", 1500


End With
ClearFields
txtPaymentDate.Text = Format(Now, "yyyy-mm-dd")
Exit Sub
End Sub
Private Sub cmdAddPayment_Click()
cmdaddpayment.Enabled = False
FramePaymentDetails.Enabled = True
txtInvoiceID.Enabled = True
txtPaymentID.Enabled = False
cmbMedicineName.Enabled = False
cmbPatientName.Enabled = False
txtTotalCost.Enabled = False
txtPaymentDate.Enabled = False
cmbPaymentStatus.Enabled = False

cmdprocesspayment.Enabled = False
cmdSubmit.Enabled = False
cmdclear.Enabled = True
cmdexit.Enabled = True
cmdretrieve.Enabled = True

ClearFields
txtInvoiceID.SetFocus
End Sub

Private Sub cmdclear_Click()
txtInvoiceID.Text = ""
txtPaymentID.Text = ""
cmbPatientName.Text = ""
cmbMedicineName.Text = ""
txtTotalCost.Text = ""
cmbPaymentStatus.Text = ""
End Sub
Private Sub ClearFields()
cmdaddpayment.Enabled = True
txtInvoiceID.Text = ""
txtPaymentID.Text = ""
cmbPatientName.Text = ""
cmbMedicineName.Text = ""
txtTotalCost.Text = ""
cmbPaymentStatus.Text = ""
End Sub
Private Sub cmdRetrieve_Click()
If Trim(txtInvoiceID.Text) = "" Then
MsgBox "Please enter an Invoice ID.", vbExclamation
Exit Sub
End If
Dim InvoiceID As Integer
InvoiceID = CInt(txtInvoiceID.Text)

Dim recordset As ADODB.recordset
Set recordset = New ADODB.recordset

Dim sql As String
sql = "SELECT PatientName, MedicineName, TotalCost FROM InvoiceTable WHERE InvoiceID = " & InvoiceID

recordset.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If Not recordset.EOF Then
cmbPatientName.Text = recordset.Fields("PatientName").Value
cmbMedicineName.Text = recordset.Fields("MedicineName").Value
txtTotalCost.Text = recordset.Fields("TotalCost").Value
Else
MsgBox "No invoice found with the given Invoice ID.", vbInformation
ClearFields
End If
txtInvoiceID.Enabled = False
cmbPaymentStatus.Enabled = True
txtPaymentID.Enabled = True
cmdSubmit.Enabled = True


recordset.Close
Set recordset = Nothing
End Sub
Private Sub cmdSubmit_Click()
dbconnect

If Trim(txtPaymentID.Text) = "" Then
MsgBox "Please enter Payment ID.", vbExclamation
Exit Sub
End If

Dim paymentID As Double
paymentID = CDbl(txtPaymentID.Text)

Dim recordsetCheck As ADODB.recordset
Set recordsetCheck = New ADODB.recordset
recordsetCheck.Open "SELECT 1 FROM Payment WHERE PaymentID = " & paymentID, dbconnection, adOpenStatic, adLockReadOnly

If Not recordsetCheck.BOF And Not recordsetCheck.EOF Then
MsgBox "Payment ID already exists in the database. Please enter a unique Payment ID.", vbExclamation
txtPaymentID.SetFocus
Exit Sub
End If

Dim i As Integer
For i = 1 To PaymentList.ListItems.Count
If PaymentList.ListItems(i).Text = txtPaymentID.Text Then
MsgBox "Payment ID already exists in the list. Please enter a unique Payment ID.", vbExclamation
txtPaymentID.SetFocus
Exit Sub
End If
Next i

Dim InvoiceID As Double
If Trim(txtInvoiceID.Text) = "" Then
MsgBox "Please enter Invoice ID.", vbExclamation
Exit Sub
End If
InvoiceID = CDbl(txtInvoiceID.Text)

For i = 1 To PaymentList.ListItems.Count
If PaymentList.ListItems(i).SubItems(1) = txtInvoiceID.Text Then
MsgBox "Invoice ID already exists in the list. Please enter a unique Invoice ID.", vbExclamation
Exit Sub
End If
Next i

If Len(Trim(cmbPaymentStatus.Text)) = 0 Then
MsgBox "The Payment Status is required. Please enter it.", vbExclamation
Exit Sub
End If

Dim Item As ListItem
Set Item = PaymentList.ListItems.Add(, , txtPaymentID.Text)
Item.SubItems(1) = txtInvoiceID.Text
Item.SubItems(2) = cmbMedicineName.Text
Item.SubItems(3) = cmbPatientName.Text
Item.SubItems(4) = txtTotalCost.Text
Item.SubItems(5) = txtPaymentDate.Text
Item.SubItems(6) = cmbPaymentStatus.Text

cmdprocesspayment.Enabled = True
cmdaddpayment.Enabled = True
cmdSubmit.Enabled = False
cmdclear.Enabled = True
cmdexit.Enabled = True
cmdretrieve.Enabled = False
txtInvoiceID.Enabled = False
cmbPaymentStatus.Enabled = False
txtPaymentID.Enabled = False

ClearFields
End Sub
Private Sub cmdprocesspayment_Click()
If PaymentList.ListItems.Count = 0 Then
MsgBox "No Payment to save.", vbExclamation
Exit Sub
End If

dbconnect

Dim i As Integer
Dim sql As String
Dim totalCost As Double
totalCost = 0

For i = 1 To PaymentList.ListItems.Count
With PaymentList.ListItems(i)
sql = "INSERT INTO Payment (PaymentID, InvoiceID, MedicineName, PatientName, TotalCost, PaymentDate, PaymentStatus) VALUES (" & _
.Text & ", " & _
.SubItems(1) & ", '" & Replace(.SubItems(2), "'", "''") & "', '" & _
Replace(.SubItems(3), "'", "''") & "', " & _
.SubItems(4) & ", #" & Format(.SubItems(5), "yyyy-mm-dd") & "#, '" & _
Replace(.SubItems(6), "'", "''") & "')"
dbconnection.Execute sql

totalCost = totalCost + CDbl(.SubItems(4))
GenerateReceiptForPayment .Text, totalCost
End With
Next i

MsgBox "All Payments saved successfully!", vbInformation
cmdprocesspayment.Enabled = False
PaymentList.ListItems.clear
End Sub

Private Sub GenerateReceiptForPayment(paymentID As String, totalCost As Double)
Set rs = New ADODB.recordset
Dim Query As String
Query = "SELECT * FROM Payment WHERE PaymentID = " & paymentID

rs.Open Query, dbconnection, adOpenDynamic, adLockOptimistic

If rs.EOF Then
MsgBox "No data found for the receipt.", vbExclamation, "No Data"
rs.Close
Exit Sub
End If
With DataReportReceipts
Set .DataSource = rs
Debug.Print "DataSource assigned successfully."
With .Sections("Section5").Controls
Dim i As Integer
For i = 1 To .Count
If TypeOf .Item(i) Is RptLabel Then
If .Item(i).Name = "lblTotalCost" Then
.Item(i).Caption = "Sub Total: " & totalCost
End If
End If
Next i
End With
.Refresh
.Show
End With
Exit Sub
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub txtPaymentID_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyNumbers(KeyAscii, txtPaymentID)
End Sub
Private Sub txtInvoiceID_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyNumbers(KeyAscii, txtInvoiceID)
End Sub
Private Sub cmbPaymentStatus_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub cmdPrintReport_Click()
dbconnect
Dim totalCost As Double
totalCost = 0

If recordset Is Nothing Then Set recordset = New ADODB.recordset
With recordset
If .State = 1 Then .Close
.Open "SELECT * FROM Payment", dbconnection, adOpenDynamic, adLockOptimistic
If Not .EOF Then
.MoveFirst
Do While Not .EOF
If Not IsNull(.Fields("TotalCost").Value) Then
totalCost = totalCost + .Fields("TotalCost").Value
End If
.MoveNext
Loop
End If
.MoveFirst
End With
MsgBox "Total Cost: " & totalCost
With DataReportPayments
Set .DataSource = recordset
With .Sections("Section5").Controls
Dim i As Integer
For i = 1 To .Count
If TypeOf .Item(i) Is RptLabel Then
If .Item(i).Name = "lblTotalCosts" Then
.Item(i).Caption = "Sub Total: " & totalCost
End If
End If
Next i
End With
.Refresh
.Show
End With
End Sub
Private Sub cmdGenerateReportByDate_Click()
Dim startDate As Date
Dim endDate As Date
startDate = dtpStartDate.Value
endDate = dtpEndDate.Value
If startDate > endDate Then
MsgBox "Start date cannot be later than end date.", vbExclamation
Exit Sub
End If

Dim rs As ADODB.recordset
Set rs = New ADODB.recordset

Dim sql As String
sql = "SELECT * FROM Payment WHERE PaymentDate BETWEEN #" & Format(startDate, "yyyy-mm-dd") & "# AND #" & Format(endDate, "yyyy-mm-dd") & "#"

rs.Open sql, dbconnection, adOpenStatic, adLockReadOnly

If rs.EOF Then
MsgBox "No records found for the specified date range.", vbExclamation
Exit Sub
End If

With DataReportPayments
Set .DataSource = rs
.Refresh
.Show
End With
End Sub
