VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H80000002&
   Caption         =   "MDIForm1"
   ClientHeight    =   4605
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6960
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu Fmac 
      Caption         =   "Accounts"
      Begin VB.Menu FmUser 
         Caption         =   "User Registration"
      End
      Begin VB.Menu fmlg 
         Caption         =   "Login"
      End
      Begin VB.Menu fmrp 
         Caption         =   "Reset Password"
      End
   End
   Begin VB.Menu Mfile 
      Caption         =   "File"
      Begin VB.Menu Frun 
         Caption         =   "Run"
         Begin VB.Menu Fmmed 
            Caption         =   "Medicine Registration"
         End
         Begin VB.Menu Fmsup 
            Caption         =   "Supplier Registration"
         End
         Begin VB.Menu Fmop 
            Caption         =   "Order Placement"
         End
         Begin VB.Menu Fmsr 
            Caption         =   "Stock Received"
         End
         Begin VB.Menu fmst 
            Caption         =   "Stock Taking"
         End
         Begin VB.Menu Fmpr 
            Caption         =   "Prescriptions"
         End
         Begin VB.Menu Fmin 
            Caption         =   "Invoice"
         End
         Begin VB.Menu Fmpc 
            Caption         =   "Payment Confirmation"
         End
         Begin VB.Menu Fmup 
            Caption         =   "User Priviledges"
         End
      End
      Begin VB.Menu Fmexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Fmvr 
      Caption         =   "View Reports"
      Begin VB.Menu rptreg 
         Caption         =   "Medicine Registration"
      End
      Begin VB.Menu rptsup 
         Caption         =   "Supplier Registration"
      End
      Begin VB.Menu rptO 
         Caption         =   "Orders"
      End
      Begin VB.Menu rptsr 
         Caption         =   "Stock Received"
      End
      Begin VB.Menu rptpr 
         Caption         =   "Prescription"
      End
      Begin VB.Menu rptpay 
         Caption         =   "Payments"
      End
      Begin VB.Menu rptInv 
         Caption         =   "Invoice"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
dbconnect
If UserLoggedIn = False Then
frmLogin.Show
If UserLoggedIn Then
Mfile.Enabled = True
Fmvr.Enabled = True
Fmac.Enabled = True
Else
Mfile.Enabled = False
Fmvr.Enabled = False
Fmac.Enabled = False
End If
End If
Me.Caption = "Nanyuki Hospital Pharmacy Management System"
End Sub

Private Sub Fmlg_Click()
frmLogin.Show
End Sub

Private Sub Fmmed_Click()
frmMedicine.Show
Unload frmInvoice
Unload frmOrderPlacement
Unload FrmPaymentConfirmation
Unload frmPrescription
Unload frmStockReceived
Unload frmStockTaking
Unload frmSupplier
End Sub

Private Sub Fmop_Click()
frmOrderPlacement.Show
Unload frmInvoice
Unload frmMedicine
Unload FrmPaymentConfirmation
Unload frmPrescription
Unload frmStockReceived
Unload frmStockTaking
Unload frmSupplier

End Sub

Private Sub Fmpc_Click()
FrmPaymentConfirmation.Show
Unload frmInvoice
Unload frmOrderPlacement
Unload frmMedicine
Unload frmPrescription
Unload frmStockReceived
Unload frmStockTaking
Unload frmSupplier

End Sub

Private Sub Fmpr_Click()
frmPrescription.Show
Unload frmInvoice
Unload frmOrderPlacement
Unload frmMedicine
Unload FrmPaymentConfirmation
Unload frmStockReceived
Unload frmStockTaking
Unload frmSupplier


End Sub
Private Sub Fmrp_Click()
frmForgotPassword.Show
End Sub

Private Sub Fmsr_Click()
frmStockReceived.Show
Unload frmInvoice
Unload frmOrderPlacement
Unload frmMedicine
Unload FrmPaymentConfirmation
Unload frmPrescription
Unload frmStockTaking
Unload frmSupplier

End Sub

Private Sub Fmst_Click()
frmStockTaking.Show
Unload frmInvoice
Unload frmOrderPlacement
Unload frmMedicine
Unload FrmPaymentConfirmation
Unload frmPrescription
Unload frmStockReceived
Unload frmSupplier
End Sub

Private Sub Fmsup_Click()
frmSupplier.Show
Unload frmInvoice
Unload frmOrderPlacement
Unload frmMedicine
Unload FrmPaymentConfirmation
Unload frmPrescription
Unload frmStockReceived
Unload frmStockTaking
End Sub

Private Sub Fmup_Click()
frmUserPrivileges.Show
End Sub

Private Sub FmUser_Click()
frmUserReg.Show
End Sub
Private Sub Fmin_Click()
frmInvoice.Show
Unload frmOrderPlacement
Unload frmMedicine
Unload FrmPaymentConfirmation
Unload frmPrescription
Unload frmStockReceived
Unload frmStockTaking
Unload frmSupplier

End Sub

Private Sub rptInv_Click()
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

Private Sub rptO_Click()
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
Private Sub rptpay_Click()
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

dbdisconnect
End Sub

Private Sub rptpr_Click()
dbconnect
If recordset Is Nothing Then Set recordset = New ADODB.recordset
With recordset
If .State = 1 Then .Close
.Open "SELECT * FROM PrescriptionTable", dbconnection, adOpenDynamic, adLockOptimistic
End With
With DataReportPrescription
Set .DataSource = recordset
.Refresh
.Show
End With
Exit Sub
dbdisconnect
End Sub
Private Sub rptreg_Click()
dbconnect
If recordset Is Nothing Then Set recordset = New ADODB.recordset
With recordset
If .State = adStateOpen Then .Close
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
Private Sub rptsr_Click()
dbconnect
Dim totalStockReceivedCost As Double
totalStockReceivedCost = 0

If recordset Is Nothing Then Set recordset = New ADODB.recordset
With recordset
If .State = 1 Then .Close
.Open "SELECT * FROM StockReceivedTable", dbconnection, adOpenDynamic, adLockOptimistic
If Not .EOF Then
.MoveFirst
Do While Not .EOF
If Not IsNull(.Fields("totalCost").Value) Then
totalStockReceivedCost = totalStockReceivedCost + .Fields("totalCost").Value
End If
.MoveNext
Loop
End If
.MoveFirst
End With

MsgBox "Total Stock Received Cost: " & totalStockReceivedCost

With DataReportStockReceived
Set .DataSource = recordset
With .Sections("Section5").Controls
Dim i As Integer
For i = 1 To .Count
If TypeOf .Item(i) Is RptLabel Then
If .Item(i).Name = "lbltotalcost" Then
.Item(i).Caption = "Total cost in Ksh: " & totalStockReceivedCost
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
.Open "SELECT * FROM StockReceivedItemsTable", dbconnection, adOpenDynamic, adLockOptimistic
End With

With DataReportStockReceivedItems
Set .DataSource = recordset
With .Sections("Section5").Controls
For i = 1 To .Count
If TypeOf .Item(i) Is RptLabel Then
If .Item(i).Name = "lbltotalcost" Then
.Item(i).Caption = "Total cost: " & totalStockReceivedCost
End If
End If
Next i
End With
.Refresh
.Show
End With

dbdisconnect
End Sub

Private Sub rptsup_Click()
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
