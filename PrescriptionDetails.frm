VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrescription 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prescription"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15630
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
   ScaleHeight     =   7530
   ScaleWidth      =   15630
   Begin VB.CommandButton cmdPrintReport 
      Caption         =   "Generate Report"
      Height          =   495
      Left            =   8400
      TabIndex        =   21
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   495
      Left            =   3480
      TabIndex        =   20
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdaddnewprescription 
      Caption         =   "Add New Prescription"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Frame FramePrescriptionDetails 
      Caption         =   "Prescription Details"
      Height          =   6495
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   7815
      Begin VB.TextBox txtPrescriptionID 
         Height          =   615
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   19
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtDateIssued 
         Height          =   615
         Left            =   3000
         MaxLength       =   25
         TabIndex        =   18
         Top             =   4680
         Width           =   2655
      End
      Begin VB.ComboBox cmbInstructions 
         Height          =   345
         Left            =   3000
         TabIndex        =   16
         Top             =   5640
         Width           =   2655
      End
      Begin VB.ComboBox cmbDosage 
         Height          =   345
         Left            =   3000
         TabIndex        =   13
         Top             =   3120
         Width           =   2655
      End
      Begin VB.ComboBox cmbMedicineName 
         Height          =   345
         Left            =   3000
         TabIndex        =   12
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtPatientName 
         Height          =   615
         Left            =   3000
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtQuantity 
         Height          =   615
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   5
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "Date Issued"
         Height          =   495
         Left            =   360
         TabIndex        =   15
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Patient Name"
         Height          =   615
         Left            =   360
         TabIndex        =   14
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "PrescriptionID"
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Medicine Name"
         Height          =   615
         Left            =   360
         TabIndex        =   10
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Dosage"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Quantity"
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Instructions"
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   5760
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   6720
      Width           =   1335
   End
   Begin MSComctlLib.ListView PrescriptionList 
      Height          =   6375
      Left            =   8160
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
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
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmPrescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
dbconnect
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
With PrescriptionList
.View = lvwReport
.ColumnHeaders.Add , , "Prescription ID", 1500
.ColumnHeaders.Add , , "Patient Name", 1500
.ColumnHeaders.Add , , "Medicine Name", 1500
.ColumnHeaders.Add , , "Dosage", 1000
.ColumnHeaders.Add , , "Quantity", 1000
.ColumnHeaders.Add , , "Date Issued", 1500
.ColumnHeaders.Add , , "Instructions", 1500
End With

LoadMedicinesIntoComboBox
LoadDosagesIntoComboBox
LoadInstructionsIntoComboBox

txtDateIssued.Text = Format(Now, "yyyy-mm-dd")


txtPrescriptionID.Enabled = True
cmbMedicineName.Enabled = False
cmbDosage.Enabled = False
txtPatientName.Enabled = False
txtQuantity.Enabled = False
txtDateIssued.Enabled = False
cmbInstructions.Enabled = False
cmdaddnewprescription.Enabled = True
cmdsave.Enabled = False
cmdSubmit.Enabled = False
cmdclear.Enabled = False
cmdexit.Enabled = True
FramePrescriptionDetails.Enabled = False
End Sub

Private Sub LoadMedicinesIntoComboBox()
Dim recordset As ADODB.recordset
Set recordset = New ADODB.recordset
recordset.Open "SELECT MedicineName FROM MedicineTable", dbconnection, adOpenStatic, adLockReadOnly

cmbMedicineName.clear

Do While Not recordset.EOF
cmbMedicineName.AddItem recordset.Fields("MedicineName").Value
recordset.MoveNext
Loop

recordset.Close
Set recordset = Nothing
End Sub

Private Sub LoadDosagesIntoComboBox()
cmbDosage.clear
cmbDosage.AddItem "250 mg"
cmbDosage.AddItem "500 mg"
cmbDosage.AddItem "1 g"
cmbDosage.AddItem "2 g"
End Sub

Private Sub LoadInstructionsIntoComboBox()
cmbInstructions.clear
cmbInstructions.AddItem "Take after meals"
cmbInstructions.AddItem "Take with water"
cmbInstructions.AddItem "Do not exceed dosage"
cmbInstructions.AddItem "Take with food"
cmbInstructions.AddItem "Take on an empty stomach"
cmbInstructions.AddItem "Take twice daily"
cmbInstructions.AddItem "Take as needed"
cmbInstructions.AddItem "Do not exceed recommended dose"
End Sub

Private Sub cmdAddNewPrescription_Click()
txtPrescriptionID.Enabled = True
cmbMedicineName.Enabled = True
cmbDosage.Enabled = True
txtPatientName.Enabled = True
txtQuantity.Enabled = True
txtDateIssued.Enabled = False
cmbInstructions.Enabled = True

cmdsave.Enabled = False
cmdSubmit.Enabled = True
cmdclear.Enabled = True
cmdexit.Enabled = True

FramePrescriptionDetails.Enabled = True

cmdaddnewprescription.Enabled = False

txtPrescriptionID.Text = ""
cmbMedicineName.Text = ""
cmbDosage.Text = ""
txtPatientName.Text = ""
txtQuantity.Text = ""
cmbInstructions.Text = ""
txtPrescriptionID.SetFocus
End Sub
Private Sub cmdSubmit_Click()
If Trim(txtPrescriptionID.Text) = "" Then
MsgBox "Please enter a Prescription ID.", vbExclamation
Exit Sub
End If

Dim prescriptionID As Double
prescriptionID = CDbl(txtPrescriptionID.Text)

Set recordsetCheck = New ADODB.recordset
recordsetCheck.Open "SELECT 1 FROM PrescriptionTable WHERE PrescriptionID = " & prescriptionID, dbconnection, adOpenStatic, adLockReadOnly

If Not recordsetCheck.BOF And Not recordsetCheck.EOF Then
MsgBox "Prescription ID already exists. Please enter a unique Prescription ID.", vbExclamation
txtPrescriptionID.SetFocus
Exit Sub
End If
Dim Item As ListItem
Set Item = PrescriptionList.ListItems.Add(, , txtPrescriptionID.Text)
Item.SubItems(1) = txtPatientName.Text
Item.SubItems(2) = cmbMedicineName.Text
Item.SubItems(3) = cmbDosage.Text
Item.SubItems(4) = txtQuantity.Text
Item.SubItems(5) = txtDateIssued.Text
Item.SubItems(6) = cmbInstructions.Text

cmdsave.Enabled = True
End Sub
Private Sub cmdSave_Click()
If PrescriptionList.ListItems.Count = 0 Then
MsgBox "No Prescriptions to save."
Exit Sub
End If

Dim i As Integer
Dim prescriptionID As Double
Dim patientName As String
Dim medicineName As String
Dim dosage As String
Dim quantity As Double
Dim DateIssued As String
Dim Instructions As String

For i = 1 To PrescriptionList.ListItems.Count
With PrescriptionList.ListItems(i)
prescriptionID = CDbl(.Text)
patientName = .SubItems(1)
medicineName = .SubItems(2)
dosage = .SubItems(3)
quantity = CDbl(.SubItems(4))
DateIssued = .SubItems(5)
Instructions = .SubItems(6)

Dim sql As String
sql = "INSERT INTO PrescriptionTable (PrescriptionID, PatientName, MedicineName, Dosage, Quantity, DateIssued, Instructions) VALUES (" & _
prescriptionID & ", '" & Replace(patientName, "'", "''") & "', '" & Replace(medicineName, "'", "''") & "', '" & _
Replace(dosage, "'", "''") & "', " & quantity & ", #" & Format(DateIssued, "yyyy-mm-dd") & "#, '" & Replace(Instructions, "'", "''") & "')"
dbconnection.Execute sql
End With
Next i

MsgBox "All prescriptions saved successfully!"
clear
PrescriptionList.ListItems.clear
End Sub
Private Sub cmdclear_Click()
txtPrescriptionID.Text = ""
cmbMedicineName.Text = ""
cmbDosage.Text = ""
txtPatientName.Text = ""
txtQuantity.Text = ""
cmbInstructions.Text = ""
End Sub
Private Sub clear()
PrescriptionList.ListItems.clear
txtPrescriptionID.Text = ""
cmbMedicineName.Text = ""
cmbDosage.Text = ""
txtPatientName.Text = ""
txtQuantity.Text = ""
cmbInstructions.Text = ""
End Sub


Private Sub cmdexit_Click()
If MsgBox("Are you sure you want to exit?", vbYesNo, "Confirm Exit") = vbYes Then
dbdisconnect
Unload Me
End If
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) And KeyAscii <> vbKeyBack Then
KeyAscii = 0
End If
End Sub
Private Sub cmbInstruction_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub cmbDosage_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub cmbMedicineName_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub txtPatientName_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyText(KeyAscii)
End Sub
Private Sub txtPrescriptionID_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyNumbers(KeyAscii, txtPrescriptionID)
End Sub
Private Sub txtDateIssued_KeyPress(KeyAscii As Integer)
KeyAscii = AllowOnlyDate(KeyAscii, txtDateIssued)
End Sub
Private Sub cmdPrintReport_Click()
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



