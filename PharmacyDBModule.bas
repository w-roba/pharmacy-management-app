Attribute VB_Name = "Module1"
Option Explicit
Public dbconnection As ADODB.Connection
Public recordset As ADODB.recordset
Public UserLoggedIn As Boolean
Public Roles As String
Public Sub dbconnect()
Set dbconnection = New ADODB.Connection
With dbconnection
 If .State = adStateOpen Then .Close
 .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PharmacyDB.mdb;Persist Security Info=False"
End With
End Sub

Public Sub dbdisconnect()
If Not recordset Is Nothing Then
 If recordset.State = adStateOpen Then recordset.Close
 Set recordset = Nothing
End If

If Not dbconnection Is Nothing Then
 If dbconnection.State = adStateOpen Then dbconnection.Close
 Set dbconnection = Nothing
End If
End Sub
Public Function AllowOnlyNumbersForCostPerUnit(KeyAscii As Integer, obj As Control) As Integer

If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Then

If Len(obj.Text) = 0 And (KeyAscii = 48 Or KeyAscii = 46) Then
 AllowOnlyNumbersForCostPerUnit = 0
ElseIf KeyAscii = 46 Then
If InStr(obj.Text, ".") Then
 AllowOnlyNumbersForCostPerUnit = 0
Else
 AllowOnlyNumbersForCostPerUnit = KeyAscii
End If
Else
 AllowOnlyNumbersForCostPerUnit = KeyAscii
End If
Else
 AllowOnlyNumbersForCostPerUnit = 0
End If
End Function
Public Function AllowOnlyNumbersForTotalCost(KeyAscii As Integer, obj As Control) As Integer

If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 46 Then

If Len(obj.Text) = 0 And (KeyAscii = 48 Or KeyAscii = 46) Then
 AllowOnlyNumbersForTotalCost = 0
ElseIf KeyAscii = 46 Then
If InStr(obj.Text, ".") Then
 AllowOnlyNumbersForTotalCost = 0
Else
 AllowOnlyNumbersForTotalCost = KeyAscii
End If
Else
 AllowOnlyNumbersForTotalCost = KeyAscii
End If
Else
 AllowOnlyNumbersForTotalCost = 0
End If
End Function

Public Function AllowOnlyNumbers(KeyAscii As Integer, obj As Control) As Integer

If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
If KeyAscii = 48 Then
If Len(obj.Text) = 0 Then
 AllowOnlyNumbers = 0
Else
 AllowOnlyNumbers = KeyAscii
End If
Else
 AllowOnlyNumbers = KeyAscii
End If
Else
 AllowOnlyNumbers = 0
End If
End Function

Public Function AllowOnlyText(KeyAscii As Integer) As Integer
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
 AllowOnlyText = KeyAscii
Else
 AllowOnlyText = 0
End If
End Function
Public Function AllowOnlyAddress(KeyAscii As Integer, CurrentText As String) As Integer
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 45 Or KeyAscii = 44 Or KeyAscii = 39 Or KeyAscii = 8 Then

 If Len(CurrentText) = 0 And KeyAscii = 32 Then
     AllowOnlyAddress = 0
 Else
     AllowOnlyAddress = KeyAscii
 End If
Else
 AllowOnlyAddress = 0
End If
End Function
Public Function AllowOnlyEmail(KeyAscii As Integer, CurrentText As String) As Integer
If (KeyAscii >= 65 And KeyAscii <= 90) Or _
(KeyAscii >= 97 And KeyAscii <= 122) Or _
(KeyAscii >= 48 And KeyAscii <= 57) Or _
KeyAscii = 64 Or _
KeyAscii = 46 Or _
KeyAscii = 95 Or _
KeyAscii = 45 Or _
KeyAscii = 8 Then

 If Len(CurrentText) = 0 And KeyAscii = 46 Then
     AllowOnlyEmail = 0
 ElseIf Right(CurrentText, 1) = "." And KeyAscii = 46 Then
     AllowOnlyEmail = 0
 Else
     AllowOnlyEmail = KeyAscii
 End If
Else
 AllowOnlyEmail = 0
End If
End Function
Public Function AllowOnlyNumbersAndSinglePlus(KeyAscii As Integer, CurrentText As String) As Integer
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
 AllowOnlyNumbersAndSinglePlus = KeyAscii
ElseIf KeyAscii = 43 Then
 If Len(CurrentText) = 0 Then
     AllowOnlyNumbersAndSinglePlus = KeyAscii
 Else
     AllowOnlyNumbersAndSinglePlus = 0
 End If
Else
 AllowOnlyNumbersAndSinglePlus = 0
End If
End Function
Public Function AllowOnlyDate(KeyAscii As Integer, obj As Control) As Integer
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = Asc("/") Or KeyAscii = Asc("-") Then
 If KeyAscii = Asc("/") Or KeyAscii = Asc("-") Then
     If Len(obj.Text) = 0 Or Right(obj.Text, 1) = "/" Or Right(obj.Text, 1) = "-" Then
         AllowOnlyDate = 0
     Else
         AllowOnlyDate = KeyAscii
     End If
 Else
     AllowOnlyDate = KeyAscii
 End If
Else
 AllowOnlyDate = 0
End If
End Function
Public Function AllowNumbers(KeyAscii As Integer, obj As Control) As Integer
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
 AllowNumbers = KeyAscii
Else
 AllowNumbers = 0
End If
End Function

