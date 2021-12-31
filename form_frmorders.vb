Option Compare Database

Private Sub btnClose_Click()
    'Ask to save record before closing form if Temp is equal to Yes
    If Me.txtTemp = "Yes" Then
        'If choice is Yes
        If MsgBox("Do you want to save the current record?", vbYesNo + vbQuestion, "Save Record") = vbYes Then
            'Call Save Button subroutine
            Call btnSave_Click
                'Close form if fields are not updated further
                If txtTemp = "No" Then
                    DoCmd.Close acForm, "frmOrders"
                End If
        'If choice is No
        ElseIf vbNo Then
            'Delete existing record without warning and close the form
            DoCmd.SetWarnings False
            DoCmd.RunCommand acCmdDeleteRecord
            DoCmd.SetWarnings True
            DoCmd.Close
        End If
    Else
        'Close form if Temp is equal to No
        DoCmd.Close acForm, "frmOrders"
    End If
End Sub

Private Sub btnDelete_Click()
    'Delete current record
    DoCmd.RunCommand acCmdDeleteRecord
End Sub

Private Sub btnFirst_Click()
    'Open first record
    DoCmd.GoToRecord , , acFirst
End Sub

Private Sub btnLast_Click()
    'Open last record
    DoCmd.GoToRecord , , acLast
End Sub

Private Sub btnNext_Click()
    'Open next record
    DoCmd.GoToRecord , , acNext
End Sub

Private Sub btnPrevious_Click()
    'Open previous record
    DoCmd.GoToRecord , , acPrevious
End Sub

Private Sub btnSave_Click()
    'Error messages if the respective fields are empty
    If IsNull(Me.SupplierID) = True Then
        MsgBox "Supplier ID is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Me.SupplierName) = True Then
        MsgBox "Supplier Name is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Me.OrderDate) = True Then
        MsgBox "Order Date is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Me.Status) = True Then
        MsgBox "Status is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Form_frmOrderDetails.ItemName) = True Then
        MsgBox "Item Name is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Form_frmOrderDetails.ItemID) = True Then
        MsgBox "Item ID is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Form_frmOrderDetails.Quantity) = True Then
        MsgBox "Quantity is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Form_frmOrderDetails.CostPrice) = True Then
        MsgBox "Cost Price is required!", vbExclamation, "Error Message"
    Else
        'Save current record and open new record
        DoCmd.RunCommand acCmdSaveRecord
        DoCmd.GoToRecord , , acNewRec
    End If
End Sub

Private Sub Form_AfterUpdate()
    'Change Temp value to Yes
    Me.txtTemp = "Yes"
End Sub

Private Sub Form_Current()
    'Deny access to subform
    Me.frmOrderDetails.Locked = True
    'Change Temp value to No
    Me.txtTemp = "No"
End Sub

Private Sub OrderDate_AfterUpdate()
    'Allow access to subform
    Me.frmOrderDetails.Locked = False
    'Change Temp value to Yes
    Me.txtTemp = "Yes"
End Sub

Private Sub Status_AfterUpdate()
    'Allow access to subform
    Me.frmOrderDetails.Locked = False
    'Change Temp value to Yes
    Me.txtTemp = "Yes"
End Sub

Private Sub SupplierID_AfterUpdate()
    'Error message if value is less and equal to zero
    If Me.SupplierID <= 0 Then
        MsgBox "Invalid Data! Make sure you enter a value greater than 0.", vbCritical, "Error Message"
        'Erase and return focus to the field
        Me.SupplierID = Null
        Me.OrderID.SetFocus
        Me.SupplierID.SetFocus
    End If
End Sub

Private Sub SupplierName_AfterUpdate()
        
    'Set values for database and recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("tblSuppliers")
    
    'Search recordset for required record
    rst.Index = "tblSuppliersSupplierName"
    rst.Seek "=", Me!SupplierName
    
    'Copy value from recordset
    Me!SupplierID = rst("SupplierID")
    'Close recordset
    rst.Close
    
    'Set variables to nothing
    Set rst = Nothing
    Set dbs = Nothing
    
    'Allow access to subform
    Me.frmOrderDetails.Locked = False
    'Change Temp value to Yes
    Me.txtTemp = "Yes"
End Sub
