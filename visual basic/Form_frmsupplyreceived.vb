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
                    DoCmd.Close acForm, "frmSupplyReceived"
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
        DoCmd.Close acForm, "frmSupplyReceived"
    End If
End Sub

Private Sub btnDelete_Click()
    'Delete current record
    DoCmd.RunCommand acCmdDeleteRecord
End Sub
Private Sub btnFill_Click()
    'Declare varibales to be used
    Dim dbs As DAO.Database 'name for current database
    Dim rst1 As DAO.Recordset 'name for new recordset
    Dim rst2 As DAO.Recordset 'name for new recordset
    Dim rst3 As DAO.Recordset 'name for new recordset
        
    'Set values for database and recordsets
    Set dbs = CurrentDb
    Set rst1 = dbs.OpenRecordset("tblOrders")
    Set rst2 = dbs.OpenRecordset("tblOrderDetails")
    Set rst3 = Me.frmSupplyReceivedDetails.Form.Recordset
     
    'Search recordset for required record
    rst1.Index = "PrimaryKey"
    rst1.Seek "=", Me!OrderID
    
    'Search recordset for required record
    rst2.Index = "tblOrderDetailsOrderID"
    rst2.Seek "=", Me!OrderID
    
    'If OrderID field is not empty then
    If IsNull(OrderID) = False Then
        'Display the order details and ask to fill fields
        'If choice is Yes
        If MsgBox("Do you want to fill the fields automatically based on the following order:" & vbCrLf & _
                    "Order ID: " & rst1("OrderID") & vbCrLf & _
                    "Supplier ID: " & rst1("SupplierID") & vbCrLf & _
                    "Supplier Name: " & rst1("SupplierName") & vbCrLf & _
                    "Order Date: " & rst1("OrderDate"), _
                    vbYesNo + vbQuestion, "The Smart Shop") = vbYes Then
            
            'Copy values from recordset
            Me!SupplierID = rst1("SupplierID")
            Me!SupplierName = rst1("SupplierName")
            
            'Close recordset
            rst1.Close
                
            'Turn focus to subform
            Me.frmSupplyReceivedDetails.SetFocus
                
            'Do while end of file is reached
            Do While Not rst2.EOF
                'If both OrderID's match then
                If Me!OrderID = rst2("OrderID") Then
                    'Add new record in the recordset
                    rst3.AddNew
                    rst3("SupplyID") = Me!SupplyID
                    'Copy values from the second recordset
                    rst3("ItemID") = rst2("ItemID")
                    rst3("ItemName") = rst2("ItemName")
                    rst3("CostPrice") = rst2("CostPrice")
                    'Update recordset
                    rst3.Update
                End If
                'Move to next record
                rst2.MoveNext
            Loop
            
            'Close recordset
            rst2.Close
            End If
    Else
        'Error message if OrderID field is empty
        MsgBox "Enter Order ID first!", vbExclamation, "Error Message"
    End If
    
    'Set variables to nothing
    Set rst1 = Nothing
    Set rst2 = Nothing
    Set dbs = Nothing
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
    If IsNull(Me.SupplierName) = True Then
        MsgBox "Supplier Name is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Me.SupplierID) = True Then
        MsgBox "Supplier ID is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Me.ReceivedDate) = True Then
        MsgBox "Received Date is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Form_frmSupplyReceivedDetails.ItemName) = True Then
        MsgBox "Item Name is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Form_frmSupplyReceivedDetails.ItemID) = True Then
        MsgBox "Item ID is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Form_frmSupplyReceivedDetails.Quantity) = True Then
        MsgBox "Quantity is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Form_frmSupplyReceivedDetails.CostPrice) = True Then
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
    'Allow access to subform
    Me.frmSupplyReceivedDetails.Locked = True
    'Change Temp value to No
    Me.txtTemp = "No"
End Sub

Private Sub OrderID_AfterUpdate()
    'Error message if value is less and equal to zero
    If Me.OrderID <= 0 Then
        MsgBox "Invalid Data! Make sure you enter a value greater than 0.", vbCritical, "Error Message"
        'Erase and return focus to the field
        Me.OrderID = Null
        Me.SupplyID.SetFocus
        Me.OrderID.SetFocus
    End If
    
    'Allow access to subform
    Me.frmSupplyReceivedDetails.Locked = False
    'Change Temp value to Yes
    Me.txtTemp = "Yes"
    
    'Declare varibales to be used
    Dim dbs As DAO.Database 'name for current database
    Dim rst As DAO.Recordset 'name for new recordset
        
    'Set values for database and recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("tblOrders")
     
    'Search recordset for required record
    rst.Index = "PrimaryKey"
    rst.Seek "=", Me!OrderID
    
    If Me.NewRecord Then
        'Edit and update the recordset
        rst.Edit
        rst("Status") = "RECEIVED"
        rst.Update
    End If
    'Close recordset
    rst.Close
        
    'Set variables to nothing
    Set rst = Nothing
    Set dbs = Nothing
End Sub

Private Sub ReceivedDate_AfterUpdate()
    'Allow access to subform
    Me.frmSupplyReceivedDetails.Locked = False
    'Change Temp value to Yes
    Me.txtTemp = "Yes"
End Sub

Private Sub SupplierID_AfterUpdate()
    'Error message if value is less and equal to zero
    If Me.SupplierID <= 0 Then
        MsgBox "Invalid Data! Make sure you enter a value greater than 0.", vbCritical, "Error Message"
        'Erase and return focus to the field
        Me.SupplierID = Null
        Me.SupplyID.SetFocus
        Me.SupplierID.SetFocus
    End If
End Sub

Private Sub SupplierName_AfterUpdate()
    'Declare varibales to be used
    Dim dbs As DAO.Database 'name for current database
    Dim rst As DAO.Recordset 'name for new recordset
        
    'Set values for database and recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("tblSuppliers")
    
    'Search recordset for required record
    rst.Index = "tblSuppliersSupplierName"
    rst.Seek "=", Me!SupplierName
    
    'Copy values from recordset
    Me!SupplierID = rst("SupplierID")
    'Close recordset
    rst.Close
    
    'Set variables to nothing
    Set rst = Nothing
    Set dbs = Nothing
    
    'Allow access to subform
    Me.frmSupplyReceivedDetails.Locked = False
    'Change Temp value to Yes
    Me.txtTemp = "Yes"
End Sub
