Option Compare Database

Private Sub CostPrice_AfterUpdate()
    'Error message if value is less and equal to zero
    If Me.CostPrice <= 0 Then
        MsgBox "Invalid Data! Make sure you enter a value greater than 0.", vbCritical, "Error Message"
        'Erase and return focus to the field
        Me.CostPrice = Null
        Me.ItemID.SetFocus
        Me.CostPrice.SetFocus
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Const conErrDataType = 2113
    
    'Replace standard error message with custom error message
    If DataErr = conErrDataType Then
        MsgBox "Invalid Data! Please enter numeric data only.", vbCritical, "Error Message"
        Response = acDataErrContinue
    Else
        'Display a standard error message
        Response = acDataErrDisplay
    End If
End Sub

Private Sub ItemID_AfterUpdate()
    'Error message if value is less and equal to zero
    If Me.ItemID <= 0 Then
        MsgBox "Invalid Data! Make sure you enter a value greater than 0", vbCritical, "Error Message"
        'Erase and return focus to the field
        Me.ItemID = Null
        Me.ItemName.SetFocus
        Me.ItemID.SetFocus
    End If
End Sub

Private Sub ItemName_AfterUpdate()
    'Declare varibales 
    Dim dbs As DAO.Database 'name for current database
    Dim rst As DAO.Recordset 'name for new recordset
        
    'Set values for database and recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("tblStock")
    
    'Search recordset for required record
    rst.Index = "tblStockItemName"
    rst.Seek "=", Me!ItemName
    
    'Copy values from recordset
    Me!CostPrice = rst("CostPrice")
    Me!ItemID = rst("ItemID")
    'Close recordset
    rst.Close
    
    'Set varibales to nothing
    Set rst = Nothing
    Set dbs = Nothing
End Sub

Private Sub Quantity_AfterUpdate()
    'Error message if value is less than zero
    If Me.Quantity < 0 Then
        MsgBox "Invalid Data! Make sure you enter a value of 0 or greater.", vbCritical, "Error Message"
        'Erase and return focus to the field
        Me.Quantity = Null
        Me.ItemID.SetFocus
        Me.Quantity.SetFocus
    End If
    
    'Declare varibales to be used
    Dim dbs As DAO.Database 'name for current database
    Dim rst As DAO.Recordset 'name for new recordset
            
    'Set values for database and recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("tblStock", dbOpenTable)
            
    'Search recordset for required record
    rst.Index = "PrimaryKey"
    rst.Seek "=", Me!ItemID
    'Edit recordset
    rst.Edit
    'Add entered quantity to current quantity
    rst("CurrentQuantity") = rst("CurrentQuantity") + Me!Quantity
    'Update and close recordset
    rst.Update
    rst.Close
            
    'Set variables to nothing
    Set rst = Nothing
    Set dbs = Nothing
End Sub
