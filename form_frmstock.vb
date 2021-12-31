Option Compare Database

Private Sub btnClear_Click()
    'Clear all fields
    Me.Undo
End Sub

Private Sub btnClose_Click()
    'Ask to save record before closing form if fields are updated
    If Me.Dirty = True Then
        'If choice is Yes
        If MsgBox("Do you want to save the current record?", vbYesNo + vbQuestion, "Save Record") = vbYes Then
            'Call Save Button subroutine
            Call btnSave_Click
                'Close form if fields are not updated further
                If Me.Dirty = False Then
                    DoCmd.Close acForm, "frmStock"
                End If
        'If choice is No
        ElseIf vbNo Then
            'Clear all fields and close the form
            Me.Undo
            DoCmd.Close
        End If
    Else
        'Close form if fields are not updated
        DoCmd.Close acForm, "frmStock"
    End If
End Sub

Private Sub btnDelete_Click()
    'Delete current record
    DoCmd.RunCommand acCmdDeleteRecord
End Sub

Private Sub btnFirst_Click()
    'Undo form if fields are updated
    If Me.Dirty = True Then
        Me.Undo
    End If
    
    'Open first record
    DoCmd.GoToRecord , , acFirst
End Sub

Private Sub btnLast_Click()
    'Undo form if fields are updated
     If Me.Dirty = True Then
        Me.Undo
    End If
    
    'Open last record
    DoCmd.GoToRecord , , acLast
End Sub

Private Sub btnNext_Click()
    'Undo form if fields are updated
     If Me.Dirty = True Then
        Me.Undo
    End If
    
    'Open next record
    DoCmd.GoToRecord , , acNext
End Sub

Private Sub btnPrevious_Click()
    'Undo form if fields are updated
    If Me.Dirty = True Then
        Me.Undo
    End If
    
    'Open previous record
    DoCmd.GoToRecord , , acPrevious
End Sub

Private Sub btnSave_Click()
    'Error messages if the respective fields are empty
    If IsNull(Me.ItemName) = True Then
        MsgBox "Item Name is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Me.Category) = True Then
        MsgBox "Category is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Me.CurrentQuantity) = True Then
        MsgBox "Current Quantity is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Me.CostPrice) = True Then
        MsgBox "Cost Price is required!", vbExclamation, "Error Message"
    Else
        'Save current record and open new record
        DoCmd.RunCommand acCmdSaveRecord
        DoCmd.GoToRecord , , acNewRec
    End If
End Sub

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

Private Sub CurrentQuantity_AfterUpdate()
    'Error message if value is less than zero
    If Me.CurrentQuantity < 0 Then
        MsgBox "Invalid Data! Make sure you enter a value of 0 or greater.", vbCritical, "Error Message"
        'Erase and return focus to the field
        Me.CurrentQuantity = Null
        Me.ItemID.SetFocus
        Me.CurrentQuantity.SetFocus
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

Private Sub ItemName_AfterUpdate()
    'Convert text to uppercase
    Me.ItemName = UCase(Me.ItemName)
    
    'Error message if length is greater than 60
    If Len(Me.ItemName) > 60 Then
        MsgBox "Field size cannot exceed 60 characters.", vbCritical, "Error Message"
        'Return focus to the field
        Me.ItemID.SetFocus
        Me.ItemName.SetFocus
    End If
End Sub

Private Sub ReorderLevel_AfterUpdate()
    'Error message if value is less and equal to zero
    If Me.ReorderLevel <= 10 Then
        MsgBox "Item currently out of stock.", vbCritical, "Error Message"
        'Erase and return focus to the field
        Me.ReorderLevel = Null
        Me.ItemID.SetFocus
        Me.ReorderLevel.SetFocus
    End If
End Sub

