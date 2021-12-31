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
                    DoCmd.Close acForm, "frmSuppliers"
                End If
        'If choice is No
        ElseIf vbNo Then
            'Clear all fields and close the form
            Me.Undo
            DoCmd.Close
        End If
    Else
        'Close form if fields are not updated
        DoCmd.Close acForm, "frmSuppliers"
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
    If IsNull(Me.SupplierName) = True Then
        MsgBox "Supplier Name is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Me.Address) = True Then
        MsgBox "Address is required!", vbExclamation, "Error Message"
    ElseIf IsNull(Me.MobileNo) = True Then
        MsgBox "Mobile No is required!", vbExclamation, "Error Message"
    Else
        'Save current record and open new record
        DoCmd.RunCommand acCmdSaveRecord
        DoCmd.GoToRecord , , acNewRec
    End If
End Sub

Private Sub LandlineNo_AfterUpdate()
    'Error message if data is not numeric
    If IsNumeric(Me.LandlineNo) = False Then
        MsgBox "Invalid Data! Please enter numeric data only.", vbCritical, "Error Message"
        'Erase and return focus to the field
        Me.LandlineNo = Null
        Me.SupplierID.SetFocus
        Me.LandlineNo.SetFocus
    End If
    
    'Error message if length is greater than 10
    If Len(Me.LandlineNo) > 10 Then
        MsgBox "Field size cannot exceed 10 characters.", vbCritical, "Error Message"
        'Erase and return focus to the field
        Me.LandlineNo = Null
        Me.SupplierID.SetFocus
        Me.LandlineNo.SetFocus
    End If
End Sub

Private Sub MobileNo_AfterUpdate()
    'Error message if data is not numeric
    If IsNumeric(Me.MobileNo) = False Then
        MsgBox "Invalid Data! Please enter numeric data only.", vbCritical, "Error Message"
        'Erase and return focus to the field
        Me.MobileNo = Null
        Me.SupplierID.SetFocus
        Me.MobileNo.SetFocus
    End If
    
    'Error message if length is greater than 11
    If Len(Me.MobileNo) > 11 Then
        MsgBox "Field size cannot exceed 11 characters.", vbCritical, "Error Message"
        'Erase and return focus to the field
        Me.MobileNo = Null
        Me.SupplierID.SetFocus
        Me.MobileNo.SetFocus
    End If
End Sub

Private Sub SupplierName_AfterUpdate()
    'Convert text to uppercase
    Me.SupplierName = UCase(Me.SupplierName)
    
    'Error message if length is greater than 40
    If Len(Me.SupplierName) > 40 Then
        MsgBox "Field size cannot exceed 40 characters.", vbCritical, "Error Message"
        'Return focus to the field
        Me.SupplierID.SetFocus
        Me.SupplierName.SetFocus
    End If
End Sub
