Option Compare Database

Private Sub btnClose_Click()
    'Close the form
    DoCmd.Close acForm, "frmItembyName"
End Sub

Private Sub btnOK_Click()
    'Declare variables to be used
    Dim DocName As String 'name for document to be opened
    Dim strCondition As String 'name for condition
    
    'Assign value to DocName
    DocName = "frmStock"
    'Assign condition where ItemName matches
    strCondition = "ItemName = '" & Me.cboItemName & "'"
    
    'Open Stock Form according to the condition
    DoCmd.OpenForm DocName, acNormal, , strCondition
    'Close the form
    DoCmd.Close acForm, "frmItembyName"
End Sub
