Option Compare Database

Private Sub btnClose_Click()
    'Close the form
    DoCmd.Close acForm, "frmItembyCategory"
End Sub

Private Sub btnOK_Click()
    'Declare variables
    Dim DocName As String 'name for document to be opened
    Dim strCondition As String 'name for condition
    
    'Assign value to DocName
    DocName = "rptStock"
    'Assign condition where Category matches
    strCondition = "Category = '" & Me.cboCategory & "'"
    
    'Open Stock Form according to the condition
    DoCmd.OpenReport DocName, acViewReport, , strCondition
    'Close the form
    DoCmd.Close acForm, "frmItembyCategory"
End Sub
