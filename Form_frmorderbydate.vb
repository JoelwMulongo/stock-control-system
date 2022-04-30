Option Compare Database
Private Sub btnClose_Click()
    'Close the form
    DoCmd.Close acForm, "frmOrderbyDate"
End Sub
Private Sub btnOK_Click()
    'Declare variables
    Dim DocName As String 'name for document to be opened
    Dim strCondition As String 'name for condition
    'Assign value to DocName
    DocName = "rptOrders"
    'Assign condition where OrderDate matches
    strCondition = "OrderDate = #" & Me.txtOrderDate & "#"
    
    'Open Orders Report
    DoCmd.OpenReport DocName, acViewReport, , strCondition
    'Close the form
    DoCmd.Close acForm, "frmOrderbyDate"
End Sub
