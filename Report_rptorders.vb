Option Compare Database

Private Sub btnClose_Click()
    'Close the report
    DoCmd.Close acReport, "rptOrders"
End Sub

Private Sub lblDetails_Click()
    'Declare variables to be used
    Dim DocName As String 'name for document to be opened
    Dim strCondition As String 'name for condition
    
    'Assign value to DocName
    DocName = "frmOrders"
    'Assign condition where OrderID matches
    strCondition = "OrderID=" & OrderID
    
    'Open Orders Form according to the condition
    DoCmd.OpenForm DocName, acNormal, , strCondition, acFormReadOnly
End Sub
