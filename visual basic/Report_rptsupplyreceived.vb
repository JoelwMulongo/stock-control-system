Option Compare Database

Private Sub btnClose_Click()
    'Close the report
    DoCmd.Close acReport, "rptSupplyReceived"
End Sub

Private Sub lblDetails_Click()
    'Declare variables to be used
    Dim DocName As String 'name for document to be opened
    Dim strCondition As String 'name for condition
    
    'Assign value to DocName
    DocName = "frmSupplyReceived"
    'Assign condition where SupplyID matches
    strCondition = "SupplyID=" & SupplyID

    'Open Supply Received Form 
    DoCmd.OpenForm DocName, acNormal, , strCondition, acFormReadOnly
End Sub
