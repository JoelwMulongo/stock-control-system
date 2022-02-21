Option Compare Database

Private Sub btnClose_Click()
    'Close the form
    DoCmd.Close acForm, "frmItembyID"
End Sub
Private Sub btnOK_Click()
    'Declare variables to be used
    Dim DocName As String 'name for document to be opened
    Dim strCondition As String 'name for condition
    
    'Assign value to DocName
    DocName = "frmStock"
    'Assign condition where ItemID matches
    strCondition = "ItemID = " & Me.txtItemID
    
    'Open Stock Form according to the condition
    DoCmd.OpenForm DocName, acNormal, , strCondition
    'Close the form
    DoCmd.Close acForm, "frmItembyID"
End Sub
