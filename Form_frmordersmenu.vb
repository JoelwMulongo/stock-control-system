Option Compare Database

Private Sub btnClose_Click()
    'Exit Database
    DoCmd.Quit
End Sub

Private Sub btnFindbyDate_Click()
    'Open Find Order by Date form
    DoCmd.OpenForm "frmOrderbyDate", acNormal
End Sub

Private Sub btnFindbyID_Click()
    'Open Find Order by ID form
    DoCmd.OpenForm "frmOrderbyID", acNormal
End Sub

Private Sub btnFindbySupplier_Click()
    'Open Find Order by Supplier form
    DoCmd.OpenForm "frmOrderbySupplier", acNormal
End Sub

Private Sub btnForm_Click()
    'Open Orders Form
    DoCmd.OpenForm "frmOrders", acNormal
End Sub

Private Sub btnFromUntil_Click()
    'Open Find Orders From Until form
    DoCmd.OpenForm "frmOrdersFromUntil", acNormal
End Sub

Private Sub btnMainMenu_Click()
    'Open Main Menu
    DoCmd.OpenForm "frmMainMenu", acNormal
    'Close Orders Menu
    DoCmd.Close acForm, "frmOrdersMenu"
End Sub

Private Sub btnPending_Click()
    'Declare variables 
    Dim DocName As String 'name for document to be opened
    Dim strCondition As String 'name for condition
    
    'Assign value to DocName
    DocName = "rptOrders"
    'Assign condition where Status is PENDING
    strCondition = "Status = 'PENDING'"
    
    'Open Orders Report according to condition
    DoCmd.OpenReport DocName, acViewReport, , strCondition
End Sub

Private Sub btnReceived_Click()
    'Declare variables
    Dim DocName As String 'name for document to be opened
    Dim strCondition As String 'name for condition
    
    'Assign value to DocName
    DocName = "rptOrders"
    'Assign condition where Status is RECEIVED
    strCondition = "Status = 'RECEIVED'"
    
    'Open Orders Report according to condition
    DoCmd.OpenReport DocName, acViewReport, , strCondition
End Sub

Private Sub btnReport_Click()
    'Open Orders Report
    DoCmd.OpenReport "rptOrders", acViewReport
End Sub

Private Sub Command5_Click()
DoCmd.OpenForm "frmItembyName", acNormal
End Sub
