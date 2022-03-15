Option Compare Database

Private Sub btnClose_Click()
    'Exit Database
    DoCmd.Quit
End Sub
Private Sub btnFindbyDate_Click()
    'Open Find Supply by Date form
    DoCmd.OpenForm "frmSupplybyDate", acNormal
End Sub

Private Sub btnFindbyID_Click()
    'Open Find Supply by ID form
    DoCmd.OpenForm "frmSupplybyID", acNormal
End Sub

Private Sub btnFindbyOrderID_Click()
    'Open Find Supply by Order ID form
    DoCmd.OpenForm "frmSupplybyOrderID", acNormal
End Sub

Private Sub btnFindbySupplier_Click()
    'Open Find Supply by Supplier form
    DoCmd.OpenForm "frmSupplybySupplier", acNormal
End Sub

Private Sub btnForm_Click()
    'Open Supply Received Form
    DoCmd.OpenForm "frmSupplyReceived", acNormal
End Sub
Private Sub btnFromUntil_Click()
    'Open Find Supply Received From Until form
    DoCmd.OpenForm "frmSupplyReceivedFromUntil", acNormal
End Sub

Private Sub btnMainMenu_Click()
    'Open Main Menu
    DoCmd.OpenForm "frmMainMenu", acNormal
    'Close Supply Received Menu
    DoCmd.Close acForm, "frmSupplyReceivedMenu"
End Sub

Private Sub btnReport_Click()
    'Open Supply Received Report
    DoCmd.OpenReport "rptSupplyReceived", acViewReport
End Sub
