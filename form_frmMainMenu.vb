Option Compare Database
Private Sub btnClose_Click()
    'Exit Database
    DoCmd.Quit
End Sub
Private Sub btnOrdersMenu_Click()
    'Open Orders Menu
    DoCmd.OpenForm "frmOrdersMenu", acNormal
    'Close Main Menu
    DoCmd.Close acForm, "frmMainMenu"
End Sub
Private Sub btnStockMenu_Click()
    'Open Stock Menu
    DoCmd.OpenForm "frmStockMenu", acNormal
    'Close Main Menu
    DoCmd.Close acForm, "frmMainMenu"
End Sub

Private Sub btnSuppliersMenu_Click()
    'Open Suppliers Menu
    DoCmd.OpenForm "frmSuppliersMenu", acNormal
    'Close Main Menu
    DoCmd.Close acForm, "frmMainMenu"
End Sub

Private Sub btnSupplyMenu_Click()
    'Open Supply Received Menu
    DoCmd.OpenForm "frmSupplyReceivedMenu", acNormal
    'Close Main Menu
    DoCmd.Close acForm, "frmMainMenu"
End Sub

Private Sub lblSuppliers_Click()
    'Open Suppliers Form
    DoCmd.OpenForm "frmSuppliers", acNormal
End Sub
