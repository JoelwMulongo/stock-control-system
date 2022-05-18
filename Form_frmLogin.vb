Option Compare Database
Private Sub btnExit_Click()
    'Exit Database
    DoCmd.Quit
End Sub
Private Sub btnLogin_Click()
     'Set values for database and recordset
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset("tbl_login")
    
    'Assign values for Username and Password
    rst.MoveFirst
    Password = rst("Password")
    UserName = rst("Username")
    
    'Error message if Username field is empty
    If IsNull(Me.Combo12) = True Then
        MsgBox "Username is required!!", vbExclamation
        Me.Combo12.SetFocus
        'Exit the subroutine due to error
        Exit Sub
    End If
    'Error if entered username doesn't match existing username
    If Me.Combo12 <> UserName Then
        'Show error label
        Me.lblWrongUsername.Visible = True
        Me.Combo12.SetFocus
        'Exit the subroutine due to error
        Exit Sub
    End If
    'Hide error label
    Me.lblWrongUsername.Visible = False
    
    'Error message if Password field is empty
    If IsNull(Me.txtPassword) = True Then
        MsgBox "Password is required!!", vbExclamation
        Me.txtPassword.SetFocus
        'Exit the subroutine due to error
        Exit Sub
    End If
    
    'Error if entered password doesn't match existing password
    If Me.txtPassword <> Password Then
        'Show error label
        Me.lblWrongPassword.Visible = True
        Me.txtPassword.SetFocus
        'Exit subroutine
        Exit Sub
    End If
    'Hide error label
    Me.lblWrongPassword.Visible = False
    
    'Open Main Menu
    DoCmd.OpenForm "frmMainMenu"
    'Close Login Form
    DoCmd.Close acForm, "frmLogin"
End Sub

