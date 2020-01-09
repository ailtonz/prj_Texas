Option Compare Database

Private Sub Form_Open(Cancel As Integer)

codPAG.DefaultValue = NewCod(Form.RecordSource, codPAG.ControlSource)

End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click
    
    Form_frmCadastros.lstCadastro.Requery
    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    MsgBox Err.Description
    Resume Exit_cmdFechar_Click
    
End Sub
Private Sub cmdSalvar_Click()
On Error GoTo Err_cmdSalvar_Click
    
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
    Form_frmCadastros.lstCadastro.Requery
    DoCmd.Close
    
Exit_cmdSalvar_Click:
    Exit Sub

Err_cmdSalvar_Click:
    MsgBox Err.Description
    Resume Exit_cmdSalvar_Click
    
End Sub
