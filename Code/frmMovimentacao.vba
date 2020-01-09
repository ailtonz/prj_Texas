Option Compare Database

Private Sub cmdEfetivar_Click()
    EfetivaMovimento (Me.CodMov)
    Form_frmCadastros.lstCadastro.Requery
    DoCmd.Close
End Sub

Private Sub Form_Current()
    mov_Tipo_Click
End Sub

Private Sub Form_Open(Cancel As Integer)

codPED.DefaultValue = NewCod(Form.RecordSource, codPED.ControlSource)

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

Private Sub mov_Tipo_Click()

    If mov_Tipo.Value = 1 Then
        Me.codPED.Enabled = False
    Else
        Me.codPED.Enabled = True
    End If
    
End Sub
