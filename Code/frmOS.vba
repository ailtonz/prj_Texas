Option Compare Database
Option Explicit
Dim WithEvents mRelatorio As Report

Private Sub cmdVisualizar_Click()
On Error GoTo Err_cmdVisualizar_Click


Set mRelatorio = New Report_rptOrdemDeServicos
 
  With mRelatorio
   .Caption = "Visualizando: " & codOS.Value
   .Filter = "codos = " & codOS.Value
   .FilterOn = True
   .Visible = True
  End With

Exit_cmdVisualizar_Click:
    Exit Sub

Err_cmdVisualizar_Click:
    MsgBox Err.Description
    Resume Exit_cmdVisualizar_Click
    
End Sub

Private Sub Form_Open(Cancel As Integer)

codOS.DefaultValue = NewCod(Form.RecordSource, codOS.ControlSource)

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
