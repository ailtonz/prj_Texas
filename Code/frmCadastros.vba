Option Compare Database
Option Explicit
Dim WithEvents mformulario As Form

Sub Filtrar(txtPesquisa)

Dim Max As Integer
Dim Sql As String
Dim SqlAux As String
Dim a As Integer
Dim Colunas As Integer

Max = 20

If strTabela = "qualquer tabela" Then
       
Else
   
   Sql = "Select "
   
   Sql = Sql & strIdentificacao & ", "
   For a = 0 To Max
       If strColuna(a) <> "" Then
          Sql = Sql & strColuna(a) & ", "
          Colunas = Colunas + 1
       End If
   Next a
   Sql = Left(Sql, Len(Sql) - 2) & " "
   
   Sql = Sql & " from " & strTabela & ", "
   For a = 0 To Max
       If strTabelaRel(a, 1) <> "" Then
          Sql = Sql & strTabelaRel(a, 1) & ", "
       End If
   Next a
   Sql = Left(Sql, Len(Sql) - 2) & " "
   
   SqlAux = ""
   For a = 0 To Max
       If strTabelaRel(a, 1) <> "" Then
          SqlAux = SqlAux & strTabelaRel(a, 1) & "." & strTabelaRel(a, 2) & " = " & strTabela & "." & strTabelaRel(a, 3) & ", "
       End If
   Next a
   If SqlAux <> "" Then
      Sql = Sql & " Where (" & SqlAux
      Sql = Left(Sql, Len(Sql) - 2) & ") "
   End If
   If Not IsNull(txtPesquisa) Then
      If SqlAux = "" Then
         Sql = Sql & " Where ("
      Else
         Sql = Sql & " AND ("
      End If
      For a = 0 To Max
          If strColuna(a) <> "" Then
             If strColunaFiltra(a) = 0 Then
                Sql = Sql & strColuna(a) & " Like '*" & LCase(Trim(txtPesquisa)) & "*' OR "
             End If
          End If
      Next a
      Sql = Left(Sql, Len(Sql) - 3) & ") "
   End If
   
   Sql = Sql & "Order By "
   If strOrdemC <> "" Then
      Sql = Sql & strOrdemC & " Asc "
   ElseIf strOrdemD <> "" Then
      Sql = Sql & strOrdemD & " Desc "
   Else
      Sql = Sql & strColuna(1) & " Asc "
   End If
   Sql = Sql & ";"
   
   lstCadastro.RowSource = Sql
   
   lstCadastro.ColumnHeads = True
   lstCadastro.ColumnCount = Colunas + 1
   lstCadastro.ColumnWidths = "0cm"
   For a = 1 To Max
       If strColunaTam(a) <> "" Then
          lstCadastro.ColumnWidths = lstCadastro.ColumnWidths & ";" & Str(strColunaTam(a)) & "cm"
       End If
   Next a
   
End If

End Sub

Private Sub Filtro_Click()

Filtrar txtPesquisa

End Sub

Private Sub Form_Load()

Caption = strTitulo
Filtrar txtPesquisa

End Sub

Private Sub lstCadastro_DblClick(Cancel As Integer)
cmdAlterar_Click
End Sub

Private Sub cmdNovo_Click()
Manipulacao strTabela, "Novo"
End Sub

Private Sub cmdAlterar_Click()
Manipulacao strTabela, "Alterar"
End Sub

Private Sub cmdExcluir_Click()
Manipulacao strTabela, "Excluir"
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click

    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    MsgBox Err.Description
    Resume Exit_cmdFechar_Click
    
End Sub

Private Sub Manipulacao(Tabela As String, Operacao As String)

If IsNull(lstCadastro.Value) And Operacao <> "Novo" Then
   Exit Sub
End If

'Formulario
Select Case Tabela
 
 Case "tblBoletos"
  Set mformulario = New Form_frmBoletos

 Case "tblClientes"
  Set mformulario = New Form_frmClientes
 
 Case "tblFuncionarios"
  Set mformulario = New Form_frmFuncionarios
 
 Case "tblPagamentos"
  Set mformulario = New Form_frmPagamentos
 
 Case "tblProdutos"
  Set mformulario = New Form_frmProdutos
 
 Case "tblFuncionarios"
  Set mformulario = New Form_frmFuncionarios
   
 Case "tblNotasFiscais"
  Set mformulario = New Form_frmNotasFiscais
 
 Case "tblOrcamentos"
  Set mformulario = New Form_frmOrcamentos
 
 Case "tblOS"
  Set mformulario = New Form_frmOS
 
 Case "tblMovimentacao"
  Set mformulario = New Form_frmMovimentacao
  
 Case "tblPedidos"
  Set mformulario = New Form_frmPedidos
  
  
   
End Select


Select Case Operacao
 Case "Novo"
  With mformulario
   .Caption = "Novo Registro"
   .AllowDeletions = False
   .AllowAdditions = True
   .Visible = True
  End With
  DoCmd.GoToRecord , , acNewRec
 Case "Alterar"
  With mformulario
   .Caption = "Alteração de registro"
   .Filter = strIdentificacao & " = " & lstCadastro.Value
   .FilterOn = True
   .AllowDeletions = False
   .AllowAdditions = False
   .Visible = True
  End With
 Case "Excluir"
  If MsgBox("Deseja excluir este registro?", vbInformation + vbOKCancel) = vbOK Then
     DoCmd.SetWarnings False
     DoCmd.RunSQL ("Delete from " & strTabela & " where " & strIdentificacao & " = " & lstCadastro.Value)
     DoCmd.SetWarnings True
  End If
End Select
lstCadastro.Requery

Saida:
End Sub
