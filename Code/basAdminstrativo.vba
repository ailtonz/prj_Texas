Option Compare Database
Option Explicit
Public strTabela As String
Public strTitulo As String

Public strIdentificacao As String
Public strColuna(20) As String
Public strColunaTam(20) As String
Public strColunaFiltra(20) As Integer

Public strTabelaRel(20, 20) As String

Public strOrdemD As String
Public strOrdemC As String

Sub AllForms()
    Dim obj As AccessObject, dbs As Object
    Set dbs = Application.CurrentProject
    ' Search for open AccessObject objects in AllForms collection.
    For Each obj In dbs.AllForms
        'If obj.IsLoaded = True Then
            ' Print name of obj.
            Debug.Print obj.Name
        'End If
    Next obj
End Sub

Public Function NewCod(Tabela, campo)

Dim rs1 As DAO.Recordset
Set rs1 = CurrentDb.OpenRecordset("SELECT Max([" & campo & "])+1 AS CodigoNovo FROM " & Tabela & ";")
If Not rs1.EOF Then
   NewCod = rs1.Fields("CodigoNovo")
   If IsNull(NewCod) Then
      NewCod = 1
   End If
Else
   NewCod = 1
End If
rs1.Close

End Function

Public Function Cadastro(Tabela As String)

Dim a As Integer

strTitulo = "Cadastro"
strTabela = Tabela
strOrdemD = ""
strOrdemC = ""
strIdentificacao = ""

For a = 1 To 20

   strColuna(a) = ""
   strColunaTam(a) = ""
   strColunaFiltra(a) = 0
   strTabelaRel(a, 1) = ""
   strTabelaRel(a, 2) = ""
   strTabelaRel(a, 3) = ""
   
Next a


Select Case Tabela
    
    Case "tblClientes"
       
       ' Código
       strIdentificacao = "codCLI"
       
       ' Campos da Tabela
       strColuna(1) = "cli_descricao"  ' Descrição
       strColuna(2) = strIdentificacao ' Código
       strColuna(3) = "cli_Telefone" ' Telefone
       
       strColunaTam(1) = "8"
       strColunaTam(2) = "4"
       strColunaTam(3) = "4"
             
    Case "tblProdutos"
       strIdentificacao = "codPRO"
       strColuna(1) = "pro_descricao"
       strColuna(2) = strIdentificacao
       strColuna(3) = "pro_Codigo"
       strColuna(4) = "pro_QTD"
       
       strColunaTam(1) = "8"
       strColunaTam(2) = "3"
       strColunaTam(3) = "3"
       strColunaTam(4) = "2"
       
       
    Case "tblFuncionarios"
       strIdentificacao = "codFUN"
       strColuna(1) = "fun_descricao"
       strColuna(2) = strIdentificacao
       strColunaTam(1) = "8"
       strColunaTam(2) = "4"
    
    Case "tblPagamentos"
       strIdentificacao = "codPAG"
       strColuna(1) = "pag_descricao"
       strColuna(2) = strIdentificacao
       strColunaTam(1) = "8"
       strColunaTam(2) = "4"
     
    Case "tblNotasFiscais"
       strIdentificacao = "codNF"
       strColuna(1) = strIdentificacao
       strColuna(2) = "cli_Descricao"
       strColunaTam(1) = "4"
       strColunaTam(2) = "8"
       strTabelaRel(1, 1) = "tblClientes"
       strTabelaRel(1, 2) = "codCLI"
       strTabelaRel(1, 3) = "codCLI"
       strOrdemD = "codNF"
    
    Case "tblBoletos"
       strIdentificacao = "codBOL"
       strColuna(1) = strIdentificacao
       strColuna(2) = "cli_Descricao"
       strColunaTam(1) = "4"
       strColunaTam(2) = "8"
       strTabelaRel(1, 1) = "tblClientes"
       strTabelaRel(1, 2) = "codCLI"
       strTabelaRel(1, 3) = "codCLI"
       strOrdemD = "codBOL"
       
    Case "tblOrcamentos"
       strIdentificacao = "codORC"
       
       strColuna(1) = strIdentificacao
       strColuna(2) = "cli_Descricao"
       strColuna(3) = "orc_data"
       strColuna(4) = "iif(orc_Acao=1,'Em aberto',iif(orc_Acao=2,'Vendeu','Cancelou')) as Status "
       
       strColunaFiltra(4) = 1
       
       strColunaTam(1) = "2"
       strColunaTam(2) = "8"
       strColunaTam(3) = "4"
       strColunaTam(4) = "4"
       
       strTabelaRel(1, 1) = "tblClientes"
       strTabelaRel(1, 2) = "codCLI"
       strTabelaRel(1, 3) = "codCLI"
       strOrdemD = "codORC"
     
    Case "tblOS"
       strIdentificacao = "codOS"
       
       strColuna(1) = strIdentificacao
       strColuna(2) = "cli_Descricao"
       strColuna(3) = "os_DataEntrega"
       strColuna(4) = "os_DataRetirada"
       
       strColunaTam(1) = "2"
       strColunaTam(2) = "6"
       strColunaTam(3) = "4"
       strColunaTam(4) = "4"
       
       strTabelaRel(1, 1) = "tblClientes"
       strTabelaRel(1, 2) = "codCLI"
       strTabelaRel(1, 3) = "codCLI"
       strOrdemD = "codOS"
     
    Case "tblPedidos"
       strIdentificacao = "codPED"
       
       strColuna(1) = strIdentificacao
       strColuna(2) = "cli_Descricao"
       strColuna(3) = "ped_data"
       
       strColunaTam(1) = "2"
       strColunaTam(2) = "8"
       strColunaTam(3) = "4"
       
       strTabelaRel(1, 1) = "tblClientes"
       strTabelaRel(1, 2) = "codCLI"
       strTabelaRel(1, 3) = "codCLI"
       strOrdemD = "codPED"
     
    Case "tblMovimentacao"
       strIdentificacao = "codMov"
       
       strColuna(1) = strIdentificacao
       strColuna(2) = "cli_Descricao"
       strColuna(3) = "iif(mov_Tipo=1,'Entrada','Saída') as Tipo"
       strColuna(4) = "mov_dtEntrada"
       
       strColunaFiltra(3) = 1
       
       strColunaTam(1) = "2"
       strColunaTam(2) = "8"
       strColunaTam(3) = "2"
       strColunaTam(4) = "2"
       
       strTabelaRel(1, 1) = "tblClientes"
       strTabelaRel(1, 2) = "codCLI"
       strTabelaRel(1, 3) = "codCLI"
       
       strOrdemD = "codMov"
     
End Select

DoCmd.OpenForm "frmCadastros"

End Function

Public Function GerarMovimentacao(codPedido As Integer) As Integer
'=================================================================================
'* Função               : GerarMovimentacao
'
'* Principal objetivo   : Gerar uma movimentação do estoque baseado em um pedido
'                         ainda ñ efetivado.
'
'* Tabelas envolvidas   : tblPedidos;tblPedidosItens;
'                         tblMovimentacao;tblMovimentacaoItens;
'
'=================================================================================

Dim dbDados As DAO.Database
Dim rstPedido As DAO.Recordset
Dim rstItensDoPedido As DAO.Recordset
Dim rstMovimento As DAO.Recordset
Dim rstItensDoMovimento As DAO.Recordset

Dim contRegistros As Long

Set dbDados = CurrentDb

Set rstPedido = dbDados.OpenRecordset("Select * from tblPedidos where codPed = " & codPedido)

'##### TESTES DE VALIDAÇÃO DO MOVIMENTO ####

If rstPedido.RecordCount > 0 Then
   If rstPedido.Fields("ped_Ok") Then
        MsgBox "ATENÇÃO: Este pedido já foi efetivado!", vbExclamation
        Exit Function
   Else
        Set rstItensDoPedido = dbDados.OpenRecordset("Select * from tblPedidosItens where codPed = " & codPedido)
        Set rstMovimento = dbDados.OpenRecordset("tblMovimentacao")
        Set rstItensDoMovimento = dbDados.OpenRecordset("tblMovimentacaoItens")
        
   End If
Else
   MsgBox "ATENÇÃO: Pedido não encontrado!", vbCritical
   Exit Function
End If

'##### GERAR MOVIMENTO E SEUS RESPECTIVOS ITENS ####

BeginTrans

rstMovimento.AddNew

rstMovimento.Fields("codPed") = rstPedido.Fields("codPed")
rstMovimento.Fields("mov_Tipo") = 2
rstMovimento.Fields("mov_dtEntrada") = Format(Now(), "dd/mm/yyyy")
rstMovimento.Fields("codCLI") = rstPedido.Fields("codCLI")
rstMovimento.Fields("mov_Obs") = rstPedido.Fields("ped_Obs")

rstMovimento.Update
rstMovimento.MoveLast

rstItensDoPedido.MoveLast
rstItensDoPedido.MoveFirst

For contRegistros = 1 To rstItensDoPedido.RecordCount

    rstItensDoMovimento.AddNew
    rstItensDoMovimento.Fields("codMov") = rstMovimento.Fields("codMov")
    rstItensDoMovimento.Fields("codPro") = rstItensDoPedido.Fields("codPro")
    rstItensDoMovimento.Fields("mvi_QTD") = rstItensDoPedido.Fields("pedi_QTD")
    rstItensDoMovimento.Fields("mvi_ValorUnitario") = rstItensDoPedido.Fields("pedi_Valor")
    rstItensDoMovimento.Update
    
    rstItensDoPedido.MoveNext

Next

rstPedido.Edit
rstPedido.Fields("ped_Ok") = True
rstPedido.Update

GerarMovimentacao = rstMovimento.Fields("codMov")

CommitTrans


End Function


Public Function EfetivaMovimento(codMovimento As Integer)
'=================================================================================
'* Função               : EfetivaMovimento
'
'* Principal objetivo   : Encontrar o movimento solicitado e baseado no tipo de
'                         movimento fazer a atualização junto ao estoque.
'
'* Tabelas envolvidas   : tblMovimentacao;tblMovimentacaoItens;tblProdutos.
'
'=================================================================================

Dim dbDados As DAO.Database
Dim rstMovimento As DAO.Recordset
Dim rstItensDoMovimento As DAO.Recordset
Dim rstProdutos As DAO.Recordset

Dim contRegistros As Long

Set dbDados = CurrentDb
Set rstMovimento = dbDados.OpenRecordset("Select * from tblMovimentacao where CodMov = " & codMovimento)

'##### TESTES DE VALIDAÇÃO DO MOVIMENTO ####

If rstMovimento.RecordCount > 0 Then
   If rstMovimento.Fields("mov_Ok") Then
        MsgBox "ATENÇÃO: Esta movimentação já foi efetivada!", vbExclamation
        Exit Function
   Else
        Set rstItensDoMovimento = dbDados.OpenRecordset("Select * from tblMovimentacaoItens where CodMov = " & codMovimento)
        Set rstProdutos = dbDados.OpenRecordset("Select * from tblProdutos")
   End If
Else
   MsgBox "ATENÇÃO: Movimento não encontrado!", vbCritical
   Exit Function
End If

'##### BAIXA DO ESTOQUE E EFETIVAÇÃO DO MOVIMENTO ####

BeginTrans

rstItensDoMovimento.MoveLast
rstItensDoMovimento.MoveFirst

If rstMovimento.Fields("mov_Tipo") = 1 Then
    For contRegistros = 1 To rstItensDoMovimento.RecordCount
    
       rstProdutos.MoveLast
       rstProdutos.FindFirst "CodPro = " & Str(rstItensDoMovimento.Fields("CodPro"))
       rstProdutos.Edit
       
       rstProdutos.Fields("pro_QTD") = rstProdutos.Fields("pro_QTD") + rstItensDoMovimento.Fields("mvi_QTD")
       rstProdutos.Update
       
       rstItensDoMovimento.MoveNext
    Next
Else
    For contRegistros = 1 To rstItensDoMovimento.RecordCount
    
       rstProdutos.MoveLast
       rstProdutos.FindFirst "CodPro = " & Str(rstItensDoMovimento.Fields("CodPro"))
       rstProdutos.Edit
       
       rstProdutos.Fields("pro_QTD") = rstProdutos.Fields("pro_QTD") - rstItensDoMovimento.Fields("mvi_QTD")
       rstProdutos.Update
       
       rstProdutos.MoveNext
    Next
End If

rstMovimento.Edit
rstMovimento.Fields("mov_Ok") = True
rstMovimento.Update

CommitTrans


End Function


Public Function CancelaMovimento(codMovimento As Integer)
'=================================================================================
'* Função               : CancelaMovimento
'
'* Principal objetivo   : Encontrar o movimento solicitado e cancelar sua
'                         operação junto ao estoque.
'
'* Tabelas envolvidas   : tblMovimentacao;tblMovimentacaoItens;tblProdutos.
'
'=================================================================================

Dim dbDados As DAO.Database
Dim rstMovimento As DAO.Recordset
Dim rstItensDoMovimento As DAO.Recordset
Dim rstProdutos As DAO.Recordset

Dim contRegistros As Long

Set dbDados = CurrentDb
Set rstMovimento = dbDados.OpenRecordset("Select * from tblMovimentacao where CodMov = " & codMovimento)

'##### TESTES DE VALIDAÇÃO DO MOVIMENTO ####

If rstMovimento.RecordCount > 0 Then
   If Not rstMovimento.Fields("mov_Ok") Then
        MsgBox "ATENÇÃO: Esta movimentação já foi efetivada!", vbExclamation
        Exit Function
   Else
        Set rstItensDoMovimento = dbDados.OpenRecordset("Select * from tblMovimentacaoItens where CodMov = " & codMovimento)
        Set rstProdutos = dbDados.OpenRecordset("Select * from tblProdutos")
   End If
Else
   MsgBox "ATENÇÃO: Movimento não encontrado!", vbCritical
   Exit Function
End If

'##### BAIXA DO ESTOQUE E EFETIVAÇÃO DO MOVIMENTO ####

BeginTrans

rstItensDoMovimento.MoveLast
rstItensDoMovimento.MoveFirst

If rstMovimento.Fields("mov_Tipo") = 1 Then
    For contRegistros = 1 To rstItensDoMovimento.RecordCount
    
       rstProdutos.MoveLast
       rstProdutos.FindFirst "CodPro = " & Str(rstItensDoMovimento.Fields("CodPro"))
       rstProdutos.Edit
       
       rstProdutos.Fields("pro_QTD") = rstProdutos.Fields("pro_QTD") - rstItensDoMovimento.Fields("mvi_QTD")
       rstProdutos.Update
       
       rstItensDoMovimento.MoveNext
    Next
Else
    For contRegistros = 1 To rstItensDoMovimento.RecordCount
    
       rstProdutos.MoveLast
       rstProdutos.FindFirst "CodPro = " & Str(rstItensDoMovimento.Fields("CodPro"))
       rstProdutos.Edit
       
       rstProdutos.Fields("pro_QTD") = rstProdutos.Fields("pro_QTD") + rstItensDoMovimento.Fields("mvi_QTD")
       rstProdutos.Update
       
       rstProdutos.MoveNext
    Next
End If

rstMovimento.Edit
rstMovimento.Fields("mov_Ok") = False
rstMovimento.Update

CommitTrans

End Function


Public Function GeraOrcamento(codOrcamento As Integer)

Dim Orcamento As DAO.Recordset
Dim ItensDoOrcamento As DAO.Recordset

Dim Cliente As DAO.Recordset
Dim Funcionario As DAO.Recordset
Dim Produtos As DAO.Recordset


Dim caminho As String

Set Orcamento = CurrentDb.OpenRecordset("Select * from tblOrcamentos where codOrc = " & codOrcamento)
Set ItensDoOrcamento = CurrentDb.OpenRecordset("Select * from tblOrcamentosItens where codOrc = " & Orcamento.Fields("codOrc"))

Set Cliente = CurrentDb.OpenRecordset("Select * from tblClientes where codCLI = " & Orcamento.Fields("codCLI"))
Set Funcionario = CurrentDb.OpenRecordset("Select * from tblFuncionarios where codFun = " & Orcamento.Fields("codFun"))
Set Produtos = CurrentDb.OpenRecordset("tblProdutos")

caminho = Application.CurrentProject.Path

'Shell "winword.exe"

'Dim i As Long

'For i = 1 To 7500
'    DoEvents
'Next


Documents.Add Template:=caminho & "\Orcamentos.dot", newtemplate:=False


'==============
'   Orçamento
'==============

    Selection.GoTo what:=wdGoToBookmark, Name:="bmNumeroDoOrcamento"
    Selection.TypeText Format(Orcamento.Fields("codOrc"), "000000")

    Selection.GoTo what:=wdGoToBookmark, Name:="bmDataDeEmissao"
    Selection.TypeText Format(Orcamento.Fields("orc_data"), "dd/mm/yyyy")

    Selection.GoTo what:=wdGoToBookmark, Name:="bmVendedor"
    Selection.TypeText Funcionario.Fields("fun_Descricao")

'====================
'   Dados do Cliente
'====================

    If Not IsNull(Cliente.Fields("cli_Descricao")) Then
        Selection.GoTo what:=wdGoToBookmark, Name:="bmCliente"
        Selection.TypeText Cliente.Fields("cli_Descricao")
    End If


    If Not IsNull(Orcamento.Fields("orc_AC")) Then
        Selection.GoTo what:=wdGoToBookmark, Name:="bmContato"
        Selection.TypeText Orcamento.Fields("orc_AC")
    End If


    If Not IsNull(Orcamento.Fields("orc_email")) Then
        Selection.GoTo what:=wdGoToBookmark, Name:="bmEMail"
        Selection.TypeText Orcamento.Fields("orc_email")
    End If

'====================
'   Pagamentos
'====================

    If Not IsNull(Orcamento.Fields("orc_Pgto1")) Then
        Selection.GoTo what:=wdGoToBookmark, Name:="bmPG_01"
        Selection.TypeText Orcamento.Fields("orc_Pgto1")
        
        If Orcamento.Fields("orc_Valor1") > 0 Then
           Selection.GoTo what:=wdGoToBookmark, Name:="bmVL_01"
           Selection.TypeText FormatNumber(Orcamento.Fields("orc_Valor1"))
        End If
    End If
    
    
    If Not IsNull(Orcamento.Fields("orc_Pgto2")) Then
        Selection.GoTo what:=wdGoToBookmark, Name:="bmPG_02"
        Selection.TypeText Orcamento.Fields("orc_Pgto2")
        
        If Orcamento.Fields("orc_Valor2") > 0 Then
           Selection.GoTo what:=wdGoToBookmark, Name:="bmVL_02"
           Selection.TypeText FormatNumber(Orcamento.Fields("orc_Valor2"))
        End If
    End If
    
    
    If Not IsNull(Orcamento.Fields("orc_Pgto3")) Then
        Selection.GoTo what:=wdGoToBookmark, Name:="bmPG_03"
        Selection.TypeText Orcamento.Fields("orc_Pgto3")
        
        If Orcamento.Fields("orc_Valor3") > 0 Then
           Selection.GoTo what:=wdGoToBookmark, Name:="bmVL_03"
           Selection.TypeText FormatNumber(Orcamento.Fields("orc_Valor3"))
        End If
    End If
    

'====================
'   Observação
'====================

    If Not IsNull(Orcamento.Fields("orc_Obs")) Then
        Selection.GoTo what:=wdGoToBookmark, Name:="bmObservacoes"
        Selection.TypeText Orcamento.Fields("orc_Obs")
    End If

'====================
'   Itens
'====================

Dim Inicio As Integer
Dim LimiteDeItens As Integer
Dim ContItens As Integer
Dim PuloDeLinha As Integer
'Dim SomaTotItens As Currency

LimiteDeItens = 24
PuloDeLinha = 2
'SomaTotItens = 0

ItensDoOrcamento.MoveLast
ItensDoOrcamento.MoveFirst
Inicio = 1
If Not ItensDoOrcamento.EOF Then

Do While Not ItensDoOrcamento.EOF

    'Especificações
    Produtos.MoveLast
    Produtos.FindFirst "codPro = " & ItensDoOrcamento.Fields("codPro")
    Selection.GoTo what:=wdGoToBookmark, Name:=IIf(Inicio < 10, "bmESP_0", "bmESP_") & Inicio
    Selection.TypeText Produtos.Fields("pro_Descricao")
    
    'Valor
    Selection.GoTo what:=wdGoToBookmark, Name:=IIf(Inicio < 10, "bmESP_VAL_0", "bmESP_VAL_") & Inicio
    Selection.TypeText FormatNumber(ItensDoOrcamento.Fields("orc_Valor"))
    
    'SomaTotItens = SomaTotItens + ItensDoOrcamento.Fields("orc_Valor")
    
    ItensDoOrcamento.MoveNext
    Inicio = Inicio + 1
        
Loop
    
    Selection.GoTo what:=wdGoToBookmark, Name:=IIf(Inicio < 10, "bmESP_0", "bmESP_") & Inicio + PuloDeLinha
    Selection.Font.Bold = True
    Selection.TypeText "Total : "
    
    'Valor
    Selection.GoTo what:=wdGoToBookmark, Name:=IIf(Inicio < 10, "bmESP_VAL_0", "bmESP_VAL_") & Inicio + PuloDeLinha
    Selection.Font.Bold = True
    Selection.TypeText FormatNumber(Orcamento.Fields("orc_Valor1")) 'FormatCurrency(SomaTotItens)
    
End If

Cliente.Close
Orcamento.Close
ItensDoOrcamento.Close

End Function


