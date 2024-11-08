VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu_Principal 
   ClientHeight    =   9195.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17985
   OleObjectBlob   =   "Menu_Principal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Menu_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Obrigatoriedade de declaração de variável
Option Explicit
Private Sub B_Exc_Vendas_Click()
On Error GoTo erro_carregamento:

    'Ativar Botões
        B_Home.BackStyle = fmBackStyleTransparent
        B_Produtos.BackStyle = fmBackStyleTransparent
        B_Vendas.BackStyle = fmBackStyleTransparent
        B_Exc_Vendas.BackStyle = fmBackStyleOpaque
        B_Lucro.BackStyle = fmBackStyleTransparent
        
        PAGINAS.Value = 3
Exit Sub
erro_carregamento:
End Sub
Sub Tela_Inicial()
On Error GoTo erro_carregamento:

    'Ativar Botões
        B_Home.BackStyle = fmBackStyleOpaque
        B_Produtos.BackStyle = fmBackStyleTransparent
        B_Vendas.BackStyle = fmBackStyleTransparent
        B_Exc_Vendas.BackStyle = fmBackStyleTransparent
        B_Lucro.BackStyle = fmBackStyleTransparent
        
        PAGINAS.Value = 0
Exit Sub
erro_carregamento:
End Sub
Private Sub B_Excluir_Click()
On Error GoTo erro_carregamento:
    
        If MsgBox("Deseja realmente continuar?", vbQuestion + vbYesNo, "Exclusão") = vbYes Then
            
           'Verifica se a linha esta vazia
           If P_Lista.Column(0) <> Empty Then
               With Sheets("PRODUTOS").Range("A:A")
                    'Localizando chave primaria na primeira coluna
                Set EncontrarID = .Find(What:=P_Lista.Column(0), _
                            LookAt:=xlWhole)
            
                    If Not EncontrarID Is Nothing Then
                    Application.Goto EncontrarID, True
                    
                    'Deletando linha selecionada
                    ActiveCell.Rows.EntireRow.Delete
                    
                        'Após deletar ele atualiza a lista de produtos
                         Call PreenchimentoProdutos
                         
                         
                         MsgBox "Exclusão realizada com sucesso!", vbInformation, "Exclusão"
                        
                    End If
                 End With
            Else
             MsgBox "Selecione um produto para exclusão!", vbExclamation, "Exclusão"
            End If
        End If
Exit Sub
erro_carregamento:
End Sub
Private Sub B_Home_Click()
On Error GoTo erro_carregamento:

        Call Tela_Inicial
        
Exit Sub
erro_carregamento:
End Sub
Private Sub B_Lucro_Click()
On Error GoTo erro_carregamento:

    'Ativar Botões
        B_Home.BackStyle = fmBackStyleTransparent
        B_Produtos.BackStyle = fmBackStyleTransparent
        B_Vendas.BackStyle = fmBackStyleTransparent
        B_Exc_Vendas.BackStyle = fmBackStyleTransparent
        B_Lucro.BackStyle = fmBackStyleOpaque
        
        PAGINAS.Value = 4
Exit Sub
erro_carregamento:
End Sub
Private Sub B_Produtos_Click()
    On Error GoTo erro_carregamento:
    'Ativar Botões
        B_Home.BackStyle = fmBackStyleTransparent
        B_Produtos.BackStyle = fmBackStyleOpaque
        B_Vendas.BackStyle = fmBackStyleTransparent
        B_Exc_Vendas.BackStyle = fmBackStyleTransparent
        B_Lucro.BackStyle = fmBackStyleTransparent
        
        PAGINAS.Value = 1
    
    Exit Sub
erro_carregamento:
End Sub
'Ao clicar no campo vendas ele fica selecionado
Private Sub B_Vendas_Click()
On Error GoTo erro_carregamento:

    'Ativar Botões
        B_Home.BackStyle = fmBackStyleTransparent
        B_Produtos.BackStyle = fmBackStyleTransparent
        B_Vendas.BackStyle = fmBackStyleOpaque
        B_Exc_Vendas.BackStyle = fmBackStyleTransparent
        B_Lucro.BackStyle = fmBackStyleTransparent
        
        PAGINAS.Value = 2
Exit Sub
erro_carregamento:
End Sub
'Ao clicar na barra de menu ele chama a tela inicial
Private Sub Barra_Menu_Click()
On Error GoTo erro_carregamento:
    
        Call Tela_Inicial
        
Exit Sub
erro_carregamento:
End Sub
Private Sub BEstImagem_Click()
On Error GoTo erro_carregamento:
    
        'Chamando variável para abrir caixa de diálogo
        Call FotoProduto
        
Exit Sub
erro_carregamento:
End Sub

 'Transformar e "travar" letras para Maiuscula
Private Sub E_busca_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
    
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
Exit Sub
erro_carregamento:
End Sub

Private Sub E_Busca_Change()

    Call PreenchimentoListaVendasExcluir
    
End Sub

Private Sub E_DataFinal_AfterUpdate()

        If E_DataFinal <> Empty Then
            'DIAS: se o numero de dias for maior que 31
            If Left(E_DataFinal, 2) > 31 Then
                'Ele limpa o campo
                E_DataFinal = Empty
                E_DataFinal.SetFocus
             Exit Sub
        
            'MÊS: se o numero de meses for maior que 12
            ElseIf Mid(E_DataFinal, 4, 2) > 12 Then
                'Ele limpa o campo
                E_DataFinal = Empty
                E_DataFinal.SetFocus
            Else
                'Verifica se a data é valida
                 If IsDate(E_DataFinal) = False Then
                     'Ele limpa o campo
                     E_DataFinal = Empty
                     E_DataFinal.SetFocus
                 End If
                 
                Call PreenchimentoListaVendasExcluir
                 
            End If
          End If
End Sub

 'Tabela Keyascii para numeros, barra e backspace(botão limpar)
Private Sub E_DataFinal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
        
        Select Case KeyAscii
            Case 8, 47 To 57
            Case Else
            KeyAscii = 0
        End Select
        
Exit Sub
erro_carregamento:
End Sub

Private Sub E_DataInicial_AfterUpdate()

        If E_DataInicial <> Empty Then
            'DIAS: se o numero de dias for maior que 31
            If Left(E_DataInicial, 2) > 31 Then
                'Ele limpa o campo
                E_DataInicial = Empty
                E_DataInicial.SetFocus
             Exit Sub
        
            'MÊS: se o numero de meses for maior que 12
            ElseIf Mid(E_DataInicial, 4, 2) > 12 Then
                'Ele limpa o campo
                E_DataInicial = Empty
                E_DataInicial.SetFocus
            Else
                'Verifica se a data é valida
                 If IsDate(E_DataInicial) = False Then
                     'Ele limpa o campo
                     E_DataInicial = Empty
                     E_DataInicial.SetFocus
                 End If
                 
                 Call PreenchimentoListaVendasExcluir
                 
            End If
          End If
End Sub

 'Tabela Keyascii para numeros, barra e backspace(botão limpar)
Private Sub E_DataInicial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
        
        Select Case KeyAscii
            Case 8, 47 To 57
            Case Else
            KeyAscii = 0
        End Select
        
Exit Sub
erro_carregamento:
End Sub
Private Sub E_Excluir_Click()
On Error GoTo erro_carregamento:

    If E_QtdeExcluir = Empty Or E_QtdeExcluir = 0 Then
        MsgBox "Informe uma quantidade para exclusão!", vbExclamation, "Exclusão"
        E_QtdeExcluir.SetFocus
    Exit Sub
    
    ElseIf MsgBox("Deseja realmente continuar?", vbQuestion + vbYesNo, "Exclusão") = vbYes Then
    
        'Estorno ao estoque
        Sheets("PRODUTOS").Select
        
            With Sheets("PRODUTOS").Range("A:A")
            'Localizando chave primaria
            Set EncontrarID = .Find(What:=E_Lista.Column(2), _
                            LookAt:=xlWhole)
            
                If Not EncontrarID Is Nothing Then
                    Application.Goto EncontrarID, True
                    
                    'Localizar coluna de quantidade e somar
                    ActiveCell.Offset(0, 5).Value = CDbl(ActiveCell.Offset(0, 5).Value) + CDbl(E_QtdeExcluir)
                End If
            End With
            
            'Exclusão
            If E_QtdeTotal.Value = True Then
                
                'Exclusão total
                Sheets("VENDAS FINALIZADAS").Select
            
                With Sheets("VENDAS FINALIZADAS").Range("A:A")
                'Localizando chave primaria
                Set EncontrarID = .Find(What:=E_Lista.Column(0), _
                                LookAt:=xlWhole)
                
                    If Not EncontrarID Is Nothing Then
                        Application.Goto EncontrarID, True
                        
                        'Deletar linha
                        ActiveCell.Rows.EntireRow.Delete
                    End If
                End With
            Else
                'Exclusão parcial ou total
                Sheets("VENDAS FINALIZADAS").Select
            
                With Sheets("VENDAS FINALIZADAS").Range("A:A")
                'Localizando chave primaria
                Set EncontrarID = .Find(What:=E_Lista.Column(0), _
                                LookAt:=xlWhole)
                
                    If Not EncontrarID Is Nothing Then
                        Application.Goto EncontrarID, True
                        
                        'Exclusão total
                        If E_QtdeFinal = Empty Or E_QtdeFinal = 0 Then
                        'Deletar linha
                        ActiveCell.Rows.EntireRow.Delete
                        'Exclusão Parcial
                        Else
                            ActiveCell.Offset(0, 6).Value = E_QtdeFinal
                        End If
                    End If
                End With
            End If
    End If
    
    
    'Limpar
    E_QtdeExcluir = Empty
    E_QtdeFinal = Empty
    
    'Atualizar formulários
    Call PreenchimentoListaVendasExcluir
    Call PreenchimentoVendas
    Call PreenchimentoProdutos
Exit Sub
erro_carregamento:
End Sub

Private Sub E_Lista_Click()
On Error GoTo erro_carregamento:
    
        Sheets("VENDAS FINALIZADAS").Select
    
    With Sheets("VENDAS FINALIZADAS").Range("A:A")
        'Localizando chave primaria
        Set EncontrarID = .Find(What:=E_Lista.Column(2), _
                        LookAt:=xlWhole)
        
            If Not EncontrarID Is Nothing Then
                Application.Goto EncontrarID, True
                
                'Seleção de controle automatico padrão
                E_QtdeTotal.Value = True
                'Limpa quantidade final
                E_QtdeFinal = Empty
                'Quantidade a excluir
                E_QtdeExcluir = E_Lista.Column(6)
                
                'Habilita parcial se a quantidade for superior a 1
                If E_Lista.Column(6) > 1 Then
                    E_QtdeParcial.Enabled = True
                Else
                    'Se não for maio que 1 ele não habilita
                    E_QtdeParcial.Enabled = False
                End If
                
                'Idenficar se a coluna tem endereço de imagem
                If ActiveCell.Offset(0, 12) <> "" Then
                    E_Imagem.Picture = LoadPicture(ActiveCell.Offset(0, 12))
                    'Formatando imagem
                    E_Imagem.PictureSizeMode = fmPictureSizeModeZoom
                    'Deixando imagem visivel
                    E_Imagem.Visible = True
                  Else
                    'Caso não tiver endereço de imagem ele fica oculto
                     E_Imagem.Visible = False
                End If
            End If
     End With
Exit Sub
erro_carregamento:
End Sub

Private Sub E_QtdeExcluir_Change()

    'Se o campo estiver vazio, ele limpa
    If E_QtdeExcluir = Empty Then
        E_QtdeFinal = Empty
    Else
        'Não permiti excluir quantidade maior que a vendida
        If CDbl(E_QtdeExcluir) > CDbl(E_Lista.Column(6)) Or E_QtdeExcluir = 0 Then
            E_QtdeExcluir = Empty
            E_QtdeFinal = Empty
        Else
            'Devolução parcial
            E_QtdeFinal = CDbl(E_Lista.Column(6)) - CDbl(E_QtdeExcluir)
        
        End If
    End If
End Sub


Private Sub E_QtdeParcial_Click()

    'Habilita a quantidade se for maior que 1
    If E_QtdeParcial.Value = True Then
        E_QtdeExcluir.Enabled = True
        E_QtdeExcluir.SetFocus
    End If
End Sub

Private Sub E_QtdeTotal_Click()
On Error GoTo erro_carregamento:

        'So permiti digitação se for maior que 1 na quantidade parcial
        If E_QtdeTotal.Value = True Then
            E_QtdeExcluir = E_Lista.Column(6)
            E_QtdeFinal = Empty
            E_QtdeExcluir.Enabled = False
        End If
Exit Sub
erro_carregamento:
End Sub

Private Sub L_Busca_Change()
On Error GoTo erro_carregamento:

    Call PreenchimentoListaVendasLucro
    
Exit Sub
erro_carregamento:
End Sub

 'Transformar e "travar" letras para Maiuscula
Private Sub L_busca_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
    
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
Exit Sub
erro_carregamento:
End Sub

Private Sub L_DataFinal_AfterUpdate()

        If L_DataFinal <> Empty Then
            'DIAS: se o numero de dias for maior que 31
            If Left(L_DataFinal, 2) > 31 Then
                'Ele limpa o campo
                L_DataFinal = Empty
                L_DataFinal.SetFocus
             Exit Sub
        
            'MÊS: se o numero de meses for maior que 12
            ElseIf Mid(L_DataFinal, 4, 2) > 12 Then
                'Ele limpa o campo
                L_DataFinal = Empty
                L_DataFinal.SetFocus
            Else
                'Verifica se a data é valida
                 If IsDate(L_DataFinal) = False Then
                     'Ele limpa o campo
                     L_DataFinal = Empty
                     L_DataFinal.SetFocus
                 End If
                 
                Call PreenchimentoListaVendasLucro
                 
            End If
          End If
End Sub

 'Tabela Keyascii para numeros, barra e backspace(botão limpar)
Private Sub L_DataFinal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
        
        Select Case KeyAscii
            Case 8, 47 To 57
            Case Else
            KeyAscii = 0
        End Select
        
Exit Sub
erro_carregamento:
End Sub

Private Sub L_DataInicial_AfterUpdate()

        If L_DataInicial <> Empty Then
            'DIAS: se o numero de dias for maior que 31
            If Left(L_DataInicial, 2) > 31 Then
                'Ele limpa o campo
                L_DataInicial = Empty
                L_DataInicial.SetFocus
             Exit Sub
        
            'MÊS: se o numero de meses for maior que 12
            ElseIf Mid(L_DataInicial, 4, 2) > 12 Then
                'Ele limpa o campo
                L_DataInicial = Empty
                L_DataInicial.SetFocus
            Else
                'Verifica se a data é valida
                 If IsDate(L_DataInicial) = False Then
                     'Ele limpa o campo
                     L_DataInicial = Empty
                     L_DataInicial.SetFocus
                 End If
                 
                 Call PreenchimentoListaVendasLucro
                 
            End If
          End If
End Sub

 'Tabela Keyascii para numeros, barra e backspace(botão limpar)
Private Sub L_DataInicial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
        
        Select Case KeyAscii
            Case 8, 47 To 57
            Case Else
            KeyAscii = 0
        End Select
        
Exit Sub
erro_carregamento:
End Sub

Private Sub L_Sair_Click()
On Error GoTo erro_carregamento:
        
        'Fechando o sistema
        If MsgBox("Deseja realmente continuar com o fechamento do sistema?", vbQuestion + vbYesNo, "Fechar") = vbYes Then
        
            Application.Quit
            ActiveWorkbook.Close savechanges:=False
            'Não emitir nenhum alerta
            Application.DisplayAlerts = False
            
        End If
Exit Sub
erro_carregamento:
End Sub

Private Sub P_Busca_Change()
On Error GoTo erro_carregamento:
        
        'Chamando variável para carregar produtos
        Call PreenchimentoProdutos
        
Exit Sub
erro_carregamento:
End Sub
'Transformar e "travar" letras para Maiuscula
Private Sub P_busca_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
    
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
Exit Sub
erro_carregamento:
End Sub
'Trava preenchimento obrigatorio / Configurando campos obrigatorios *
Private Sub P_Cadastrar_Click()
On Error GoTo erro_carregamento:

        'Se este campo estiver vazio
        If P_Produto = Empty Then
            'Então, apareça essa msg
            MsgBox "Informe o nome do produto!", vbExclamation, "Cadastro"
            'Colorir o campo vazio
            P_Produto.BackColor = &HFFFF&
            'Após msg, campo será focado
            P_Produto.SetFocus
        Exit Sub
        
  'Só após o preenchido segue para o próximo campo
  
        'Se este campo estiver vazio
        ElseIf P_Qtde = Empty Then
            'Então, apareça essa msg
            MsgBox "Informe a quantidade!", vbExclamation, "Cadastro"
            'Colorir o campo vazio
            P_Qtde.BackColor = &HFFFF&
            'Após msg, campo será focado
            P_Qtde.SetFocus
                        
         Exit Sub
                
  'Só após o preenchido segue para o próximo campo
  
        'Se este campo estiver vazio
        ElseIf P_ValorCompra = Empty Then
            'Então, apareça essa msg
            MsgBox "Informe o valor de compra!", vbExclamation, "Cadastro"
            'Colorir o campo vazio
            P_ValorCompra.BackColor = &HFFFF&
            'Após msg, campo será focado
            P_ValorCompra.SetFocus
        
        Exit Sub
                
   'Só após o preenchido segue para o próximo campo
   
        'Se este campo estiver vazio
        ElseIf P_ValorVenda = Empty Then
            'Então, apareça essa msg
            MsgBox "Informe valor de venda!", vbExclamation, "Cadastro"
            'Colorir o campo vazio
            P_ValorVenda.BackColor = &HFFFF&
            'Após msg, campo será focado
            P_ValorVenda.SetFocus
            
         Exit Sub
            'Após todos os campos preenchidos confirmar cadastro
        ElseIf MsgBox("Deseja realmente continuar?", vbQuestion + vbYesNo, "Cadastro") = vbYes Then
        
            'Selecionar planilha
            Sheets("PRODUTOS").Select
            'Referencia a uma célula especifica "ultima celula do excel e primeira coluna"
            Application.Goto reference:="R1048576C1"
            'Ultima linha preenchida
            Selection.End(xlUp).Select
            'Localizar primeira linha vazia(descer uma linha e manter na mesma coluna)
            ActiveCell.Offset(1, 0).Select
            
            ContCodigo = Range("AR1").Value + 1
            
            
            'Lançamento de dados
            ActiveCell.Value = ContCodigo
            
            'Manter na linha atual (0) e andar duas colunas (2)
            ActiveCell.Offset(0, 2).Value = Format(P_CodBarras)
            ActiveCell.Offset(0, 3).Value = Format(P_Produto)
            ActiveCell.Offset(0, 4).Value = Format(P_Marca)
            ActiveCell.Offset(0, 5).Value = Format(P_Qtde)
            ActiveCell.Offset(0, 6).Value = Format(P_UnidadeMedida)
            
            'Formatar em moeda
            ActiveCell.Offset(0, 7).Value = Format(P_ValorCompra, "Currency")
            ActiveCell.Offset(0, 8).Value = Format(P_ValorVenda, "Currency")
            
            'Formatar como data
            ActiveCell.Offset(0, 9).Value = CDate(P_DataCompra)
            ActiveCell.Offset(0, 10).Value = CDate(P_DataValidade)
   
            'Formatado
            ActiveCell.Offset(0, 11).Value = Format(P_Outras)
            
            'Salvar endereço da imagem selecionada
            ActiveCell.Offset(0, 12).Value = SalvarFoto
            
            'Limpar para o próximo produto
            P_CodBarras = Empty
            P_Produto = Empty
            P_Marca = Empty
            P_Qtde = Empty
            P_UnidadeMedida = Empty
            P_ValorCompra = Empty
            P_ValorVenda = Empty
            P_DataCompra = Empty
            P_DataValidade = Empty
            P_Outras = Empty
            SalvarFoto = Empty
            
            Range("AR1").Value = ContCodigo
            
            'Classificar em ordem alfabetica
            Call Classificar_Produtos
            
            
            MsgBox "Cadastro realizado com sucesso!", vbInformation, "Cadastro"
            
        End If
Exit Sub
erro_carregamento:
End Sub
'Bloquear para aceitar apenas numero
Private Sub P_codbarras_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
        
        'Tabela Keyascii para numeros e backspace(botão limpar)
        Select Case KeyAscii
            Case 8, 47 To 57
            Case Else
            KeyAscii = 0
        End Select
        
Exit Sub
erro_carregamento:
End Sub
'Ao sair do campo ele limpa se for invalido
Private Sub P_DataCompra_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo erro_carregamento:
        
        If P_DataCompra <> Empty Then
            'DIAS: se o numero de dias for maior que 31
            If Left(P_DataCompra, 2) > 31 Then
                'Ele limpa o campo
                P_DataCompra = Empty
                P_DataCompra.SetFocus
             Exit Sub
        
            'MÊS: se o numero de meses for maior que 12
            ElseIf Mid(P_DataCompra, 4, 2) > 12 Then
                'Ele limpa o campo
                P_DataCompra = Empty
                P_DataCompra.SetFocus
            Else
                'Verifica se a data é valida
                 If IsDate(P_DataCompra) = False Then
                     'Ele limpa o campo
                     P_DataCompra = Empty
                     P_DataCompra.SetFocus
                 End If
            End If
          End If
Exit Sub
erro_carregamento:
End Sub
 'Tabela Keyascii para numeros, barra e backspace(botão limpar)
Private Sub P_DataCompra_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
        
        Select Case KeyAscii
            Case 8, 47 To 57
            Case Else
            KeyAscii = 0
        End Select
        
Exit Sub
erro_carregamento:
End Sub
    'Ao sair do campo ele limpa se for invalido
Private Sub P_DataValidade_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo erro_carregamento:
        
        If P_DataValidade <> Empty Then
            'DIAS: se o numero de dias for maior que 31
            If Left(P_DataValidade, 2) > 31 Then
                'Ele limpa o campo
                P_DataValidade = Empty
                P_DataValidade.SetFocus
             Exit Sub
        
            'MÊS: se o numero de meses for maior que 12
            ElseIf Mid(P_DataValidade, 4, 2) > 12 Then
                'Ele limpa o campo
                P_DataValidade = Empty
                P_DataValidade.SetFocus
            Else
                'Verifica se a data é valida
                 If Not IsDate(P_DataValidade) Then
                     'Ele limpa o campo
                     P_DataValidade = Empty
                     P_DataValidade.SetFocus
                 End If
            End If
          End If
Exit Sub
erro_carregamento:
End Sub
'Tabela Keyascii para numeros, barra e backspace(botão limpar)
Private Sub P_DataValidade_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
        
        Select Case KeyAscii
            Case 8, 47 To 57
            Case Else
            KeyAscii = 0
        End Select
        
Exit Sub
erro_carregamento:
End Sub
'Carregando imagem dos produtos no formulario
Private Sub P_Lista_Click()
On Error GoTo erro_carregamento:
    
        Sheets("PRODUTOS").Select
    
    With Sheets("PRODUTOS").Range("A:A")
        'Localizando chave primaria na primeira coluna
        Set EncontrarID = .Find(What:=P_Lista.Column(0), _
                        LookAt:=xlWhole)
        
            If Not EncontrarID Is Nothing Then
                Application.Goto EncontrarID, True
                    'Idenficar se a coluna tem endereço de imagem
                    If ActiveCell.Offset(0, 12) <> "" Then
                        P_Imagem.Picture = LoadPicture(ActiveCell.Offset(0, 12))
                        'Formatando imagem
                        P_Imagem.PictureSizeMode = fmPictureSizeModeZoom
                        'Deixando imagem visivel
                        P_Imagem.Visible = True
                      Else
                        'Caso não tiver endereço de imagem ele fica oculto
                         P_Imagem.Visible = False
                    End If
            End If
     End With
Exit Sub
erro_carregamento:
End Sub
'Ao cliclar duas vezes permiti fazer alteração
Private Sub P_Lista_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo erro_carregamento:
    
        Alterar_Produtos.P_CodBarras = P_Lista.Column(2)
        Alterar_Produtos.P_Produto = P_Lista.Column(3)
        Alterar_Produtos.P_Marca = P_Lista.Column(4)
        Alterar_Produtos.P_Qtde = P_Lista.Column(5)
        Alterar_Produtos.P_UnidadeMedida = P_Lista.Column(6)
        Alterar_Produtos.P_ValorCompra = P_Lista.Column(7)
        Alterar_Produtos.P_ValorVenda = P_Lista.Column(8)
        Alterar_Produtos.P_DataCompra = P_Lista.Column(9)
        Alterar_Produtos.P_DataValidade = P_Lista.Column(10)
        Alterar_Produtos.P_Outras = P_Lista.Column(11)
        
        'Verifica se coluna imagem esta vazia
         If P_Lista.Column(12) <> Empty Then
            
            'Carregar imagem
            Alterar_Produtos.P_Imagem.Picture = LoadPicture(P_Lista.Column(12))
            'Formatar tamanho da imagem
            Alterar_Produtos.P_Imagem.PictureSizeMode = fmPictureSizeModeZoom
            'Tornar imagem visivel
            Alterar_Produtos.P_Imagem.Visible = True
        
         End If
        
        'Apresenta o produto após o carregamento
        Alterar_Produtos.Show
        
Exit Sub
erro_carregamento:
End Sub
'Transformar e "travar" letras para  Maiuscula
Private Sub P_marca_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
    
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
Exit Sub
erro_carregamento:
End Sub
'Transformar e "travar" letras para Maiuscula
Private Sub P_outras_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
    
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
Exit Sub
erro_carregamento:
End Sub
'Após o preenchimendo do campo ele retorna branco
Private Sub P_Produto_Change()
On Error GoTo erro_carregamento:

    P_Produto.BackColor = &HFFFFFF
    
Exit Sub
erro_carregamento:
End Sub
 'Transformar e "travar" letras para Maiuscula
Private Sub P_Produto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
    
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
Exit Sub
erro_carregamento:
End Sub
'Após o preenchimendo do campo ele retorna branco
Private Sub P_Qtde_Change()
On Error GoTo erro_carregamento:

    P_Qtde.BackColor = &HFFFFFF
    
Exit Sub
erro_carregamento:
End Sub
'Bloquear para aceitar apenas numero
Private Sub P_Qtde_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
        
        'Tabela Keyascii para numeros e backspace(botão limpar)
        Select Case KeyAscii
            Case 8, 48 To 57
            Case Else
            KeyAscii = 0
        End Select
        
Exit Sub
erro_carregamento:
End Sub
Private Sub P_SCodigo_Click()
On Error GoTo erro_carregamento:

        'Esvazia o campo
        P_Busca = Empty
        'Campo fica selecionado
        P_Busca.SetFocus
        
Exit Sub
erro_carregamento:
End Sub
Private Sub P_SProduto_Click()
On Error GoTo erro_carregamento:
        
        'Esvazia o campo
        P_Busca = Empty
        'Campo fica selecionado
        P_Busca.SetFocus
        
Exit Sub
erro_carregamento:
End Sub
Private Sub P_ValorCompra_Change()
On Error GoTo erro_carregamento:
        
        'Após o preenchimendo do campo ele retorna branco
        P_ValorCompra.BackColor = &HFFFFFF
        
        'Bloquear virgula a esqueda
        If Left(P_ValorCompra, 1) = "," Then
            P_ValorCompra = Empty
        End If
        
Exit Sub
erro_carregamento:
End Sub
'Ao cliclar no campo ele desfaz a formatação em moeda
Private Sub P_ValorCompra_Enter()
On Error GoTo erro_carregamento:
        
        P_ValorCompra = Format(P_ValorCompra)
        
Exit Sub
erro_carregamento:
End Sub
'Ao sair do campo, formatar valor em moeda
Private Sub P_ValorCompra_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo erro_carregamento:
        
        P_ValorCompra = Format(P_ValorCompra, "Currency")
        
Exit Sub
erro_carregamento:
End Sub
'Bloquear para aceitar apenas numero
Private Sub P_valorcompra_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
        
        'Tabela Keyascii para numeros, virgula e backspace(botão limpar)
        Select Case KeyAscii
            Case 8, 44, 48 To 57
            
                'Permitir apenas uma virgula no campo
                If KeyAscii = 44 Then
                    If InStr(1, P_ValorCompra, ",", vbTextCompare) > 1 Then
                    KeyAscii = 0
                    End If
                End If
            Case Else
            KeyAscii = 0
        End Select
        
Exit Sub
erro_carregamento:
End Sub
Private Sub P_Valorvenda_Change()
On Error GoTo erro_carregamento:
        
        'Após o preenchimendo do campo ele retorna branco
        P_ValorVenda.BackColor = &HFFFFFF
        
        'Bloquear virgula a esqueda
        If Left(P_ValorVenda, 1) = "," Then
            P_ValorVenda = Empty
        End If
        
Exit Sub
erro_carregamento:
End Sub
'Ao cliclar no campo ele desfaz a formatação em moeda
Private Sub P_Valorvenda_Enter()
On Error GoTo erro_carregamento:
        
        P_ValorVenda = Format(P_ValorVenda)
        
Exit Sub
erro_carregamento:
End Sub
'Ao sair do campo, formatar valor em moeda
Private Sub P_Valorvenda_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo erro_carregamento:
        
        P_ValorVenda = Format(P_ValorVenda, "Currency")
        
Exit Sub
erro_carregamento:
End Sub
'Bloquear para aceitar apenas numero
Private Sub P_valorvenda_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
        
        'Tabela Keyascii para numeros, virgula e backspace(botão limpar)
        Select Case KeyAscii
            Case 8, 44, 48 To 57
            
                'Permitir apenas uma virgula no campo
                If KeyAscii = 44 Then
                    If InStr(1, P_ValorVenda, ",", vbTextCompare) > 1 Then
                    KeyAscii = 0
                    End If
                End If
            Case Else
            KeyAscii = 0
        End Select
        
Exit Sub
erro_carregamento:
End Sub
Private Sub Texto_Menu1_Click()
On Error GoTo erro_carregamento:
        Call Tela_Inicial
        
Exit Sub
erro_carregamento:
End Sub
Private Sub Texto_Menu2_Click()
On Error GoTo erro_carregamento:
        Call Tela_Inicial
        
Exit Sub
erro_carregamento:
End Sub
'Ocultar menu(botões)da pagina inicial
Private Sub UserForm_Initialize()
On Error GoTo erro_carregamento:

       'Carregar automaticamente
       Call PreenchimentoProdutos
       Call PreenchimentoVendasProdutos
       Call PreenchimentoVendas
       
       'Carregar data
       L_DataInicial = Format(Now - 3, "dd/mm/yyyy")
       L_DataFinal = Format(Now + 3, "dd/mm/yyyy")
       
       Call PreenchimentoListaVendasLucro
              
       'Carregar data
       E_DataInicial = Format(Now - 3, "dd/mm/yyyy")
       E_DataFinal = Format(Now + 3, "dd/mm/yyyy")
       
       Call PreenchimentoListaVendasExcluir
       
       PAGINAS.Style = fmTabStyleNone
       
       'Carregar combobox
       'Adiciona uma lista que contem o que esta descrito em AddItem
        With Me.P_UnidadeMedida
            .AddItem "PEÇA"
            .AddItem "UNIDADE"
            .AddItem "PAR"
        End With
        
        With Me.V_FormaPagto
            .AddItem "À VISTA"
            .AddItem "À PRAZO"
            .AddItem "CARTÃO"
        End With
        
Exit Sub
erro_carregamento:
End Sub
Private Sub V_Adicionar_Click()
On Error GoTo erro_carregamento:

    If V_Produto = Empty Then
        MsgBox "Selecione um produto para adição!", vbExclamation, "Venda"
    Exit Sub
    
    'Bloquear quando estoque estiver zerado, transformar em numero para leitura certa
    ElseIf CDbl(V_Qtde) > CDbl(V_Lista.Column(5)) Then
        MsgBox "Quantidade solicitada é maior que o estoque atual!", vbExclamation, "Venda"
            
    Else
            'Selecionar planilha
            Sheets("LISTA DE VENDAS").Select
            'Referencia a uma célula especifica "ultima celula do excel e primeira coluna"
            Application.Goto reference:="R1048576C1"
            'Ultima linha preenchida
            Selection.End(xlUp).Select
            'Localizar primeira linha vazia(descer uma linha e manter na mesma coluna)
            ActiveCell.Offset(1, 0).Select
            
            'localizar valor da primeira chave primaria e adicionar mais 1
            ContCodigo = Range("AR1").Value + 1
            
            'Colocando chave primaria na primeira celula ativa
            ActiveCell.Value = ContCodigo
            'Adicionando daodos e formatando
            ActiveCell.Offset(0, 1).Value = V_Lista.Column(0)
            'Código de barras
            ActiveCell.Offset(0, 2).Value = Format(V_Lista.Column(2))
            'Produto
            ActiveCell.Offset(0, 3).Value = Format(V_Lista.Column(3))
            'Marca
            ActiveCell.Offset(0, 4).Value = Format(V_Lista.Column(4))
            'Quantidade, consegue alterar quantidade
            ActiveCell.Offset(0, 5).Value = Format(V_Qtde)
            'Unidade de Medida
            ActiveCell.Offset(0, 6).Value = Format(V_Lista.Column(6))
            'Valor de Venda, consegue alterar valor
            ActiveCell.Offset(0, 7).Value = Format(V_ValorVenda, "currency")
            'Valor total de venda, transformando informações em número décimal
            ActiveCell.Offset(0, 8).Value = Format(CDbl(V_Qtde) * CDbl(V_ValorVenda), "currency")
            'Lucro
            ActiveCell.Offset(0, 9).Value = Format(CDbl(ActiveCell.Offset(0, 8).Value) - (CDbl(V_Qtde) * CDbl(V_Lista.Column(7))), "currency")
            
            'Salva dentro da planilha o código primário
            Range("AR1").Value = ContCodigo
            
            'Baixa de estoque na planilha produtos
            Sheets("PRODUTOS").Select
        
            With Sheets("PRODUTOS").Range("A:A")
            'Localizando chave primaria na primeira coluna
            Set EncontrarID = .Find(What:=V_Lista.Column(0), _
                            LookAt:=xlWhole)
            
                If Not EncontrarID Is Nothing Then
                    Application.Goto EncontrarID, True
                    
                    'Deduzir quantidade na planilha, cdbl converter para décimal
                    ActiveCell.Offset(0, 5).Value = CDbl(ActiveCell.Offset(0, 5).Value) - CDbl(V_Qtde)

                End If
            End With
            
            'Carregar listbox
            Call PreenchimentoVendas
            Call PreenchimentoProdutos
            'Carregar produtos na tabelinha amarela
            Call PreenchimentoVendasProdutos
            
    End If
    
Exit Sub
erro_carregamento:
End Sub
Private Sub V_Busca_Change()
On Error GoTo erro_carregamento:
    'Chamando variável para carregar produtos
    Call PreenchimentoVendas
    
Exit Sub
erro_carregamento:
End Sub

 'Transformar e "travar" letras para Maiuscula
Private Sub v_busca_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
    
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
Exit Sub
erro_carregamento:
End Sub

Private Sub V_Excluir_Click()
On Error GoTo erro_carregamento:
    
    'Estorno ao estoque
    Sheets("PRODUTOS").Select
    
        With Sheets("PRODUTOS").Range("A:A")
        'Localizando chave primaria do produto selecionado
        Set EncontrarID = .Find(What:=V_ListaProdutos.Column(1), _
                        LookAt:=xlWhole)
        
            If Not EncontrarID Is Nothing Then
                'Localizar chave primária
                Application.Goto EncontrarID, True
                
                ActiveCell.Offset(0, 5) = CDbl(ActiveCell.Offset(0, 5)) + CDbl(V_ListaProdutos.Column(5))
            End If
    End With
        
    'Excluir produto
    Sheets("LISTA DE VENDAS").Select
    
        With Sheets("LISTA DE VENDAS").Range("A:A")
        'Localizando chave primaria do produto selecionado
        Set EncontrarID = .Find(What:=V_ListaProdutos.Column(0), _
                        LookAt:=xlWhole)
        
            If Not EncontrarID Is Nothing Then
                'Localizar chave primária
                Application.Goto EncontrarID, True
                'Apaga toda a linha que foi selecionada
                ActiveCell.Rows.EntireRow.Delete
            End If
    End With
        
        'Limpar listbox
        V_ListaProdutos.Clear
        
            'Carregar listbox
            Call PreenchimentoVendas
            Call PreenchimentoProdutos
            'Carregar produtos na tabelinha amarela
            Call PreenchimentoVendasProdutos
        
Exit Sub
erro_carregamento:
End Sub

Private Sub V_ExcluirTudo_Click()
On Error GoTo erro_carregamento:
    
    Sheets("LISTA DE VENDAS").Select
    Range("A2").Select
    
    Do Until ActiveCell.Value = Empty
    
    ChaveProduto = ActiveCell.Offset(0, 1).Value
    QtdeProduto = ActiveCell.Offset(0, 5).Value
        'Estorno ao estoque
        Sheets("PRODUTOS").Select
        
            With Sheets("PRODUTOS").Range("A:A")
            'Localizando chave primaria do produto selecionado
            Set EncontrarID = .Find(What:=ChaveProduto, _
                            LookAt:=xlWhole)
            
                If Not EncontrarID Is Nothing Then
                    'Localizar chave primária
                    Application.Goto EncontrarID, True
                    
                    ActiveCell.Offset(0, 5) = CDbl(ActiveCell.Offset(0, 5)) + CDbl(QtdeProduto)
                End If
        End With
            
        'Excluir produto
        Sheets("LISTA DE VENDAS").Select
        'Exclui linha ativa
        ActiveCell.Rows.EntireRow.Delete
            
        'Limpar listbox
        V_ListaProdutos.Clear
            
     Loop
     
        'Renovar chave primária do zero
        Sheets("LISTA DE VENDAS").Range("AR1").Value = 0
     
        'Carregar listbox
        Call PreenchimentoVendas
        Call PreenchimentoProdutos
        'Carregar produtos na tabelinha amarela
        Call PreenchimentoVendasProdutos
      
Exit Sub
erro_carregamento:
End Sub

Private Sub V_Finalizar_Click()
On Error GoTo erro_carregamento:

ContProxLinha = 2

    'Verificar se exixte produtos na aba de vendas
    If V_QtdeItens = 0 Or V_QtdeItens = Empty Then
        MsgBox "Adicione um produto para venda!", vbExclamation, "Venda"
    Exit Sub
    
    'Verificar forma de pagamento
    ElseIf Me.V_FormaPagto = Empty Then
        'Abrir opções de pagamento
        V_FormaPagto.DropDown
        MsgBox "Selecione uma forma de pagamento!", vbExclamation, "Venda"
    Exit Sub
        
     ElseIf Me.V_TipoPagto = Empty Then
        'Abrir tipos de pagamento
        V_TipoPagto.DropDown
        MsgBox "Selecione um tipo de pagamento!", vbExclamation, "Venda"
     Exit Sub
     
     ElseIf MsgBox("Deseja realmente continuar?", vbQuestion + vbYesNo, "Venda") = vbYes Then
     
        On Error Resume Next
        'Seleciona planilha para lançar as vendas
        Sheets("LISTA DE VENDAS").Select
        Range("B2").Select
        
        'Concatenando código para gerar numero de venda
        SalvarNumeroVenda = (Format(Now, "yyyy") & Format(1, "0000")) + Range("AS1").Value
        
        Do Until ActiveCell.Value = Empty
            
            Range("B" & ContProxLinha & ": J" & ContProxLinha).Select
            
            'Fazendo cópia da venda na planilha selecionada
            Selection.Copy
            
            'Vai para planilha onde deseja colar
            Sheets("VENDAS FINALIZADAS").Select
            
            'Referencia a uma célula especifica "ultima celula do excel e primeira coluna"
            Application.Goto reference:="R1048576C1"
            'Ultima linha preenchida
            Selection.End(xlUp).Select
            'Localizar primeira linha vazia(desce uma linha e manter na mesma coluna)
            ActiveCell.Offset(1, 0).Select
            
            'localizar valor da primeira chave primaria e adicionar mais 1
            ContCodigo = Range("AR1").Value + 1
            
            ActiveCell.Value = ContCodigo
            'Selecionar duas colunas pra frente
            ActiveCell.Offset(0, 2).Select
            'Colando
            ActiveSheet.Paste
            'Inserindo numero da venda na coluna
            ActiveCell.Offset(0, 9).Value = SalvarNumeroVenda
            'Inserindo data da venda
            ActiveCell.Offset(0, 10).Value = Format(Now, "mm/dd/yyyy")
            'Inserindo forma de pagamento
            ActiveCell.Offset(0, 11).Value = Me.V_FormaPagto
            'Inserindo tipo de pagamento
            ActiveCell.Offset(0, 12).Value = Me.V_TipoPagto
            
            'Renovar chave primária
            Range("AR1").Value = ContCodigo
            
            'Selecionar planilha
            Sheets("LISTA DE VENDAS").Select
            'Fazer copia na próxima linha vazia
            ActiveCell.Offset(1, 0).Select
            'Adicionar linha
            ContProxLinha = ContProxLinha + 1
        Loop
        
            'Seleciona planilha para limpeza
            Sheets("LISTA DE VENDAS").Select
            UltimaLinha = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
            Range("A2:J" & UltimaLinha).Rows.EntireRow.Delete
            
            'Renovar chave primária
            Range("AR1").Value = 0
            'Renovar contagem numero de venda
            Range("AS1").Value = Range("AS1").Value + 1
            
            'Limpar listbox
            V_ListaProdutos.Clear
            Call PreenchimentoVendasProdutos
            
            
            MsgBox "Venda nº " & SalvarNumeroVenda & " finalizada com sucesso!", vbInformation, "Vendas"
            Err
            
    End If
Exit Sub
erro_carregamento:
End Sub

Private Sub V_FormaPagto_Change()
On Error GoTo erro_carregamento:
            
     'Encadeamento na forma de pagamento à vista
     If V_FormaPagto = "À VISTA" Then
        'Carregar o segundo combobox
        With Me.V_TipoPagto
            'Limpar
            .Clear
            .AddItem "DINHEIRO"
            .AddItem "DOC"
            .AddItem "TED"
        End With
      End If
            
     'Encadeamento na forma de pagamento à prazo
     If V_FormaPagto = "À PRAZO" Then
        'Carregar o segundo combobox
        With Me.V_TipoPagto
            'Limpar
            .Clear
            .AddItem "FIADO"
            .AddItem "BOLETO"
            .AddItem "CARNÊ"
        End With
      End If
      
     'Encadeamento na forma de pagamento cartão
     If V_FormaPagto = "CARTÃO" Then
        'Carregar o segundo combobox
        With Me.V_TipoPagto
            'Limpar
            .Clear
            .AddItem "DÉBITO - VISA"
            .AddItem "DÉBITO - MASTER"
            .AddItem "DÉBITO - ELO"
            .AddItem "CRÉDITO - VISA"
            .AddItem "CRÉDITO - MASTER"
            .AddItem "CRÉDITO - ELO"
        End With
      End If
Exit Sub
erro_carregamento:
End Sub

'Carregando imagem dos produtos no formulario
Private Sub V_Lista_Click()
On Error GoTo erro_carregamento:
    
        Sheets("PRODUTOS").Select
    
        With Sheets("PRODUTOS").Range("A:A")
        'Localizando chave primaria na primeira coluna
        Set EncontrarID = .Find(What:=V_Lista.Column(0), _
                        LookAt:=xlWhole)
        
            If Not EncontrarID Is Nothing Then
                'Localizar chave primária
                Application.Goto EncontrarID, True
                
                'Se a chave primaria estiver vazia, ele faz a limpeza dos campos
                If V_Lista.Column(0) = Empty Then
                    V_Produto = Empty
                    V_Qtde = Empty
                    V_UnidadeMedida = Empty
                    V_ValorVenda = Empty
                    
                    'Caso não tiver endereço de imagem ele fica oculto
                    V_Imagem.Visible = False
                Else
                    'Ao ecnontrar o produto ele carrega as informações
                    V_Produto = V_Lista.Column(3)
                    'Quantidade padrão
                    V_Qtde = 1
                    V_UnidadeMedida = V_Lista.Column(6)
                    V_ValorVenda = V_Lista.Column(8)
                    
                        'Idenfica se a coluna tem endereço de imagem
                        If ActiveCell.Offset(0, 12) <> "" Then
                            V_Imagem.Picture = LoadPicture(ActiveCell.Offset(0, 12))
                            'Ajusta a imagem
                            V_Imagem.PictureSizeMode = fmPictureSizeModeZoom
                            'Deixando imagem visivel
                            V_Imagem.Visible = True
                          Else
                            'Caso não tiver endereço de imagem ele fica oculto
                             V_Imagem.Visible = False
                        End If
                    End If
              Else
                'Se a chave primaria estiver em branco, ele faz a limpeza dos campos
                 V_Produto = Empty
                 V_Qtde = Empty
                 V_UnidadeMedida = Empty
                 V_ValorVenda = Empty
                 'Caso não tiver endereço de imagem ele fica oculto
                 V_Imagem.Visible = False
            End If
    End With
Exit Sub
erro_carregamento:
End Sub



    'Bloquear para aceitar apenas numero
Private Sub V_Qtde_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
        
        'Tabela Keyascii para numeros e backspace(botão limpar)
        Select Case KeyAscii
            Case 8, 48 To 57
            Case Else
            KeyAscii = 0
        End Select
Exit Sub
erro_carregamento:
End Sub
Private Sub V_Valorvenda_Change()
On Error GoTo erro_carregamento:
        
        'Após o preenchimendo do campo ele retorna branco
        V_ValorVenda.BackColor = &HFFFFFF
        
        'Bloquear virgula a esqueda
        If Left(V_ValorVenda, 1) = "," Then
            V_ValorVenda = Empty
        End If
        
Exit Sub
erro_carregamento:
End Sub
'Ao cliclar no campo ele desfaz a formatação em moeda
Private Sub V_Valorvenda_Enter()
On Error GoTo erro_carregamento:
        
        V_ValorVenda = Format(V_ValorVenda)
        
Exit Sub
erro_carregamento:
End Sub
'Ao sair do campo, formatar valor em moeda
Private Sub V_Valorvenda_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo erro_carregamento:
        
        V_ValorVenda = Format(V_ValorVenda, "Currency")
        
Exit Sub
erro_carregamento:
End Sub
'Bloquear para aceitar apenas numero
Private Sub V_valorvenda_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:
        
        'Tabela Keyascii para numeros, virgula e backspace(botão limpar)
        Select Case KeyAscii
            Case 8, 44, 48 To 57
            
                'Permitir apenas uma virgula no campo
                If KeyAscii = 44 Then
                    If InStr(1, V_ValorVenda, ",", vbTextCompare) > 1 Then
                    KeyAscii = 0
                    End If
                End If
            Case Else
            KeyAscii = 0
        End Select
        
Exit Sub
erro_carregamento:
End Sub
