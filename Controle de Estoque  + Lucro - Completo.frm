VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu_Principal 
   ClientHeight    =   9190.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17985
   OleObjectBlob   =   "Controle de Estoque  + Lucro - Completo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Menu_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub B_Exc_Vendas_Click()
On Error GoTo erro_carregamento:

    'Ativar Botões
        B_Home.BackStyle = fmBackStyleTransparent
        B_Produtos.BackStyle = fmBackStyleTransparent
        Me.B_Vendas.BackStyle = fmBackStyleTransparent
        Me.B_Exc_Vendas.BackStyle = fmBackStyleOpaque
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
        Me.B_Vendas.BackStyle = fmBackStyleTransparent
        Me.B_Exc_Vendas.BackStyle = fmBackStyleTransparent
        B_Lucro.BackStyle = fmBackStyleTransparent
        
        PAGINAS.Value = 0
        
Exit Sub
erro_carregamento:
End Sub

Private Sub B_Excluir_Click()
On Error GoTo erro_carregamento:

    If MsgBox("Deseja realmente continuar?", vbQuestion + vbYesNo, "Exclusão") = vbYes Then
        
        If P_Lista.Column(0) <> Empty Then
            
            On Error Resume Next
            With Sheets("PRODUTOS").Range("A:A")
                       
            Set EncontrarID = .Find(What:=P_Lista.Column(0), _
                                LookAt:=xlWhole)
                            
                If Not EncontrarID Is Nothing Then
                    Application.GoTo EncontrarID, True
                    
                    ActiveCell.Rows.EntireRow.Delete
                    
                    Call PreenchimentoProdutos
                    
                    ProgressBar.Show
                    MsgBox "Exclusão realizada com sucesso!", vbInformation, "Exclusao"
                                
                End If
            End With
            Err
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
        Me.B_Vendas.BackStyle = fmBackStyleTransparent
        Me.B_Exc_Vendas.BackStyle = fmBackStyleTransparent
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
        Me.B_Vendas.BackStyle = fmBackStyleTransparent
        Me.B_Exc_Vendas.BackStyle = fmBackStyleTransparent
        B_Lucro.BackStyle = fmBackStyleTransparent
        
        PAGINAS.Value = 1

Exit Sub
erro_carregamento:
End Sub

Private Sub B_Vendas_Click()
On Error GoTo erro_carregamento:

    'Ativar Botões
        B_Home.BackStyle = fmBackStyleTransparent
        B_Produtos.BackStyle = fmBackStyleTransparent
        Me.B_Vendas.BackStyle = fmBackStyleOpaque
        Me.B_Exc_Vendas.BackStyle = fmBackStyleTransparent
        B_Lucro.BackStyle = fmBackStyleTransparent

        PAGINAS.Value = 2
        
Exit Sub
erro_carregamento:
End Sub

Private Sub Barra_Menu_Click()
On Error GoTo erro_carregamento:

    Call Tela_Inicial
        
Exit Sub
erro_carregamento:
End Sub

Private Sub BEstImagem_Click()
On Error GoTo erro_carregamento:

    Call FotoProduto
    
Exit Sub
erro_carregamento:
End Sub


Private Sub E_Busca_Change()
On Error GoTo erro_carregamento:

    Call PreenchimentoListaVendasExcluir
    
Exit Sub
erro_carregamento:
End Sub

Private Sub E_Busca_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
Exit Sub
erro_carregamento:
End Sub

Private Sub E_DataFinal_AfterUpdate()
On Error GoTo erro_carregamento:

    If E_DataFinal <> Empty Then
    
        'Dia
        If Left(E_DataFinal, 2) > 31 Then
            E_DataFinal = Empty
            E_DataFinal.SetFocus
        Exit Sub
        
        'Mês
        ElseIf Mid(E_DataFinal, 4, 2) > 12 Then
            E_DataFinal = Empty
            E_DataFinal.SetFocus
        Else
        
            If Not IsDate(E_DataFinal) Then
                E_DataFinal = Empty
                E_DataFinal.SetFocus
            End If
            
            Call PreenchimentoListaVendasExcluir
        End If
    End If
        
Exit Sub
erro_carregamento:
End Sub


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
    
        'Dia
        If Left(E_DataInicial, 2) > 31 Then
            E_DataInicial = Empty
            E_DataInicial.SetFocus
        Exit Sub
        
        'Mês
        ElseIf Mid(E_DataInicial, 4, 2) > 12 Then
            E_DataInicial = Empty
            E_DataInicial.SetFocus
        Else
        
            If Not IsDate(E_DataInicial) Then
                E_DataInicial = Empty
                E_DataInicial.SetFocus
            End If
            
            Call PreenchimentoListaVendasExcluir
        End If
    End If
End Sub

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
    
    ElseIf MsgBox("Deseja relamente continuar?", vbQuestion + vbYesNo, "Exclusão") = vbYes Then
    
        'Estorno ao estoque
        Sheets("PRODUTOS").Select
        
        With Sheets("PRODUTOS").Range("A:A")
        
        Set EncontrarID = .Find(What:=E_Lista.Column(2), _
                            LookAt:=xlWhole)
                        
            If Not EncontrarID Is Nothing Then
                Application.GoTo EncontrarID, True
                
                ActiveCell.Offset(0, 5).Value = CDbl(ActiveCell.Offset(0, 5).Value) + CDbl(E_QtdeExcluir)
    
            End If
        End With
    
        'Exclusão
        If E_QtdeTotal.Value = True Then
            
            'Exclusão Total
            Sheets("VENDAS FINALIZADAS").Select
        
            With Sheets("VENDAS FINALIZADAS").Range("A:A")
            
            Set EncontrarID = .Find(What:=E_Lista.Column(0), _
                                LookAt:=xlWhole)
                            
                If Not EncontrarID Is Nothing Then
                    Application.GoTo EncontrarID, True
                    
                    ActiveCell.Rows.EntireRow.Delete
        
                End If
            End With
        
        Else
        
            'Exclusão Parcial ou Total
            Sheets("VENDAS FINALIZADAS").Select
        
            With Sheets("VENDAS FINALIZADAS").Range("A:A")
            
            Set EncontrarID = .Find(What:=E_Lista.Column(0), _
                                LookAt:=xlWhole)
                            
                If Not EncontrarID Is Nothing Then
                    Application.GoTo EncontrarID, True
                    
                    'Total
                    If E_QtdeFinal = Empty Or E_QtdeFinal = 0 Then
                        ActiveCell.Rows.EntireRow.Delete
                    'Parcial
                    Else
                        ActiveCell.Offset(0, 6).Value = E_QtdeFinal
                    End If
        
                End If
            End With
        End If
   
    End If
    
    ProgressBar.Show
    
    E_QtdeExcluir = Empty
    E_QtdeFinal = Empty
    
    Call PreenchimentoListaVendasExcluir
    Call PreenchimentoVendas
    Call PreenchimentoProdutos


Exit Sub
erro_carregamento:
End Sub

Private Sub E_Lista_Click()
On Error GoTo erro_carregamento:

    Sheets("PRODUTOS").Select
    
    With Sheets("PRODUTOS").Range("A:A")
    
    Set EncontrarID = .Find(What:=E_Lista.Column(2), _
                        LookAt:=xlWhole)
                    
        If Not EncontrarID Is Nothing Then
            Application.GoTo EncontrarID, True
            
            E_QtdeTotal.Value = True
            E_QtdeFinal = Empty
            
            E_QtdeExcluir = E_Lista.Column(6)
            
            If E_Lista.Column(6) > 1 Then
                E_QtdeParcial.Enabled = True
            Else
                E_QtdeParcial.Enabled = False
            End If
            
            
            If ActiveCell.Offset(0, 12) <> "" Then
                E_Imagem.Picture = LoadPicture(ActiveCell.Offset(0, 12))
                E_Imagem.PictureSizeMode = fmPictureSizeModeZoom
                E_Imagem.Visible = True
            Else
                E_Imagem.Visible = False
            End If
        End If
    End With
    
Exit Sub
erro_carregamento:
End Sub


Private Sub E_QtdeExcluir_Change()
On Error GoTo erro_carregamento:
    
    If E_QtdeExcluir = Empty Then
        E_QtdeFinal = Empty
    Else
        If CDbl(E_QtdeExcluir) > CDbl(E_Lista.Column(6)) Or E_QtdeExcluir = 0 Then
            E_QtdeExcluir = Empty
            E_QtdeFinal = Empty
        Else
            E_QtdeFinal = CDbl(E_Lista.Column(6)) - CDbl(E_QtdeExcluir)
        
        End If
    End If

Exit Sub
erro_carregamento:
End Sub

Private Sub E_QtdeParcial_Click()
On Error GoTo erro_carregamento:

    If E_QtdeParcial.Value = True Then
        E_QtdeExcluir.Enabled = True
        E_QtdeExcluir.SetFocus
    
    End If
    
Exit Sub
erro_carregamento:
End Sub

Private Sub E_QtdeTotal_Click()
On Error GoTo erro_carregamento:

    If E_QtdeTotal.Value = True Then
        E_QtdeExcluir = E_Lista.Column(6)
        E_QtdeFinal = Empty
        E_QtdeExcluir.Enabled = False
        
    End If
    
Exit Sub
erro_carregamento:
End Sub

Private Sub E_SCodigo_Click()
On Error GoTo erro_carregamento:

    E_Busca = Empty
    E_Busca.SetFocus
    
Exit Sub
erro_carregamento:
End Sub

Private Sub E_SProduto_Click()
On Error GoTo erro_carregamento:

    E_Busca = Empty
    E_Busca.SetFocus
    
Exit Sub
erro_carregamento:
End Sub

Private Sub L_Busca_Change()
On Error GoTo erro_carregamento:

    Call PreenchimentoListaVendasLucro
    
Exit Sub
erro_carregamento:
End Sub

Private Sub L_Busca_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
Exit Sub
erro_carregamento:
End Sub

Private Sub L_DataFinal_AfterUpdate()
On Error GoTo erro_carregamento:

    If L_DataFinal <> Empty Then
    
        'Dia
        If Left(L_DataFinal, 2) > 31 Then
            L_DataFinal = Empty
            L_DataFinal.SetFocus
        Exit Sub
        
        'Mês
        ElseIf Mid(L_DataFinal, 4, 2) > 12 Then
            L_DataFinal = Empty
            L_DataFinal.SetFocus
        Else
        
            If Not IsDate(L_DataFinal) Then
                L_DataFinal = Empty
                L_DataFinal.SetFocus
            End If
            
            Call PreenchimentoListaVendasLucro
        End If
    End If
        
Exit Sub
erro_carregamento:
End Sub


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
    
        'Dia
        If Left(L_DataInicial, 2) > 31 Then
            L_DataInicial = Empty
            L_DataInicial.SetFocus
        Exit Sub
        
        'Mês
        ElseIf Mid(L_DataInicial, 4, 2) > 12 Then
            L_DataInicial = Empty
            L_DataInicial.SetFocus
        Else
        
            If Not IsDate(L_DataInicial) Then
                L_DataInicial = Empty
                L_DataInicial.SetFocus
            End If
            
            Call PreenchimentoListaVendasLucro
        End If
    End If
End Sub

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

    If MsgBox("Deseja realmente continuar com o fechamento do sistema?", vbQuestion + vbYesNo, "Fechar") = vbYes Then
        
        ProgressBar.Show
        Application.Quit
        ActiveWorkbook.Close savechanges:=False
        Application.DisplayAlerts = False
        
    End If
    
Exit Sub
erro_carregamento:
End Sub

Private Sub L_SCodigo_Click()
On Error GoTo erro_carregamento:

    L_Busca = Empty
    L_Busca.SetFocus
    
Exit Sub
erro_carregamento:
End Sub

Private Sub L_SProduto_Click()
On Error GoTo erro_carregamento:

    L_Busca = Empty
    L_Busca.SetFocus
    
Exit Sub
erro_carregamento:
End Sub

Private Sub P_Busca_Change()
On Error GoTo erro_carregamento:

    Call PreenchimentoProdutos
    
Exit Sub
erro_carregamento:
End Sub

Private Sub P_Busca_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
Exit Sub
erro_carregamento:
End Sub



Private Sub P_Cadastrar_Click()
On Error GoTo erro_carregamento:

    If P_Produto = Empty Then
    
        MsgBox "Informe o nome do Produto!", vbExclamation, "Cadastro"
        P_Produto.BackColor = &HFFFF&
        P_Produto.SetFocus
        
    Exit Sub
    
    ElseIf P_Qtde = Empty Then
    
        MsgBox "Informe a quantidade!", vbExclamation, "Cadastro"
        P_Qtde.BackColor = &HFFFF&
        P_Qtde.SetFocus
    
    Exit Sub
    
    ElseIf P_ValorCompra = Empty Then
    
        MsgBox "Informe o valor de compra!", vbExclamation, "Cadastro"
        P_ValorCompra.BackColor = &HFFFF&
        P_ValorCompra.SetFocus
    
    Exit Sub
    
    ElseIf P_ValorVenda = Empty Then
    
        MsgBox "Informe o valor de venda!", vbExclamation, "Cadastro"
        P_ValorVenda.BackColor = &HFFFF&
        P_ValorVenda.SetFocus
        
    Exit Sub
    
    ElseIf MsgBox("Deseja realmente continuar?", vbQuestion + vbYesNo, "Cadastro") = vbYes Then

        On Error Resume Next
        Sheets("PRODUTOS").Select
        Application.GoTo Reference:="R1048576C1"
        Selection.End(xlUp).Select
        ActiveCell.Offset(1, 0).Select
        
        ContCodigo = Range("AR1").Value + 1
        
        'Lançamento de dados
        ActiveCell.Value = ContCodigo
        ActiveCell.Offset(0, 2).Value = Format(P_CodBarras)
        ActiveCell.Offset(0, 3).Value = Format(P_Produto)
        ActiveCell.Offset(0, 4).Value = Format(P_Marca)
        ActiveCell.Offset(0, 5).Value = Format(P_Qtde)
        ActiveCell.Offset(0, 6).Value = Format(P_UnidadeMedida)
        ActiveCell.Offset(0, 7).Value = Format(P_ValorCompra, "Currency")
        ActiveCell.Offset(0, 8).Value = Format(P_ValorVenda, "Currency")
        ActiveCell.Offset(0, 9).Value = CDate(P_DataCompra)
        ActiveCell.Offset(0, 10).Value = CDate(P_DataValidade)
        ActiveCell.Offset(0, 11).Value = Format(P_Outras)
        ActiveCell.Offset(0, 12).Value = SalvarFoto
        
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
        
        Call ClassificarProdutos
        Call PreenchimentoProdutos
        
        ProgressBar.Show
        MsgBox "Cadastro realizado com sucesso!", vbInformation, "Cadastro"
        Err
    End If
    
    
Exit Sub
erro_carregamento:
End Sub

Private Sub P_CodBarras_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:

    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else
        KeyAscii = 0
    End Select
        
Exit Sub
erro_carregamento:
End Sub

Private Sub P_DataCompra_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo erro_carregamento:

    If P_DataCompra <> Empty Then
    
        'Dia
        If Left(P_DataCompra, 2) > 31 Then
            P_DataCompra = Empty
            P_DataCompra.SetFocus
        Exit Sub
        
        'Mês
        ElseIf Mid(P_DataCompra, 4, 2) > 12 Then
            P_DataCompra = Empty
            P_DataCompra.SetFocus
        Else
        
            If Not IsDate(P_DataCompra) Then
                P_DataCompra = Empty
                P_DataCompra.SetFocus
            End If
        End If
    End If
        
Exit Sub
erro_carregamento:
End Sub

Private Sub P_DataVenda_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo erro_carregamento:

    If P_DataVenda <> Empty Then
    
        'Dia
        If Left(P_DataVenda, 2) > 31 Then
            P_DataVenda = Empty
            P_DataVenda.SetFocus
        Exit Sub
        
        'Mês
        ElseIf Mid(P_DataVenda, 4, 2) > 12 Then
            P_DataVenda = Empty
            P_DataVenda.SetFocus
        Else
        
            If Not IsDate(P_DataVenda) Then
                P_DataVenda = Empty
                P_DataVenda.SetFocus
            End If
        End If
    End If
        
Exit Sub
erro_carregamento:
End Sub

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

Private Sub P_DataValidade_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo erro_carregamento:

    If P_DataValidade <> Empty Then
    
        'Dia
        If Left(P_DataValidade, 2) > 31 Then
            P_DataValidade = Empty
            P_DataValidade.SetFocus
        Exit Sub
        
        'Mês
        ElseIf Mid(P_DataValidade, 4, 2) > 12 Then
            P_DataValidade = Empty
            P_DataValidade.SetFocus
        Else
        
            If Not IsDate(P_DataValidade) Then
                P_DataValidade = Empty
                P_DataValidade.SetFocus
            End If
        End If
    End If
        
Exit Sub
erro_carregamento:
End Sub

Private Sub P_Lista_Click()
On Error GoTo erro_carregamento:

    Sheets("PRODUTOS").Select
    
    With Sheets("PRODUTOS").Range("A:A")
    
    Set EncontrarID = .Find(What:=P_Lista.Column(0), _
                        LookAt:=xlWhole)
                    
        If Not EncontrarID Is Nothing Then
            Application.GoTo EncontrarID, True
            
            If ActiveCell.Offset(0, 12) <> "" Then
                P_Imagem.Picture = LoadPicture(ActiveCell.Offset(0, 12))
                P_Imagem.PictureSizeMode = fmPictureSizeModeZoom
                P_Imagem.Visible = True
            Else
                P_Imagem.Visible = False
            End If
        End If
    End With
    
Exit Sub
erro_carregamento:
End Sub

Private Sub P_Lista_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo erro_carregamento:

    On Error Resume Next
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
    
    If P_Lista.Column(12) <> Empty Then
    
        Alterar_Produtos.P_Imagem.Picture = LoadPicture(P_Lista.Column(12))
        Alterar_Produtos.P_Imagem.PictureSizeMode = fmPictureSizeModeZoom
        Alterar_Produtos.P_Imagem.Visible = True

    End If
    Err
    
    Alterar_Produtos.Show
    
Exit Sub
erro_carregamento:
End Sub

Private Sub P_Marca_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
Exit Sub
erro_carregamento:
End Sub

Private Sub P_Outros_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
Exit Sub
erro_carregamento:
End Sub


Private Sub P_Outras_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
Exit Sub
erro_carregamento:
End Sub

Private Sub P_Produto_Change()
On Error GoTo erro_carregamento:

    P_Produto.BackColor = &HFFFFFF
    
Exit Sub
erro_carregamento:
End Sub

Private Sub P_Produto_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
Exit Sub
erro_carregamento:
End Sub

Private Sub P_Qtde_Change()
On Error GoTo erro_carregamento:

    P_Qtde.BackColor = &HFFFFFF
    
Exit Sub
erro_carregamento:
End Sub

Private Sub P_Qtde_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:

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

    P_Busca = Empty
    P_Busca.SetFocus
    
Exit Sub
erro_carregamento:
End Sub

Private Sub P_SProduto_Click()
On Error GoTo erro_carregamento:

    P_Busca = Empty
    P_Busca.SetFocus
    
Exit Sub
erro_carregamento:
End Sub

Private Sub P_ValorCompra_Change()
On Error GoTo erro_carregamento:

    P_ValorCompra.BackColor = &H80000005
    
    If Left(P_ValorCompra, 1) = "," Then
        P_ValorCompra = Empty
    End If
        
Exit Sub
erro_carregamento:
End Sub

Private Sub P_ValorCompra_Enter()
On Error GoTo erro_carregamento:
    
    P_ValorCompra = Format(P_ValorCompra)
        
Exit Sub
erro_carregamento:
End Sub

Private Sub P_ValorVenda_Enter()
On Error GoTo erro_carregamento:
    
    P_ValorVenda = Format(P_ValorVenda)
        
Exit Sub
erro_carregamento:
End Sub

Private Sub P_ValorCompra_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo erro_carregamento:
    
    P_ValorCompra = Format(P_ValorCompra, "Currency")
        
Exit Sub
erro_carregamento:
End Sub

Private Sub P_ValorCompra_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:

    Select Case KeyAscii
        Case 8, 44, 48 To 57
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

Private Sub P_ValorVenda_Change()
On Error GoTo erro_carregamento:
    
    P_ValorVenda.BackColor = &H80000005

    If Left(P_ValorVenda, 1) = "," Then
        P_ValorVenda = Empty
    End If
        
Exit Sub
erro_carregamento:
End Sub

Private Sub P_ValorVenda_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:

    Select Case KeyAscii
        Case 8, 44, 48 To 57
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

Private Sub P_ValorVenda_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo erro_carregamento:
    
    P_ValorVenda = Format(P_ValorVenda, "Currency")
        
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

Private Sub UserForm_Initialize()
On Error GoTo erro_carregamento:
    
    Call PreenchimentoProdutos
    Call PreenchimentoVendasProdutos
    Call PreenchimentoVendas
    
    L_DataInicial = Format(Now - 1825, "dd/mm/yyyy")
    L_DataFinal = Format(Now + 365, "dd/mm/yyyy")

    Call PreenchimentoListaVendasLucro
    
    E_DataInicial = Format(Now - 1825, "dd/mm/yyyy")
    E_DataFinal = Format(Now + 365, "dd/mm/yyyy")

    Call PreenchimentoListaVendasExcluir
   

    PAGINAS.Style = fmTabStyleNone
    
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
    
    ElseIf CDbl(V_Qtde) > CDbl(V_Lista.Column(5)) Then
    
        MsgBox "Quantidade solicitada é maior que o estoque atual!", vbExclamation, "Venda"
                
    Else
        On Error Resume Next
        Sheets("LISTA DE VENDAS").Select
        
        Application.GoTo Reference:="R1048576C1"
        Selection.End(xlUp).Select
        ActiveCell.Offset(1, 0).Select
        
        ContCodigo = Range("AR1").Value + 1
        
        ActiveCell.Value = ContCodigo
        ActiveCell.Offset(0, 1).Value = V_Lista.Column(0)
        ActiveCell.Offset(0, 2).Value = Format(V_Lista.Column(2)) 'Cod. Barras
        ActiveCell.Offset(0, 3).Value = Format(V_Lista.Column(3)) 'Produtos
        ActiveCell.Offset(0, 4).Value = Format(V_Lista.Column(4)) 'Marca
        ActiveCell.Offset(0, 5).Value = Format(V_Qtde)            'Qtde
        ActiveCell.Offset(0, 6).Value = Format(V_Lista.Column(6)) 'Unid. Medida
        ActiveCell.Offset(0, 7).Value = Format(V_ValorVenda, "Currency")     'Valor Venda
        ActiveCell.Offset(0, 8).Value = Format(CDbl(V_Qtde) * CDbl(V_ValorVenda), "Currency") 'Valor Total de Venda
        
        'Lucro
        ActiveCell.Offset(0, 9).Value = Format(CDbl(ActiveCell.Offset(0, 8).Value) - (CDbl(V_Qtde) * CDbl(V_Lista.Column(7))), "Currency")
        
        Range("AR1").Value = ContCodigo
        
        Sheets("PRODUTOS").Select
        
        With Sheets("PRODUTOS").Range("A:A")
        
        Set EncontrarID = .Find(What:=V_Lista.Column(0), _
                            LookAt:=xlWhole)
                        
            If Not EncontrarID Is Nothing Then
                Application.GoTo EncontrarID, True
                
                ActiveCell.Offset(0, 5).Value = CDbl(ActiveCell.Offset(0, 5).Value) - CDbl(V_Qtde)
            End If
        End With
        Err
        
        Call PreenchimentoVendas
        Call PreenchimentoProdutos
        Call PreenchimentoVendasProdutos
        
    
    End If

    
Exit Sub
erro_carregamento:
End Sub

Private Sub V_Busca_Change()
On Error GoTo erro_carregamento:

    Call PreenchimentoVendas
    
Exit Sub
erro_carregamento:
End Sub

Private Sub V_Busca_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:

   KeyAscii = Asc(UCase(Chr(KeyAscii)))
        
Exit Sub
erro_carregamento:
End Sub

Private Sub V_Excluir_Click()
On Error GoTo erro_carregamento:
    
    On Error Resume Next
    'Estorno ao estoque
    Sheets("PRODUTOS").Select
    
    With Sheets("PRODUTOS").Range("A:A")
    
    Set EncontrarID = .Find(What:=V_ListaProdutos.Column(1), _
                        LookAt:=xlWhole)
                    
        If Not EncontrarID Is Nothing Then
            Application.GoTo EncontrarID, True
            
            ActiveCell.Offset(0, 5) = CDbl(ActiveCell.Offset(0, 5)) + CDbl(V_ListaProdutos.Column(5))
        End If
    End With

    'Excluir Produto
    Sheets("LISTA DE VENDAS").Select
    
    With Sheets("LISTA DE VENDAS").Range("A:A")
    
    Set EncontrarID = .Find(What:=V_ListaProdutos.Column(0), _
                        LookAt:=xlWhole)
                    
        If Not EncontrarID Is Nothing Then
            Application.GoTo EncontrarID, True
            
            ActiveCell.Rows.EntireRow.Delete
        End If
    End With
    Err
    
    V_ListaProdutos.Clear
    Call PreenchimentoVendas
    Call PreenchimentoProdutos
    Call PreenchimentoVendasProdutos
    


Exit Sub
erro_carregamento:
End Sub

Private Sub V_ExcluirTudo_Click()
On Error GoTo erro_carregamento:
    
    On Error Resume Next
    Sheets("LISTA DE VENDAS").Select
    Range("A2").Select
    
    Do Until ActiveCell.Value = Empty
    
    ChaveProduto = ActiveCell.Offset(0, 1).Value
    QtdeProduto = ActiveCell.Offset(0, 5).Value
    
        'Estorno ao estoque
        Sheets("PRODUTOS").Select
        
        With Sheets("PRODUTOS").Range("A:A")
        
        Set EncontrarID = .Find(What:=ChaveProduto, _
                            LookAt:=xlWhole)
                        
            If Not EncontrarID Is Nothing Then
                Application.GoTo EncontrarID, True
                
                ActiveCell.Offset(0, 5) = CDbl(ActiveCell.Offset(0, 5)) + CDbl(QtdeProduto)
            End If
        End With
    
        'Excluir Produto
        Sheets("LISTA DE VENDAS").Select
       
        ActiveCell.Rows.EntireRow.Delete
    
    
    Loop
    
    Sheets("LISTA DE VENDAS").Range("AR1").Value = 0
    Err
    
    V_ListaProdutos.Clear
    Call PreenchimentoVendas
    Call PreenchimentoProdutos
    Call PreenchimentoVendasProdutos

Exit Sub
erro_carregamento:
End Sub

Private Sub V_Finalizar_Click()
On Error GoTo erro_carregamento:

ContProxLinha = 2

    If V_QtdeItens = 0 Or V_QtdeItens = Empty Then
    
        MsgBox "Adicione um produto para venda!", vbExclamation, "Venda"
        
    Exit Sub
    
    ElseIf Me.V_FormaPagto = Empty Then
    
        V_FormaPagto.DropDown
        MsgBox "Selecione uma forma de pagamento!", vbExclamation, "Venda"
    
    Exit Sub
    
    ElseIf Me.V_TipoPagto = Empty Then
    
        V_TipoPagto.DropDown
        MsgBox "Selecione um tipo de pagamento!", vbExclamation, "Venda"
    
    Exit Sub
    
    ElseIf MsgBox("Deseja realmente continuar?", vbQuestion + vbYesNo, "Venda") = vbYes Then
    
        On Error Resume Next
        Sheets("LISTA DE VENDAS").Select
        Range("B2").Select
        
        SalvarNumeroVenda = (Format(Now, "yyyy") & Format(1, "0000")) + Range("AS1").Value
        
        Do Until ActiveCell.Value = Empty
        
            Range("B" & ContProxLinha & ": J" & ContProxLinha).Select
            
            Selection.Copy
            
            Sheets("VENDAS FINALIZADAS").Select
            
            Application.GoTo Reference:="R1048576C1"
            Selection.End(xlUp).Select
            ActiveCell.Offset(1, 0).Select
            
            ContCodigo = Range("AR1").Value + 1
            
            ActiveCell.Value = ContCodigo
            ActiveCell.Offset(0, 2).Select
            
            ActiveSheet.Paste
            
            ActiveCell.Offset(0, 9).Value = SalvarNumeroVenda
            ActiveCell.Offset(0, 10).Value = Format(Now, "mm/dd/yyyy") 'Data da venda
            ActiveCell.Offset(0, 11).Value = Me.V_FormaPagto
            ActiveCell.Offset(0, 12).Value = Me.V_TipoPagto
            
            Range("AR1").Value = ContCodigo
            
            Sheets("LISTA DE VENDAS").Select
            ActiveCell.Offset(1, 0).Select
                   
            ContProxLinha = ContProxLinha + 1
                    
        Loop
        
        Sheets("LISTA DE VENDAS").Select
        Ultimalinha = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
        Range("A2:J" & Ultimalinha).Rows.EntireRow.Delete

        Range("AR1").Value = 0
        Range("AS1").Value = Range("AS1").Value + 1
        
        V_ListaProdutos.Clear
        Call PreenchimentoVendasProdutos
        
        ProgressBar.Show
        MsgBox "Venda nº " & SalvarNumeroVenda & " finalizada com sucesso!", vbInformation, "Venda"
        Err
        
        
    End If

Exit Sub
erro_carregamento:
End Sub

Private Sub V_FormaPagto_Change()
On Error GoTo erro_carregamento:

    If V_FormaPagto = "À VISTA" Then
        With Me.V_TipoPagto
            .Clear
            .AddItem "DINHEIRO"
            .AddItem "DOC"
            .AddItem "TED"
        End With
    End If
        
    If V_FormaPagto = "À PRAZO" Then
        With Me.V_TipoPagto
            .Clear
            .AddItem "FIADO"
            .AddItem "BOLETO"
            .AddItem "CARNÊ"
        End With
    End If
        
    If V_FormaPagto = "CARTÃO" Then
        With Me.V_TipoPagto
            .Clear
            .AddItem "DÉBITO - VISA"
            .AddItem "DÉBITO - MASTER"
            .AddItem "DÉBITO ELO"
            .AddItem "CRÉDITO - VISA"
            .AddItem "CRÉDITO - MASTER"
            .AddItem "CRÉDITO - ELO"
        End With
    End If
    
Exit Sub
erro_carregamento:
End Sub

Private Sub V_Lista_Click()
On Error GoTo erro_carregamento:

    Sheets("PRODUTOS").Select
    
    With Sheets("PRODUTOS").Range("A:A")
    
    Set EncontrarID = .Find(What:=V_Lista.Column(0), _
                        LookAt:=xlWhole)
                    
        If Not EncontrarID Is Nothing Then
            Application.GoTo EncontrarID, True
            
            If V_Lista.Column(0) = Empty Then
                V_Produto = Empty
                V_Qtde = Empty
                V_UnidadeMedida = Empty
                V_ValorVenda = Empty
                V_Imagem.Visible = False
            Else
            
                V_Produto = V_Lista.Column(3)
                V_Qtde = 1
                V_UnidadeMedida = V_Lista.Column(6)
                V_ValorVenda = V_Lista.Column(8)
                
                If ActiveCell.Offset(0, 12) <> "" Then
                    V_Imagem.Picture = LoadPicture(ActiveCell.Offset(0, 12))
                    V_Imagem.PictureSizeMode = fmPictureSizeModeZoom
                    V_Imagem.Visible = True
                Else
                    V_Imagem.Visible = False
                End If
            
            End If
            
        Else
            V_Produto = Empty
            V_Qtde = Empty
            V_UnidadeMedida = Empty
            V_ValorVenda = Empty
            V_Imagem.Visible = False
        End If
    End With
    
    
Exit Sub
erro_carregamento:
End Sub

Private Sub V_Qtde_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:

    Select Case KeyAscii
        Case 8, 48 To 57
        Case Else
        KeyAscii = 0
    End Select
        
Exit Sub
erro_carregamento:
End Sub

Private Sub V_SCodigo_Click()
On Error GoTo erro_carregamento:

    V_Busca = Empty
    V_Busca.SetFocus
    
Exit Sub
erro_carregamento:
End Sub

Private Sub V_SProduto_Click()
On Error GoTo erro_carregamento:

    V_Busca = Empty
    V_Busca.SetFocus
    
Exit Sub
erro_carregamento:
End Sub

Private Sub V_ValorVenda_Enter()
On Error GoTo erro_carregamento:
    
    V_ValorVenda = Format(V_ValorVenda)
        
Exit Sub
erro_carregamento:
End Sub

Private Sub V_ValorVenda_Change()
On Error GoTo erro_carregamento:
    
    V_ValorVenda.BackColor = &H80000005

    If Left(V_ValorVenda, 1) = "," Then
        V_ValorVenda = Empty
    End If
        
Exit Sub
erro_carregamento:
End Sub

Private Sub V_ValorVenda_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
On Error GoTo erro_carregamento:

    Select Case KeyAscii
        Case 8, 44, 48 To 57
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

Private Sub V_ValorVenda_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo erro_carregamento:
    
    V_ValorVenda = Format(V_ValorVenda, "Currency")
        
Exit Sub
erro_carregamento:
End Sub

