Imports System.CodeDom.Compiler
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices
Imports System.Security.Policy
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox
Imports Guna.UI2.Native.WinApi
Imports Guna.UI2.WinForms
Imports Microsoft.VisualBasic.ApplicationServices

Public Class Form1
    'Click no botão de cadastrar cliente
    Private Sub btn_cadastrar_cliente_Click(sender As Object, e As EventArgs) Handles btn_cadastrar_cliente.Click
        'Armazenando os valores digitados
        Dim nomeCliente As String = txt_nome_cliente.Text
        Dim cpf As String = txt_cpf_cliente.Text
        Dim email As String = txt_email_cliente.Text
        Dim celular As String = txt_celular_cliente.Text
        Dim endereco As String = txt_endereco.Text
        Dim numero As String = txt_numero.Text
        Dim bairro As String = txt_bairro.Text
        Dim cidade As String = txt_cidade.Text
        Dim uf As String = txt_uf.Text

        Try
            'Verificando se há campos em branco
            If String.IsNullOrWhiteSpace(nomeCliente) OrElse
                String.IsNullOrWhiteSpace(cpf) OrElse
                String.IsNullOrWhiteSpace(email) OrElse
                String.IsNullOrWhiteSpace(celular) OrElse
                String.IsNullOrWhiteSpace(endereco) OrElse
                String.IsNullOrWhiteSpace(numero) OrElse
                String.IsNullOrWhiteSpace(bairro) OrElse
                String.IsNullOrWhiteSpace(cidade) OrElse
                String.IsNullOrWhiteSpace(uf) Then
                MsgBox("Por favor, preencha todos os campos.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "ERRO!")
                Return
            End If

            'Realizando consulta SQL
            sql = "select * from tb_clientes where cpf = '" & cpf & "'"
            rs = db.Execute(sql)
            'Verificando se o CPF já foi cadastrado
            If Not rs.EOF Then
                'Confirma a alteração dos dados se o CPF já existir
                resp = MsgBox("Registro já existente, deseja atualizar os dados?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "ATENÇÃO!")
                If resp = vbYes Then
                    sql = "update tb_clientes set nome = '" & nomeCliente & "', email = '" & email & "', celular = '" & celular & "', endereco = '" & endereco & "', numero = '" & numero & "', bairro = '" & bairro & "', cidade = '" & cidade & "', uf = '" & uf & "' where cpf = '" & cpf & "'"
                    rs = db.Execute(UCase(sql))
                    MsgBox("Dados Atualizados com Sucesso!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCESSO!")
                    'Esvazia os campos
                    limpar_clientes()
                Else
                    Exit Sub
                End If
                'Se o CPF não existir, o registro é criado
            Else
                sql = "insert into tb_clientes (nome, cpf, email, celular, endereco, numero, bairro, cidade, uf) values ('" & nomeCliente & "', " &
                "'" & cpf & "', '" & email & "', '" & celular & "', '" & endereco & "', '" & numero & "', '" & bairro & "', '" & cidade & "', '" & uf & "')"
                rs = db.Execute(UCase(sql))
                MsgBox("Dados Gravados com Sucesso!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCESSO!")

                'Esvazia os campos
                limpar_clientes()
            End If
        Catch ex As Exception
            MsgBox("Erro ao Cadastrar o Cliente", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ERRO!")
        End Try
    End Sub

    'Carregamento do formulário
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Carregar dados do DataGridView
        carregar_pedidos()
        'Foca no campo de login
        txt_senha.Focus()
        'Conecta o banco de dados assim que o formulário é aberto
        conectar_banco()
        'Carrega os dados do DataGridView
        carregar_pedidos()
        'Torna o botão de encerrar sessão invisível
        btn_encerrar.Visible = False

        'Deixa as guias de funcionários ocultas
        If TabControl1.TabPages.Contains(TabPage2) Or
           TabControl1.TabPages.Contains(TabPage3) Or
           TabControl1.TabPages.Contains(TabPage4) Then
            TabControl1.TabPages.Remove(TabPage2)
            TabControl1.TabPages.Remove(TabPage3)
            TabControl1.TabPages.Remove(TabPage4)
            TabControl1.TabPages.Remove(TabPage5)
        End If
    End Sub

    'Click no botão de login
    Private Sub btn_login_Click(sender As Object, e As EventArgs) Handles btn_login.Click
        'Armazena os valores digitados
        Dim login As String = txt_senha.Text
        Dim senha As String = txt_login.Text

        'Verifica se o usuário já está logado e se deseja sair
        If TabControl1.TabPages.Contains(TabPage2) Then
            resp = MsgBox("Opa, você já está logado! Deseja encerrar sessão?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "ERRO")
            If resp = vbYes Then
                TabControl1.TabPages.Remove(TabPage2)
                TabControl1.TabPages.Remove(TabPage3)
                TabControl1.TabPages.Remove(TabPage4)
                TabControl1.TabPages.Remove(TabPage5)
                btn_encerrar.Visible = False
            End If
        Else
            'Verifica se há algum campo vazio
            If String.IsNullOrWhiteSpace(login) OrElse String.IsNullOrWhiteSpace(senha) Then
                MsgBox("Por favor, preencha todos os campos!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ERRO!")
                Exit Sub
            End If

            Try
                'Realiza a consulta SQL
                sql = "select * from tb_funcionarios where cpf = '" & login & "' and senha = '" & senha & "'"
                rs = db.Execute(sql)

                'Se os dados baterem, o login é efetuado
                If Not rs.EOF Then
                    MsgBox("Login bem-sucedido!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCESSO!")

                    'Limpa os campos
                    limpar_login()
                    'Exibe o botão de encerrar sessão
                    btn_encerrar.Visible = True

                    'Exibe as guias para funcionários
                    TabControl1.TabPages.Add(TabPage2)
                    TabControl1.TabPages.Add(TabPage3)
                    TabControl1.TabPages.Add(TabPage4)
                    TabControl1.TabPages.Add(TabPage5)
                Else
                    MsgBox("Credenciais inválidas!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ERRO")
                    Exit Sub
                End If
            Catch ex As Exception
                MsgBox("Erro ao realizar o login", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ERRO")
            End Try
        End If
    End Sub

    'Click no botão de cadastrar funcionários
    Private Sub btn_cadastrar_func_Click(sender As Object, e As EventArgs) Handles btn_cadastrar_func.Click
        'Armazena os valores digitados
        Dim nomeFuncionario As String = txt_nome_func.Text
        Dim cpf As String = txt_cpf_func.Text
        Dim email As String = txt_email_func.Text
        Dim celular As String = txt_celular_func.Text
        Dim senha As String = txt_criar_senha.Text
        Dim cargo As String = cmb_cargo.Text
        Dim chave As String = txt_chave.Text

        'Verifica se existem campos vazios
        If String.IsNullOrWhiteSpace(nomeFuncionario) OrElse
            String.IsNullOrWhiteSpace(cpf) OrElse
            String.IsNullOrWhiteSpace(email) OrElse
            String.IsNullOrWhiteSpace(celular) OrElse
            String.IsNullOrWhiteSpace(senha) OrElse
            String.IsNullOrWhiteSpace(cargo) Then
            MsgBox("Por favor, preencha todos os campos.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "ERRO!")
            Return
        End If

        If String.IsNullOrWhiteSpace(chave) Then
            resp = MsgBox("Tem certeza que não deseja criar uma chave? Ela serve para recuperar a sua senha caso a esqueça", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "ATENÇÃO")
            If resp = vbNo Then
                txt_chave.Focus()
                Return
            End If
        End If

        'Realiza a consulta SQL
        sql = "select * from tb_funcionarios where cpf = '" & cpf & "'"
        rs = db.Execute(sql)

        'Verifica se o registro já existe
        If Not rs.EOF Then
            'Confirma a atualização dos dados já existentes
            resp = MsgBox("Registro já existente, deseja atualizar os dados?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "ATENÇÃO!")
            If resp = vbYes Then
                sql = "update tb_funcionarios set nome = '" & nomeFuncionario & "', email = '" & email & "', celular = '" & celular & "', senha = '" & senha & "', cargo = '" & cargo & "', chave = '" & chave & "' where cpf = '" & cpf & "'"
                rs = db.Execute(UCase(sql))
                MsgBox("Dados Atualizados com Sucesso!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCESSO!")
                limpar_func()
            Else
                Exit Sub
            End If
        Else
            'Gera um novo registro
            sql = "insert into tb_funcionarios (nome, cpf, email, celular, senha, cargo, chave) values ('" & nomeFuncionario & "', " &
                "'" & cpf & "', '" & email & "', '" & celular & "', '" & senha & "', '" & cargo & "', '" & chave & "')"
            rs = db.Execute(UCase(sql))
            MsgBox("Dados Gravados com Sucesso!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCESSO!")

            'Esvazia os campos
            limpar_func()
        End If
    End Sub

    'Troca de valor na ComboBox de massas
    Private Sub cmb_massa_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmb_massa.SelectedValueChanged
        'Confirma que o valor da ComboBox não é nulo e o armazena
        If cmb_massa.SelectedItem IsNot Nothing Then
            massa = cmb_massa.SelectedItem.ToString()
        End If

        'Exibe preços diferentes dependendo da escolha do usuário
        Select Case massa
            Case "Bolo"
                precoBase = 20.0
            Case "Torta"
                precoBase = 25.0
            Case "Cupcake"
                precoBase = 5.0
        End Select

        'Atualiza os preços exibidos
        lbl_preco_massa.Text = precoBase
        lbl_preco_total.Text = precoBase + precoTamanho + precoSabor
    End Sub

    'Troca de valor na ComboBox de tamanhos
    Private Sub cmb_tamanho_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmb_tamanho.SelectedValueChanged
        'Confirma que o valor da ComboBox não é nulo e o armazena
        If cmb_tamanho.SelectedItem IsNot Nothing Then
            tamanho = cmb_tamanho.SelectedItem.ToString()
        End If

        'Exibe preços diferentes dependendo da escolha do usuário
        Select Case tamanho
            Case "P (Individual)"
                precoTamanho = precoBase * 0.8
            Case "M (2-4 pessoas)"
                precoTamanho = precoBase
            Case "G (5-8 pessoas)"
                precoTamanho = precoBase * 1.2
        End Select

        'Atualiza os preços exibidos
        lbl_preco_tam.Text = precoTamanho
        lbl_preco_total.Text = precoBase + precoTamanho + precoSabor
    End Sub

    'Troca de valor na ComboBox de sabores
    Private Sub cmb_sabor_SelectedValueChanged(sender As Object, e As EventArgs) Handles cmb_sabor.SelectedValueChanged
        'Confirma que o valor da ComboBox não é nulo e o armazena
        If cmb_sabor.SelectedItem IsNot Nothing Then
            sabor = cmb_sabor.SelectedItem.ToString()
        End If

        'Exibe preços diferentes dependendo da escolha do usuário
        Select Case sabor
            Case "Morango"
                precoSabor = 5.0
            Case "Chocolate"
                precoSabor = 4.0
            Case "Baunilha"
                precoSabor = 3.0
        End Select

        'Atualiza os preços exibidos
        lbl_preco_sabor.Text = precoSabor
        lbl_preco_total.Text = precoBase + precoTamanho + precoSabor
    End Sub

    'Click no botão de confirmar pedido
    Private Sub btn_confirmar_Click(sender As Object, e As EventArgs) Handles btn_confirmar.Click
        'Armazena os valores informados
        Dim cpf = txt_cpf_pedido.Text
        Dim massa = cmb_massa.Text
        Dim tamanho = cmb_tamanho.Text
        Dim sabor = cmb_sabor.Text
        Dim valor = lbl_preco_total.Text
        Dim status = "EM PREPARO"

        'Verifica se o CPF informado está cadastrado
        sql = "select count(*) from tb_clientes where cpf = '" & cpf & "'"
        rs = db.Execute(sql)
        If rs.Fields(0).Value = 0 Then
            resp = MsgBox("Esse CPF não está registrado no sistema, deseja adicioná-lo?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "ERRO!")
            If resp = vbYes Then
                'Redireciona para a tela de cadastro
                redireciona_cadastro(cpf)
            Else
                MsgBox("Por favor, insira um CPF cadastrado", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ATENÇÃO")
            End If
            Return
        End If

        'Insere os valores no banco de dados
        sql = "INSERT INTO tb_pedidos (cliente, tipo_massa, sabor, tamanho, valor, status) VALUES ('" & cpf & "', " &
                "'" & massa & "', '" & sabor & "', '" & tamanho & "', '" & valor & "', '" & status & "')"
        rs = db.Execute(UCase(sql))
        MsgBox("Pedido Registrado!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCESSO!")

        'Atualiza os pedidos no DataGridView
        carregar_pedidos()
        'Limpa os campos
        limpar_pedido()
    End Sub

    'Click no botão de limpar pedido
    Private Sub btn_limpar_Click(sender As Object, e As EventArgs) Handles btn_limpar.Click
        limpar_pedido()
    End Sub

    'Click no botão de encerrar sessão
    Private Sub btn_encerrar_Click(sender As Object, e As EventArgs) Handles btn_encerrar.Click
        resp = MsgBox("Tem certeza que deseja sair?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "ATENÇÃO")

        If resp = vbYes Then
            'Oculta as guias de funcionário
            TabControl1.TabPages.Remove(TabPage2)
            TabControl1.TabPages.Remove(TabPage3)
            TabControl1.TabPages.Remove(TabPage4)
            TabControl1.TabPages.Remove(TabPage5)
            'Oculta o botão de sair 
            btn_encerrar.Visible = False
        End If
    End Sub

    'Click no Label de esqueci a senha
    Private Sub llb_esqueceu_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles llb_esqueceu.LinkClicked
        'MessageBox para recuperação de senha
        Dim user = InputBox("Digite o CPF ou e-mail:", "RECUPERAR SENHA")
        'Consulta SQL para localizar o usuário
        sql = "select * from tb_funcionarios where cpf='" & user & "' or email='" & user & "'"
        rs = db.Execute(sql)
        If rs.EOF = False Then
            'Armazenando os valores dos campos
            Dim nome As String = rs.Fields(0).Value
            Dim senha As String = rs.Fields(4).Value
            Dim fone As String = rs.Fields(3).Value
            Dim chave = rs.Fields(6).Value
            'Verifica se a chave existe, visto que é opcional no cadastro
            If IsDBNull(chave) Or chave.ToString() = "" Then
                resp = MsgBox("Opa, parece que você não criou uma chave de recuperação. Por favor, fale com um administrador que tenha acesso ao banco de dados.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ERRO")
            Else
                Process.Start("https://api.callmebot.com/whatsapp.php?phone=" & fone & "&text='Prezado(a) " & nome & ", sua senha é: " & senha & "'&apikey=" & chave & "")
            End If
        Else
            MsgBox("Conta Inválida!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "ATENÇÃO")
        End If
    End Sub

    'Click nas células do DataGridView
    Private Sub dgv_pedidos_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_pedidos.CellContentClick
        'Verifica se o clique ocorreu na coluna "Excluir"
        If e.ColumnIndex = dgv_pedidos.Columns("Column9").Index AndAlso e.RowIndex >= 0 Then
            'Verifica se a linha está vazia
            Dim id As Object = dgv_pedidos.Rows(e.RowIndex).Cells("Column1").Value
            If id IsNot Nothing AndAlso Not String.IsNullOrEmpty(id.ToString()) Then
                resp = MsgBox("Tem certeza que deseja excluir o pedido?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo, "ATENÇÃO")
                If resp = vbYes Then
                    'Armazena o ID para ser apagado do banco de dados
                    Dim valor_chave = dgv_pedidos.Rows(e.RowIndex).Cells("Column1").Value

                    'Remover a linha do DataGridView
                    dgv_pedidos.Rows.RemoveAt(e.RowIndex)

                    'Chama a função para apagar o registro do banco
                    excluir_registro(valor_chave)
                End If
            Else
                'Se a linha estiver vazia, exibe a mensagem de erro
                MsgBox("Não é possível excluir uma linha vazia.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ERRO")
            End If
        End If

        'Verifica se o clique ocorreu na coluna "Alterar Status"
        If e.ColumnIndex = dgv_pedidos.Columns("Column8").Index AndAlso e.RowIndex >= 0 Then
            'Verifica se a linha está vazia
            Dim id As Object = dgv_pedidos.Rows(e.RowIndex).Cells("Column1").Value
            If id IsNot Nothing AndAlso Not String.IsNullOrEmpty(id.ToString()) Then
                'Armazena o número do ID da linha clicada
                Dim valor_chave = Convert.ToInt32(dgv_pedidos.Rows(e.RowIndex).Cells("Column1").Value)
                'Realiza a consulta SQL
                sql = "select status from tb_pedidos where id = " & valor_chave
                rs = db.Execute(sql)

                'Se a consulta retornar e for igual a "EM PREPARO"...
                If rs IsNot Nothing AndAlso Not rs.EOF AndAlso rs.Fields("status").Value IsNot Nothing AndAlso rs.Fields("status").Value.ToString() = "EM PREPARO" Then
                    'O status é atualizado para "PRONTO"
                    sql = "update tb_pedidos set status = 'PRONTO' where id = " & valor_chave
                    rs = db.Execute(sql)
                    carregar_pedidos()
                    'Se a consulta retornar e for igual a "PRONTO"...
                ElseIf rs IsNot Nothing AndAlso Not rs.EOF AndAlso rs.Fields("status").Value IsNot Nothing AndAlso rs.Fields("status").Value.ToString() = "PRONTO" Then
                    'O status é atualizado para "EM PREPARO"
                    sql = "update tb_pedidos set status = 'EM PREPARO' where id = " & valor_chave
                    rs = db.Execute(sql)
                    'Atualiza os pedidos no DataGridView
                    carregar_pedidos()
                End If
            Else
                'Se a linha estiver vazia, exibe a mensagem de erro
                MsgBox("Não é possível alterar uma linha vazia.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "ERRO")
            End If
        End If
    End Sub

    'Click da CheckBox de Mostrar Senha
    Private Sub cb_mostrar_senha_CheckedChanged(sender As Object, e As EventArgs) Handles cb_mostrar_senha.CheckedChanged
        If cb_mostrar_senha.Checked Then
            txt_criar_senha.PasswordChar = ""
        Else
            txt_criar_senha.PasswordChar = "•"
        End If
    End Sub

    'Click na CheckBox de mostrar senha na tela de login
    Private Sub cb_mostrar_senha_login_CheckedChanged(sender As Object, e As EventArgs) Handles cb_mostrar_senha_login.CheckedChanged
        If cb_mostrar_senha_login.Checked Then
            txt_senha.PasswordChar = ""
        Else
            txt_senha.PasswordChar = "•"
        End If
    End Sub
End Class
