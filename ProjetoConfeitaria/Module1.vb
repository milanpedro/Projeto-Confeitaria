Imports System.Data.OleDb
Imports System.Web.UI.WebControls
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports Guna.UI.WinForms

Module Module1
    Public cont, aux_id As Integer
    Public sql, status, massa, tamanho, sabor, resp As String
    Public precoBase As Decimal = 0
    Public precoTamanho As Decimal = 0
    Public precoSabor As Decimal = 0
    Public db As New ADODB.Connection
    Public rs As New ADODB.Recordset
    Public dir_banco = Application.StartupPath & "\banco\banco_confeitaria.mdb"

    'Subrotina de conexão ao banco de dados
    Sub conectar_banco()
        Try
            db = CreateObject("ADODB.Connection")
            db.Open("Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & dir_banco)
            'MsgBox("Conectado!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AVISO")
        Catch ex As Exception
            MsgBox("Erro ao conectar com o banco!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AVISO")
        End Try
    End Sub

    'Subrotina de limpar os campos de login
    Sub limpar_login()
        With Form1
            .txt_senha.Clear()
            .txt_login.Clear()
        End With
    End Sub

    'Subrotina de limpar os campos de cadastro de clientes
    Sub limpar_clientes()
        With Form1
            .txt_nome_cliente.Clear()
            .txt_cpf_cliente.Clear()
            .txt_email_cliente.Clear()
            .txt_celular_cliente.Clear()
            .txt_endereco.Clear()
            .txt_numero.Clear()
            .txt_bairro.Clear()
            .txt_cidade.Clear()
            .txt_uf.Clear()
        End With
    End Sub

    'Subrotina de limpar os campos de cadastro de funcionários
    Sub limpar_func()
        With Form1
            .txt_nome_func.Clear()
            .txt_cpf_func.Clear()
            .txt_email_func.Clear()
            .txt_celular_func.Clear()
            .txt_criar_senha.Clear()
            .cmb_cargo.SelectedItem = Nothing
            .txt_chave.Clear()
        End With
    End Sub

    'Subrotina de limpar os campos do pedido
    Sub limpar_pedido()
        With Form1
            .txt_cpf_pedido.Clear()
            .cmb_massa.SelectedItem = Nothing
            .cmb_sabor.SelectedItem = Nothing
            .cmb_tamanho.SelectedItem = Nothing
            .lbl_preco_massa.Text = "-"
            .lbl_preco_tam.Text = "-"
            .lbl_preco_sabor.Text = "-"
            .lbl_preco_total.Text = "-"
        End With
    End Sub

    'Subrotina de exibir dados no DataGridView
    Sub carregar_pedidos()
        'Conecta ao banco de dados
        conectar_banco()

        'Realiza consulta SQL
        sql = "select * from tb_pedidos order by id asc"
        rs = db.Execute(sql)

        'Limpa as linhas do DataGridView
        Form1.dgv_pedidos.Rows.Clear()

        'Preenche o DataGridView com os dados do banco
        Do While Not rs.EOF
            Form1.dgv_pedidos.Rows.Add(rs.Fields("id").Value, rs.Fields("cliente").Value, rs.Fields("tipo_massa").Value, rs.Fields("sabor").Value, rs.Fields("tamanho").Value, rs.Fields("valor").Value, rs.Fields("status").Value)
            rs.MoveNext()
        Loop
    End Sub

    'Subrotina de excluir registros no banco de dados a partir do botão do DataGridView
    Sub excluir_registro(valor_chave As String)
        'Conecta ao banco de dados
        conectar_banco()

        'Apaga o registro do banco
        sql = "delete * from tb_pedidos where id = " & valor_chave & ""
        rs = db.Execute(sql)

        MsgBox("Registro excluído com sucesso!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCESSO")
    End Sub

    'Redireciona para o cadastro do cliente
    Sub redireciona_cadastro(cpf As String)
        With Form1
            'Redireciona para a tela de cadastro
            .TabControl1.SelectedTab = .TabControl1.TabPages("TabPage3")
            'Preenche a TextBox de CPF com o CPF que a pessoa deseja cadastrar
            .txt_cpf_cliente.Text = cpf
            'Troca o foco para a TextBox de nome
            .txt_nome_cliente.Focus()
        End With
    End Sub
End Module

