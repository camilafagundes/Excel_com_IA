## Código VBA para o Agregador de Dados no Excel

Este documento contém o código VBA (Visual Basic for Applications) que pode ser inserido no seu arquivo `Agregador_de_Dados.xlsx` para adicionar funcionalidades de automação, como a transferência de dados da planilha de entrada para o banco de dados e a limpeza do formulário.

### Como Inserir o Código VBA no Excel:

1.  **Abra o Excel:** Abra o arquivo `Agregador_de_Dados.xlsx`.
2.  **Abra o Editor VBA:** Pressione `Alt + F11` para abrir o Editor do Visual Basic for Applications.
3.  **Insira um Módulo:** No painel esquerdo (Project Explorer), clique com o botão direito em `VBAProject (Agregador_de_Dados.xlsx)`, vá em `Inserir` e selecione `Módulo`.
4.  **Cole o Código:** Copie o código abaixo e cole-o no módulo em branco que foi aberto.
5.  **Crie os Botões:** Volte para a planilha "Entrada de Dados" no Excel. Vá em `Desenvolvedor` > `Inserir` > `Controles de Formulário` > `Botão (Controle de Formulário)`. Desenhe o botão onde desejar. Ao soltar o botão, uma janela aparecerá para `Atribuir Macro`. Selecione a macro correspondente (ex: `AdicionarDados` ou `LimparFormulario`). Repita para o segundo botão.
6.  **Formatação Condicional:** Para a formatação condicional, você precisará aplicá-la diretamente no Excel, pois o `openpyxl` não suporta a criação de regras de formatação condicional complexas que dependem de fórmulas ou referências relativas da mesma forma que o Excel.

---

### Módulo 1: `MacrosAgregador`

```vba
Option Explicit

Sub AdicionarDados()
    Dim wsEntrada As Worksheet
    Dim wsBanco As Worksheet
    Dim lastRowEntrada As Long
    Dim lastRowBanco As Long
    Dim i As Long
    Dim nextID As Long
    
    Set wsEntrada = ThisWorkbook.Sheets("Entrada de Dados")
    Set wsBanco = ThisWorkbook.Sheets("Banco de Dados")
    
    ' Encontra a última linha preenchida no Banco de Dados para determinar o próximo ID
    If wsBanco.Cells(Rows.Count, "A").End(xlUp).Row < 5 Then ' Se não houver dados ou apenas cabeçalho
        nextID = 1
    Else
        nextID = wsBanco.Cells(Rows.Count, "A").End(xlUp).Value + 1
    End If
    
    ' Encontra a próxima linha vazia no Banco de Dados
    lastRowBanco = wsBanco.Cells(Rows.Count, "A").End(xlUp).Row + 1
    If lastRowBanco < 5 Then lastRowBanco = 5 ' Garante que comece na linha 5 se o banco estiver vazio
    
    ' Validação básica dos campos de entrada
    If IsEmpty(wsEntrada.Range("A5").Value) Or _
       IsEmpty(wsEntrada.Range("B5").Value) Or _
       IsEmpty(wsEntrada.Range("D5").Value) Then
        MsgBox "Por favor, preencha todos os campos obrigatórios (Data, Categoria, Valor).", vbCritical, "Erro de Validação"
        Exit Sub
    End If
    
    ' Transfere os dados da Entrada para o Banco de Dados
    wsBanco.Cells(lastRowBanco, "A").Value = nextID ' ID
    wsBanco.Cells(lastRowBanco, "B").Value = wsEntrada.Range("A5").Value ' Data
    wsBanco.Cells(lastRowBanco, "C").Value = wsEntrada.Range("B5").Value ' Categoria
    wsBanco.Cells(lastRowBanco, "D").Value = wsEntrada.Range("C5").Value ' Descrição
    wsBanco.Cells(lastRowBanco, "E").Value = wsEntrada.Range("D5").Value ' Valor
    wsBanco.Cells(lastRowBanco, "F").Value = wsEntrada.Range("E5").Value ' Status
    wsBanco.Cells(lastRowBanco, "G").Value = Now ' Data Inserção
    
    ' Formata a coluna de Data no Banco de Dados
    wsBanco.Cells(lastRowBanco, "B").NumberFormat = "dd/mm/yyyy"
    wsBanco.Cells(lastRowBanco, "G").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    
    MsgBox "Dados adicionados com sucesso!", vbInformation, "Sucesso"
    
    ' Limpa o formulário após adicionar os dados
    Call LimparFormulario
End Sub

Sub LimparFormulario()
    Dim wsEntrada As Worksheet
    Set wsEntrada = ThisWorkbook.Sheets("Entrada de Dados")
    
    ' Limpa as células de entrada
    wsEntrada.Range("A5:E5").ClearContents
    
    ' Opcional: Redefinir a seleção para a primeira célula de entrada
    wsEntrada.Range("A5").Select
End Sub

```

### Formatação Condicional (Manual no Excel)

Para realçar campos obrigatórios ou dados inválidos na planilha "Entrada de Dados", siga estes passos no Excel:

1.  **Selecione a Célula/Intervalo:** Por exemplo, selecione `A5:E100` na planilha "Entrada de Dados".
2.  **Acesse Formatação Condicional:** Vá em `Página Inicial` > `Formatação Condicional` > `Nova Regra...`.
3.  **Use uma Fórmula:** Selecione `Usar uma fórmula para determinar quais células devem ser formatadas`.

    *   **Exemplo 1: Realçar campos vazios (obrigatórios):**
        *   **Fórmula:** `=E(OU(CÉLULA("folha",A5)="Entrada de Dados";CÉLULA("folha",A5)="Entrada de Dados");ÉCÉLULA.VAZIA(A5);ÉCÉLULA.VAZIA(B5);ÉCÉLULA.VAZIA(D5))`
        *   **Formato:** Escolha um preenchimento vermelho claro e fonte em negrito.
        *   **Aplicar a:** `=$A$5:$E$100` (ajuste o intervalo conforme necessário).

    *   **Exemplo 2: Realçar valores negativos (se Valor deve ser positivo):**
        *   **Fórmula:** `=D5<0`
        *   **Formato:** Escolha um preenchimento laranja claro.
        *   **Aplicar a:** `=$D$5:$D$100`

    *   **Exemplo 3: Realçar datas futuras:**
        *   **Fórmula:** `=A5>HOJE()`
        *   **Formato:** Escolha um preenchimento amarelo claro.
        *   **Aplicar a:** `=$A$5:$A$100`

Lembre-se de que a formatação condicional é aplicada por regras e pode ser ajustada conforme suas necessidades de feedback visual.


