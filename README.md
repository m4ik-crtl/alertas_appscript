# Automação de Alertas por E-mail com Google Apps Script

Este projeto utiliza Google Apps Script para automatizar o envio de e-mails de alerta com base em datas de vencimento monitoradas em uma Planilha Google. É uma solução flexível e poderosa para gestão de prazos em diversas áreas, como compliance, gerenciamento de projetos, controle de documentos e muito mais.

## 🚀 Funcionalidades

- **Monitoramento Múltiplo**: Monitore diferentes tipos de tarefas e documentos em abas separadas da mesma planilha.
- **Alertas Configuráveis**: Defina com quantos dias de antecedência os alertas devem ser enviados (ex: 7, 15, 30 dias).
- **Regras de Negócio**: Implemente lógicas customizadas, como não enviar alertas para tarefas já marcadas como "concluídas".
- **Design Reutilizável**: O código é estruturado com uma classe (`AlertGenerator`) que pode ser facilmente estendida para novos tipos de alertas.
- **Templates de E-mail em HTML**: Envia e-mails formatados e profissionais para os destinatários.
- **Fácil de Configurar**: Todas as configurações (ID da planilha, nomes das abas, mapeamento de colunas, etc.) estão centralizadas no topo do arquivo para fácil customização.

## ⚙️ Como Funciona

O projeto é composto por um único arquivo `Code.gs` que contém:

1.  **Objetos de Configuração**: Para cada tipo de alerta (ex: `CONFIG_NC_A`, `CONFIG_DOCS`), um objeto define todas as regras: nome da aba, colunas a serem lidas, dias para alertar, etc.
2.  **Funções de Gatilho**: Funções como `runNCTypeA_Alerts()` servem como ponto de entrada para iniciar a verificação de um tipo específico de alerta.
3.  **Classe `AlertGenerator`**: O "motor" do script. Esta classe reutilizável lê os dados da planilha, aplica as regras definidas na configuração, encontra os itens que precisam de alerta e monta o e-mail em HTML.

## 🔧 Configuração do Projeto

Siga os passos abaixo para implementar esta automação na sua própria Planilha Google.

### 1. Crie uma Planilha Google

- Crie uma nova planilha ou use uma existente.
- Organize os dados em abas, como "NaoConformidades_TipoA", "ListaMestra_Documentos", etc.

### 2. Crie o Projeto Google Apps Script

- Na sua Planilha Google, vá em `Extensões` > `Apps Script`.
- Apague todo o código de exemplo no arquivo `Code.gs`.
- Copie e cole todo o conteúdo do arquivo `Code.gs` deste repositório.

### 3. Configure as Variáveis Globais

No topo do arquivo `Code.gs`, você encontrará a seção `CONFIGURAÇÃO GLOBAL`. Preencha as variáveis com os seus dados:

- **`SPREADSHEET_ID`**: O ID da sua planilha. Você pode encontrá-lo na URL da planilha (ex: `.../spreadsheets/d/`**`1a2b3c4d5e6f7g8h9i0j`**`/edit`).
- **`DEFAULT_DESTINATION_EMAIL`**: A lista de e-mails padrão para receber os alertas.
- **`DEFAULT_BCC_EMAIL`**: O e-mail que receberá uma cópia oculta (útil para registro).

### 4. Customize os Objetos de Configuração

Para cada alerta que você deseja criar, ajuste o objeto de configuração correspondente (ex: `CONFIG_NC_A`).

- **`SHEET_NAME`**: O nome exato da aba na sua planilha.
- **`ALERT_DAYS`**: Um array com os dias de antecedência para enviar o alerta (ex: `[7, 15, 30]`).
- **`START_ROW`**: O número da linha onde seus dados realmente começam (ignorando o cabeçalho).
- **`COLUMNS`**: Mapeie os índices das colunas (lembre-se que a coluna A é o índice `0`, B é `1`, e assim por diante).

### 5. Configure os Gatilhos (Triggers)

Para que o script rode automaticamente, você precisa criar gatilhos.

- No editor do Apps Script, clique no ícone de relógio ("Acionadores") no menu à esquerda.
- Clique em `+ Adicionar acionador`.
- Configure um acionador para cada função de alerta que você deseja executar (ex: `runNCTypeA_Alerts`).
  - **Função a ser executada**: `runNCTypeA_Alerts`
  - **Tipo de evento**: `Baseado em tempo`
  - **Tipo de acionador baseado em tempo**: `Acionador diário`
  - **Horário**: Escolha um horário para a execução (ex: `8h às 9h`).
- Salve o acionador. Você precisará autorizar o script a acessar suas planilhas e enviar e-mails em seu nome.
- Repita o processo para as outras funções de alerta (`runNCTypeB_Alerts`, `runDocs_Alerts`).

## 🎨 Customização

### Mapeamento de Colunas

A parte mais importante da customização é o mapeamento de colunas dentro de cada objeto de configuração. Ajuste os números dos índices para que correspondam à estrutura da sua planilha.

### Adicionando Novos Alertas

Para criar um novo tipo de alerta:
1.  Copie um dos objetos de configuração existentes (ex: `CONFIG_DOCS`).
2.  Renomeie-o (ex: `CONFIG_FORNECEDORES`).
3.  Ajuste todas as propriedades (nome da aba, colunas, etc.) para o novo alerta.
4.  Crie uma nova função de gatilho (ex: `runFornecedores_Alerts()`) que instancie e execute o `AlertGenerator` com a nova configuração.
5.  Adicione um novo acionador para esta função.

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.
