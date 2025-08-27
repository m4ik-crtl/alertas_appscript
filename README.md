# 📊 Alertas Automáticos por Vencimento para Planilhas Google

Este é um script para o Google Apps Script que automatiza o monitoramento de datas em qualquer Planilha Google. Ele envia um e-mail de alerta customizável quando um item se aproxima de sua data de vencimento, eliminando a necessidade de verificações manuais e prevenindo a perda de prazos.

Ideal para gestão de projetos, controle de qualidade, acompanhamento de contratos, inventário, ou qualquer processo que dependa de prazos.

## ✨ Funcionalidades Principais

* **Notificações Automáticas:** Envia e-mails automaticamente em dias e horários pré-configurados.
* **Altamente Customizável:** Configure facilmente quais colunas, prazos e destinatários devem ser notificados.
* **Múltiplas Colunas de Data:** Monitore vários prazos (ex: "Prazo da Ação", "Data de Abrangência") na mesma linha da planilha.
* **Relatório Profissional em HTML:** Gera um e-mail com uma tabela clara e organizada contendo todos os itens que exigem atenção.
* **Validador de Datas:** Inclui uma função para ajudar a encontrar células com formatos de data inválidos que poderiam causar erros.
* **Autônomo e Seguro:** Roda inteiramente no ambiente seguro do Google, sem necessidade de serviços ou servidores externos.

## 📧 O E-mail de Alerta

O script gera um relatório em HTML limpo e responsivo, enviado diretamente para a caixa de entrada dos responsáveis.

> 💡 **Dica:** Adicione uma captura de tela do e-mail real aqui para mostrar o resultado final aos seus usuários!

A estrutura do e-mail se parece com esta:

| NC  | Desvio                                | Ação                                  | Status      | Data       | Data Limite | Responsável   |
| --- | ------------------------------------- | ------------------------------------- | ----------- | ---------- | ----------- | ------------- |
| 101 | Falha no procedimento X               | Revisar o documento Y                 | Prazo ação  | 11/09/2025 | 11/09/2025  | João da Silva |
| 102 | Equipamento Z sem calibração          | Agendar calibração com fornecedor     | Abrangência | 26/09/2025 | 26/09/2025  | Maria Souza   |

## 🚀 Começando

Siga os passos abaixo para implementar o script em sua Planilha Google.

### Pré-requisitos

* Uma Conta Google (Gmail, Google Workspace, etc.).
* Uma Planilha Google com dados organizados em colunas, contendo pelo menos uma coluna com datas de vencimento.

### Instalação e Configuração

1.  **Abra sua Planilha Google.**
2.  No menu superior, clique em **Extensões** > **Apps Script**.
3.  Apague todo o conteúdo do arquivo `Código.gs` que aparece.
4.  **Copie e cole** o conteúdo do arquivo `codigo.gs` deste repositório no editor do Apps Script.
5.  **Configure o script:** No topo do arquivo, localize a seção `--- CONFIGURAÇÕES DO USUÁRIO ---` e altere as variáveis para corresponder às suas necessidades.

    ```javascript
    // --- CONFIGURAÇÕES DO USUÁRIO (ALTERE OS VALORES ABAIXO) ---

    // 1. Informe o nome exato da aba que o script deve monitorar.
    var NOME_DA_ABA = "Controle de Prazos";

    // 2. Informe a partir de qual linha começam os seus dados (ignorando o cabeçalho).
    var LINHA_INICIAL = 5;

    // 3. Defina os dias de antecedência para o envio do alerta.
    var DIAS_PARA_ALERTA = [15, 30];

    // 4. Configure o destinatário e o assunto do e-mail.
    var EMAIL_DESTINATARIO = "gerente.projetos@suaempresa.com";
    var ASSUNTO_DO_EMAIL = "Alerta de Prazos - Itens Vencendo em 30 ou 15 Dias";
    var TITULO_DO_EMAIL = "📌 Alerta de Itens com Vencimento Próximo";

    // 5. Mapeie as colunas da sua planilha (A=0, B=1, C=2, ...).
    var colNumero = 0;
    var colDesvio = 1;
    var colAcao = 10;
    var colResponsavel = 13;

    // 6. Mapeie as COLUNAS DE DATA que você deseja monitorar.
    var colPrazoAcao = 11;
    var colAbrangencia = 14;
    ```

6.  **Salve o projeto** clicando no ícone de disquete.
7.  **Autorize o script:**
    * No menu suspenso de funções, selecione `enviarAlertaNC_manual` e clique em **Executar**.
    * O Google pedirá sua permissão. Clique em **Revisar permissões**.
    * Escolha sua conta Google.
    * Você verá um aviso de "app não verificado". Clique em **Avançado** e depois em **Acessar [Nome do Projeto] (não seguro)**.
    * Revise as permissões e clique em **Permitir**.

### Uso

#### Execução Automática (Acionador)

Para que o script rode todos os dias sem intervenção manual:

1.  No menu à esquerda do editor do Apps Script, clique em **Acionadores** (ícone de relógio).
2.  Clique no botão **+ Adicionar acionador**.
3.  Configure o acionador da seguinte forma:
    * Função a ser executada: `enviarAlertaNC`
    * Implantação: `Principal`
    * Fonte do evento: `Baseado em tempo`
    * Tipo de acionador: `Contador de dias`
    * Horário do dia: `Entre 8h e 9h` (ou o horário de sua preferência).
4.  Clique em **Salvar**.

#### Execução Manual (Para Testes)

Você pode executar o script a qualquer momento para testar:

1.  Abra o editor do Apps Script.
2.  Selecione a função `enviarAlertaNC_manual` no menu.
3.  Clique em **Executar**.
4.  Para ver o que aconteceu, vá em **Registros de execução**.

## 🔧 Customização Avançada

Para monitorar uma terceira coluna de data (ou mais), siga este padrão:

1.  **Adicione uma nova variável** de coluna na seção de configuração:
    ```javascript
    var colRevisaoFinal = 16; // Exemplo: monitorando a coluna Q
    ```
2.  **Adicione um novo bloco `if`** dentro do `for` loop na função principal, seguindo o modelo dos existentes:
    ```javascript
    // Converte a nova data
    var revisaoFinalRaw = linha[colRevisaoFinal];
    var revisaoFinal = parseDateFromCell(revisaoFinalRaw);

    // Adiciona a validação de data inválida
    if (revisaoFinalRaw && !revisaoFinal) invalidates.push({ row: i + 1, col: colRevisaoFinal + 1, value: revisaoFinalRaw });

    // 3. Verifica a data de "Revisão Final"
    if (revisaoFinal && isWithinDaysRange(revisaoFinal, hoje, DIAS_PARA_ALERTA)) {
      mensagens.push(formatMensagem(numero, desvio, acao, revisaoFinal, "Revisão Final", responsavel, revisaoFinal));
      matchesCount++;
    }
    ```

## ⚠️ Solução de Problemas

* **O e-mail não foi enviado?**
    * Verifique os **Registros de execução** no Apps Script para ver se há erros.
    * Confirme se o `NOME_DA_ABA` no script corresponde exatamente ao nome da aba na sua planilha.
    * Execute a autorização novamente.

* **As datas não estão sendo reconhecidas?**
    * Certifique-se de que as datas na sua planilha estão em um formato válido (ex: `DD/MM/AAAA` ou `AAAA-MM-DD`).
    * Use a função de teste `validarDatas_manual` para verificar uma coluna específica por formatos inválidos.

## 📄 Licença

Este projeto está licenciado sob a Licença MIT. Veja o arquivo `LICENSE` para mais detalhes.
