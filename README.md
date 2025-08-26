# 🚨 Alertas de Não Conformidades & Lista Mestra (Google Apps Script)

Este repositório contém um único arquivo `.gs` com **todas as funções originais** fornecidas, organizadas em três blocos independentes:

1. **AE – Não Conformidades (Auditorias Externas)**  
   - Função principal: `enviarAlertaNC()`
   - Funções de teste: `enviarAlertaNC_manual()` e `validarDatas_AE()` (exposta com sufixo para evitar conflito)

2. **AI – Não Conformidades (Auditorias Internas)**  
   - Função principal: `enviarAlertaNCAi()`
   - Funções de teste: `enviarAlertaNC_manualAi()` e `validarDatas_AI()` (exposta com sufixo)

3. **Lista Mestra Interno – Documentos**  
   - Função principal: `alertaListaMestra()`
   - Funções de teste: `alertaListaMestra_manual()` e `validarDatas_LM()` (exposta com sufixo)

> 🔒 **Importante:** As funções internas auxiliares (ex.: `parseDateFromCell`, `formatMensagem`, etc.) foram mantidas **iguais** às originais. Para evitar conflitos de nomes no mesmo arquivo, elas ficam isoladas dentro de escopos (IIFEs). As funções principais continuam **globais**, prontas para uso em gatilhos.

---

## ✉️ Destinatários e Assuntos
Os endereços de e-mail e assuntos usados nas funções **são exatamente os mesmos** enviados por você:
- **AE / AI:** assunto `"Alerta de Não Conformidades - Vencendo em 30 ou 15 Dias"`
- **Lista Mestra:** assunto `"Alerta de Não Conformidades - Vencendo em 7 ou 14 Dias"`
- Destinatários: os mesmos (`to` e `bcc`) conforme o seu código.

Se desejar personalizar no futuro, edite diretamente dentro de cada função `MailApp.sendEmail({...})`.

---

## ▶️ Como usar no Google Apps Script

1. Abra sua planilha no **Google Sheets**.
2. Vá em **Extensões → Apps Script**.
3. Crie um projeto e **cole o conteúdo de `alertas_unificado.gs`** (arquivo abaixo).
4. Salve.
5. Execute manualmente uma das funções principais para conceder permissões:
   - `enviarAlertaNC` (AE)
   - `enviarAlertaNCAi` (AI)
   - `alertaListaMestra` (Lista Mestra)

---

## ⏰ Agendamento (gatilhos)

No editor do Apps Script:
1. Clique em **Relógio (Gatilhos)**.
2. **Adicionar gatilho** e selecione a função desejada (ex.: `enviarAlertaNC`).
3. Configure a periodicidade (ex.: diária) e o horário.

---

## 🧪 Funções de teste / validação de dados

- **AE:**  
  - `enviarAlertaNC_manual()`  
  - `validarDatas_AE()` *(mesma lógica da sua `validarDatas`, exposta com sufixo para não colidir com as demais)*

- **AI:**  
  - `enviarAlertaNC_manualAi()`  
  - `validarDatas_AI()`

- **Lista Mestra:**  
  - `alertaListaMestra_manual()`  
  - `validarDatas_LM()`

Os logs exibirão:
- Linhas processadas
- Alertas encontrados
- Datas inválidas (linha/coluna/valor)

---
