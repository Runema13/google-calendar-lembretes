# 📅 Lembretes de Vencimento - Google Calendar

Script em Google Apps Script que cria lembretes automáticos no Google Calendar para vencimentos de seguros.

## Funcionalidades
✅ Agrupa múltiplos seguros do mesmo cliente na mesma data
✅ Evita duplicatas automaticamente
✅ Envia lembretes por email (6 dias, 2 dias e no dia)
✅ Notificações pop-up

## Como usar
1. Abra sua planilha Google Sheets
2. Vá em **Extensões > Apps Script**
3. Cole o código de `codigo.js`
4. Execute a função `enviarNotificacoesECriarEventos()`

## Estrutura da planilha
| Fim de Vigência | Segurados |
|-----------------|-----------|
| 15/12/2025      | João Silva |
| 15/12/2025      | João Silva |
