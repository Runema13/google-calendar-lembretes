# 📅 Lembretes de Vencimento - Google Calendar

Script automatizado para criar lembretes de vencimento de seguros no Google Calendar a partir de uma planilha Google Sheets.

## ✨ Funcionalidades

- ✅ **Agrupa múltiplos seguros** do mesmo cliente na mesma data em um único evento
- ✅ **Previne duplicatas** automaticamente
- ✅ **Lembretes por email** (6 dias antes, 2 dias antes e no dia)
- ✅ **Notificações pop-up** no celular/desktop
- ✅ **Ignora datas passadas** automaticamente
- ✅ **Logs detalhados** de execução

## 📋 Estrutura da Planilha

| Fim de Vigência | Segurados    |
|-----------------|--------------|
| 15/12/2025      | João Silva   |
| 15/12/2025      | João Silva   |
| 20/12/2025      | Maria Santos |

**Resultado:** João Silva terá 1 evento com "(2 seguros)" na descrição.

## 🚀 Como Usar

1. Abra sua planilha no Google Sheets
2. Vá em **Extensões > Apps Script**
3. Cole o código do arquivo `codigo.js`
4. Substitua o email `SEU_EMAIL_AQUI@gmail.com` pelo seu calendário
5. Execute a função `enviarNotificacoesECriarEventos()`
6. Autorize as permissões necessárias

## 🔧 Funções Disponíveis

### `enviarNotificacoesECriarEventos()`
Função principal que lê a planilha e cria os eventos.

### `limparEventosDuplicados()`
Remove eventos duplicados do calendário (útil para limpeza).

## 📝 Logs

Para ver o que aconteceu:
- No Apps Script, vá em **Execuções**
- Veja quantos eventos foram criados/ignorados

## 🛠️ Tecnologias

- Google Apps Script (JavaScript)
- Google Sheets API
- Google Calendar API

## 📄 Licença

MIT License - Sinta-se livre para usar e modificar!

## 👤 Autor

Criado por [@Runema13](https://github.com/Runema13)
```

### **2. Adicione um arquivo `.gitignore`**
```
# Logs
*.log

# Arquivos temporários
*.tmp
~$*

# Sistema
.DS_Store
Thumbs.db

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo [LICENSE](LICENSE) para mais detalhes.
