// FUNÇÃO PRINCIPAL - É o que roda quando você executa
function enviarNotificacoesECriarEventos() {
  
  // TRY-CATCH: Captura erros para não quebrar tudo
  try {
    
    // PEGAR A PLANILHA ATIVA
    // SpreadsheetApp = API do Google Sheets
    // getActiveSpreadsheet() = pega a planilha aberta
    // getActiveSheet() = pega a aba atual
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // PEGAR TODOS OS DADOS DA PLANILHA
    // getDataRange() = pega o intervalo com dados (A1 até última célula preenchida)
    // getValues() = transforma em array (lista) JavaScript
    // Resultado: [[linha1], [linha2], [linha3]...]
    var dataValues = sheet.getDataRange().getValues();
    
    // VALIDAÇÃO: Tem dados além do cabeçalho?
    // dataValues.length = quantidade de linhas
    // Se <= 1, significa que só tem cabeçalho (linha 1)
    if (dataValues.length <= 1) {
      Logger.log('Nenhum dado encontrado na planilha.'); // Escreve no log
      return; // PARA a execução aqui
    }
    
    // PEGAR O CALENDÁRIO
    // CalendarApp = API do Google Calendar
    // getCalendarById() = busca calendário pelo email/ID
    var calendar = CalendarApp.getCalendarById('SEU_EMAIL_AQUI@gmail.com');
    
    // CRIAR DATA DE HOJE (para comparar depois)
    var hoje = new Date(); // new Date() = cria objeto de data/hora atual
    hoje.setHours(0, 0, 0, 0); // Zera horário (meia-noite) para comparar só a data
    
    // VARIÁVEIS PARA CONTAR
    var eventosProcessados = 0;
    var eventosIgnorados = 0;
    
    // ===== PARTE NOVA: AGRUPAR SEGURADOS =====
    
    // OBJETO VAZIO para guardar dados agrupados
    // Objetos em JS são como dicionários: {chave: valor}
    var agrupados = {};
    
    // LOOP: Percorre todas as linhas (começando da linha 2, pulando cabeçalho)
    // i = 1 porque array começa em 0 (linha 0 = cabeçalho)
    for (var i = 1; i < dataValues.length; i++) {
      
      // PEGAR DADOS DA LINHA ATUAL
      // dataValues[i] = linha atual
      // [0] = primeira coluna (data de vencimento)
      // [1] = segunda coluna (nome do segurado)
      var dataVenc = dataValues[i][0];
      var segurado = dataValues[i][1];
      
      // PULAR se dados estão vazios
      // ! = "não" (negação)
      // || = "OU"
      // toString().trim() = converte pra texto e remove espaços nas pontas
      if (!dataVenc || !segurado || segurado.toString().trim() === '') {
        continue; // PULA para próxima iteração do loop
      }
      
      // CONVERTER para Date se não for (às vezes vem como texto)
      // instanceof = verifica se é do tipo Date
      if (!(dataVenc instanceof Date)) {
        dataVenc = new Date(dataVenc);
      }
      
      // VERIFICAR se data é válida E se não passou
      // isNaN() = "is Not a Number" - verifica se é número inválido
      // getTime() = retorna milissegundos desde 1970 (para comparar)
      // < hoje = se for menor que hoje (data passada)
      if (isNaN(dataVenc.getTime()) || dataVenc < hoje) {
        continue; // PULA essa linha
      }
      
      // CRIAR CHAVE ÚNICA: "Nome do Segurado|Data"
      // Exemplo: "João Silva|Mon Dec 15 2025"
      // Isso garante que João com 3 seguros no mesmo dia = mesma chave
      var chave = segurado.toString().trim() + '|' + dataVenc.toDateString();
      
      // SE essa chave NÃO EXISTE ainda no objeto
      if (!agrupados[chave]) {
        // CRIA nova entrada com dados iniciais
        agrupados[chave] = {
          segurado: segurado.toString().trim(),
          data: dataVenc,
          count: 1  // Primeiro seguro = 1
        };
      } else {
        // SE JÁ EXISTE, apenas aumenta o contador
        agrupados[chave].count++; // count++ = count = count + 1
      }
    }
    
    // ===== CRIAR EVENTOS ÚNICOS =====
    
    // FOR...IN: Loop especial para objetos
    // Percorre cada chave do objeto "agrupados"
    for (var chave in agrupados) {
      
      // PEGAR dados dessa chave
      var item = agrupados[chave];
      var segurado = item.segurado;
      var dataVenc = item.data;
      var qtdSeguros = item.count; // Quantos seguros esse pessoa tem na mesma data
      
      // VERIFICAR se já existe evento no calendário
      var tituloEvento = 'Vencimento: ' + segurado;
      // getEventsForDay() = busca eventos naquele dia
      // {search: ...} = filtra pelo título
      var eventosExistentes = calendar.getEventsForDay(dataVenc, {search: tituloEvento});
      
      // SE JÁ EXISTE (.length > 0 = tem elementos na lista)
      if (eventosExistentes.length > 0) {
        Logger.log('Evento já existe para ' + segurado + ' em ' + dataVenc.toLocaleDateString());
        eventosIgnorados++;
        continue; // Pula para próximo
      }
      
      // CONFIGURAR HORÁRIO DO EVENTO
      // new Date(dataVenc) = cria cópia da data para não alterar original
      var dataInicio = new Date(dataVenc);
      dataInicio.setHours(10, 0, 0, 0); // Define: 10:00:00.000
      
      // DATA FIM = 1 hora depois do início
      // getTime() = milissegundos
      // 60 * 60 * 1000 = 60min × 60seg × 1000ms = 1 hora em milissegundos
      var dataFim = new Date(dataInicio.getTime() + 60 * 60 * 1000);
      
      // MONTAR DESCRIÇÃO DO EVENTO
      // \n = quebra de linha
      // + = concatenação (juntar textos)
      var descricao = 'Lembrete de Vencimento\n\n' +
                      'Segurado: ' + segurado + '\n' +
                      'Data de Vencimento: ' + dataVenc.toLocaleDateString('pt-BR') + '\n';
      
      // SE TEM MAIS DE 1 SEGURO, adiciona essa info
      if (qtdSeguros > 1) {
        descricao += 'Quantidade de seguros vencendo: ' + qtdSeguros + '\n';
        // += é o mesmo que: descricao = descricao + '...'
      }
      
      descricao += '\nAção necessária: Entrar em contato para renovação.';
      
      // TRY-CATCH interno: tenta criar evento
      try {
        
        // CRIAR EVENTO NO CALENDÁRIO
        // createEvent(título, início, fim, opções)
        var evento = calendar.createEvent(tituloEvento, dataInicio, dataFim, {
          description: descricao,
          location: 'Medseg Corretora'
        });
        
        // ADICIONAR LEMBRETES
        // addEmailReminder(minutos antes)
        // 6 dias = 6 × 24 horas × 60 minutos = 8640 minutos
        evento.addEmailReminder(6 * 24 * 60);  // 6 dias antes por email
        evento.addEmailReminder(2 * 24 * 60);  // 2 dias antes por email
        evento.addEmailReminder(10);            // 10 minutos antes por email
        
        // addPopupReminder = notificação pop-up no celular/desktop
        evento.addPopupReminder(2 * 24 * 60);  // 2 dias antes
        evento.addPopupReminder(10);            // 10 minutos antes
        
        // MENSAGEM DE SUCESSO
        // Operador ternário: condição ? seVerdadeiro : seFalso
        var msgSeguro = qtdSeguros > 1 ? ' (' + qtdSeguros + ' seguros)' : '';
        Logger.log('✓ Evento criado: ' + segurado + msgSeguro + ' - ' + dataVenc.toLocaleDateString());
        eventosProcessados++;
        
      } catch (e) {
        // SE DEU ERRO ao criar evento
        // e.message = mensagem do erro
        Logger.log('✗ Erro ao criar evento para ' + segurado + ': ' + e.message);
      }
    }
    
    // ===== RESUMO FINAL =====
    
    Logger.log('\n=== RESUMO ===');
    Logger.log('Eventos criados: ' + eventosProcessados);
    Logger.log('Eventos ignorados: ' + eventosIgnorados);
    Logger.log('Total processado: ' + (dataValues.length - 1));
    
    // MOSTRAR ALERTA NA TELA
    // SpreadsheetApp.getUi() = Interface do usuário
    // alert() = mostra caixa de mensagem
    SpreadsheetApp.getUi().alert(
      'Processo concluído!\n\n' +
      'Eventos criados: ' + eventosProcessados + '\n' +
      'Eventos ignorados: ' + eventosIgnorados
    );
    
  } catch (e) {
    // SE DEU ERRO GERAL (no try lá em cima)
    Logger.log('ERRO GERAL: ' + e.message);
    SpreadsheetApp.getUi().alert('Erro: ' + e.message);
  }
}

// ===== FUNÇÃO AUXILIAR (BÔNUS) =====

function limparEventosDuplicados() {
  
  var calendar = CalendarApp.getCalendarById('medsegcorretora@gmail.com');
  var hoje = new Date();
  
  // Data daqui 1 ano
  // 365 dias × 24h × 60min × 60seg × 1000ms
  var futuro = new Date(hoje.getTime() + 365 * 24 * 60 * 60 * 1000);
  
  // PEGAR TODOS EVENTOS entre hoje e 1 ano
  var eventos = calendar.getEvents(hoje, futuro);
  var eventosRemovidos = 0;
  
  // OBJETO para rastrear o que já vimos
  var titulos = {};
  
  // PERCORRER todos eventos
  for (var i = 0; i < eventos.length; i++) {
    var evento = eventos[i];
    var titulo = evento.getTitle();
    var data = evento.getStartTime().toDateString();
    
    // CHAVE ÚNICA: título + data
    var chave = titulo + '|' + data;
    
    // SE JÁ EXISTE essa chave (é duplicata)
    if (titulos[chave]) {
      evento.deleteEvent(); // DELETA o evento
      eventosRemovidos++;
      Logger.log('Duplicata removida: ' + titulo);
    } else {
      // MARCAR que já vimos esse título+data
      titulos[chave] = true;
    }
  }
  
  Logger.log('Total de duplicatas removidas: ' + eventosRemovidos);
  SpreadsheetApp.getUi().alert('Duplicatas removidas: ' + eventosRemovidos);
} 
