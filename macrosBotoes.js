function copiarEAdicionarNaPagina2() {
  // Abre a planilha de origem (pendente de cadastro)
  var planilhaOrigem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PENDENTES DE CADASTRO');
  
  // Abre a planilha de destino (BASE CADASTRO)
  var planilhaDestino = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BASE CADASTRO');

  // Encontra a última linha na BASE CADASTO
  var ultimaLinhaDestino = planilhaDestino.getLastRow();

  // Define as linhas e colunas que deseja copiar (exemplo: linhas 1 a 1000, colunas A a J)
  var linhasParaCopiar = 1000;
  var colunasParaCopiar = 1000;

  // Obtém os valores de texto com formatação da planilha de origem a partir da última linha copiada
  var dadosTexto = planilhaOrigem.getRange(ultimaLinhaDestino + 1, 1, linhasParaCopiar, colunasParaCopiar).getRichTextValues();
  
  // Obtém os valores numéricos da planilha de origem a partir da última linha copiada
  var dadosNumericos = planilhaOrigem.getRange(ultimaLinhaDestino + 1, 1, linhasParaCopiar, colunasParaCopiar).getValues();

  // Filtra apenas as linhas que não estão presentes na Página2
  var dadosTextoFiltrados = dadosTexto.filter(function(row) {
    return !planilhaDestino.getRange(1, 1, planilhaDestino.getLastRow(), colunasParaCopiar).getRichTextValues().some(function(existingRow) {
      return JSON.stringify(existingRow) === JSON.stringify(row);
    });
  });

  var dadosNumericosFiltrados = dadosNumericos.filter(function(row) {
    return !planilhaDestino.getRange(1, 1, planilhaDestino.getLastRow(), colunasParaCopiar).getValues().some(function(existingRow) {
      return JSON.stringify(existingRow) === JSON.stringify(row);
    });
  });

  // Adiciona apenas as linhas não copiadas anteriormente na Página2
  planilhaDestino.getRange(ultimaLinhaDestino + 1, 1, dadosTextoFiltrados.length, colunasParaCopiar).setRichTextValues(dadosTextoFiltrados);
  planilhaDestino.getRange(ultimaLinhaDestino + 1, 1, dadosNumericosFiltrados.length, colunasParaCopiar).setValues(dadosNumericosFiltrados);

  // Vai para a Página2
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(planilhaDestino);
}
