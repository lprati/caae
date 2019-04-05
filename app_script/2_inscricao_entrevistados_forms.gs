/*
 *
 *  Esta função é a responsável por criar um link entre a submissão do formulário e o envio do resumo das respostas para o email da CAAE
 *  Para tornar o formulário funcional, deve-se executar a função pelo editor de scritps
 *  
 *  1) No menu "Selecionar função" acima, clique para exibir a lista de opções e escolha "link_script_to_form_submit"
 *  2) Clique no símbolo de "play", logo ao lado do desenho de um insetinho, para executar a função
 *  3) Só é necessário realizar esse procedimento uma única vez! Execute novamente caso crie um formulário novo. Múltiplas execuções num mesmo formulário
 *     podem acarretar em problemas no script.
 *
 */
function link_script_to_form_submit() {

  // Obtem referêcia para o formulário
  var form = FormApp.getActiveForm();
  
  // Configura o formulário para realizar a função "onFormSubmit" abaixo
  ScriptApp.newTrigger('onFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();
}

/*
 * Função executada ao receber uma resposta.
 * ALTERAR: Enviar apenas última resposta recebida.
 */
function onFormSubmit(e) {
  
  // Abre planilha de respostas e recupera os valores
  var ss = SpreadsheetApp.openById('1NTPb4dYCEbYEDSWYjo2VFxYu6L7nsh_gVOYdTwiPdSs').getSheetByName('Respostas ao formulário 1');
  var range = ss.getDataRange();
  var values = range.getValues();
  
  // Para cada linha da planilha (para cada resposta)
  for (var i = 1; i < values.length; i++) {
    // Inicia uma unordered list (html)
    var message = '<ul>';
    
    // Para cada coluna
    for (var j = 1; j < values[i].length; j++) {
      
      // Se o campo é uma string
      if (typeof(values[i][j]) == "string") {
        // Cria uma lista dos valores separados por vírgula
        // O objetivo é criar uma lista onde cada posição é um dos links para documentos do drive
        // enviados para a resposta [j]
        var splited_values = values[i][j].split(',');  
      }
      // Se o campo não é uma string (ex: num usp)
      else {
        var splited_values = [values[i][j]];
      }
      
      // Uma lista contendo os valores da resposta [j]
      message += values[0][j] + '<ul>';
      for (var k = 0; k < splited_values.length; k++) {
        message += '<li>' + splited_values[k] + '</li>';
      }
      message += '</ul>'
    }
      
    message += '</ul>';
  
    // Envia o resumo da resposta para o email cadastrado
    MailApp.sendEmail({
    to: "caae.dev@gmail.com",
    subject: values[i][9].toString().split('.')[0],
    htmlBody: message
    });

    // Reseta o corpo do email caso esta função envie múltiplos emails
    message = ''
  }  
}

function onOpen() {
  

}
