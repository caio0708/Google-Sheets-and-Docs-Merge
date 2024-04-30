//Variáveis da Planilha de Banco de Dados
let app = SpreadsheetApp;
let spreadsheet = app.getActiveSpreadsheet();
let sheet = spreadsheet.getSheetByName('Nome'); //Nome da página da planilha

//Variável do arquivo de Mala Direta
let docsMalaDireta = DocumentApp.openByUrl("https://docs.google.com/document/d/xxxxxxxxxx/edit"); //link do docs que será trasnformado nos contratos

//Variável do arquivo de Template
let docsTemplate = DocumentApp.openByUrl("https://docs.google.com/document/d/yyyyyyyyyyyyyy/edit"); //template para os contratos

//Função que gera a mala direta
function criarMalaDireta()
{
  docsMalaDireta.getBody().clear();
  let paragrafos = docsTemplate.getBody().getParagraphs();

  let dados = sheet.getRange("A1:H").getValues(); // Definir de acordo com as cédulas que estão sendo utilizadas na planilha

  dados.map((linha, ind, obj) => {
    if (linha[0] != '') {
      paragrafos.map((paragrafo, ind2, obj2) => {
        let textoFormatado = paragrafo.copy().replaceText("{{NOME}}", linha[0]);
        textoFormatado = textoFormatado.replaceText("{{ANIV}}", linha[1]); 
        textoFormatado = textoFormatado.replaceText("{{RG}}", formatarRG(linha[2])); 
        textoFormatado = textoFormatado.replaceText("{{CPF}}", formatarCPF(linha[3]));
        textoFormatado = textoFormatado.replaceText("{{CIDADE}}", linha[4]); 
        textoFormatado = textoFormatado.replaceText("{{CURSO}}", linha[5]); 
        textoFormatado = textoFormatado.replaceText("{{FACUL}}", linha[6]); 
        textoFormatado = textoFormatado.replaceText("{{HORA}}", linha[7]); //se utilizar google forms sera salvo na linha[0] automaticamente, e assim deve mudar o restante do codigo
 
        docsMalaDireta.getBody().appendParagraph(textoFormatado);
      });
      docsMalaDireta.getBody().appendPageBreak();
    }
  });

}

function formatarCPF(cpf) {
  // Converte o CPF para string, caso não seja
  cpf = String(cpf);

  // Remove caracteres não numéricos do CPF
  cpf = cpf.replace(/\D/g, '');

  // Adiciona os pontos e traço ao CPF
  return cpf.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, '$1.$2.$3-$4');
}

function formatarRG(rg) {
  // Converte o RG para string, caso não seja
  rg = String(rg);

  // Remove caracteres não numéricos do RG
  rg = rg.replace(/\D/g, '');

  // Adiciona os pontos e traço ao RG
  return rg.replace(/^(\d{2})(\d{3})(\d{3})(\d{1})$/, '$1.$2.$3-$4');
}


