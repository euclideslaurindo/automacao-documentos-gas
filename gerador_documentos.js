function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📄 Gerar Documentos')
    .addItem('📝 Gerar Contrato Mudança de UT', 'gerarMudanca')
    .addSeparator()
    .addItem('📝 Gerar Termo de Rescisão (Iniciativa)', 'gerarRescisaoIniciativa')
    .addSeparator()
    .addItem('📝 Gerar Termo de Rescisão (Necessidade)', 'gerarRescisaoNecessidade')
    .addSeparator() 
    .addItem('📝 Gerar Aditivo (Ampliação/Redução)', 'gerarAditivo')
    .addSeparator()
    .addItem('📝 Gerar Declaração de Acúmulo', 'gerarDeclaracaoAcumulo')
    .addSeparator()
    .addItem('📝 Gerar Atualização Cadastral', 'gerarAtualizacaoCadastral') 
    .addSeparator()
    .addItem('📝 Gerar Declaração Contratual', 'gerarDeclaracaoContrato') 
    .addSeparator()
    .addItem('📝 Gerar Ficha Funcional', 'gerarFichaFuncional') 
    .addToUi();
}

const ID_MUDANCA              = '1gg0ilmmOb8iCEtWdkjK2dTAZt6BzxeZApTwQxdcBT-Q'; 
const ID_RESCISAO_INICIATIVA  = '1tjYcSYJJ98N9FzEPXf5vYV4xdMi7vWRgDa6X1y-63ws';
const ID_RESCISAO_NECESSIDADE = '1IqR2F0ArKJezVQfe74DWpqxIdGmb4ZpkNKBFyb6yMLU';
const ID_ADITIVO              = '1HvPG0IV8OhOi3lTePiRQxXdos3yddyD63M2boig-D3o'; 
const ID_DECLARACAO_ACUMULO   = '1x9ytB1FuS60-YqVXX_L3OO2eVc5yD2LdExk8xoKJzQc'; 
const ID_ATUALIZACAO_CADASTRO = '1SObZUK-wX6zoKKeMcePtS3NJv92T_5QlA3H1p84Qgks'; 
const ID_DECLARACAO_CONTRATO  = '19rOUbs-8zXMziJU6BDydydC16T4V7dPvziUk_yc9pd0'; 
const ID_FICHA_FUNCIONAL      = '1gZyCwnxf0PC537eDhvzFJtlJHdaLoRHYsaQ18QdGVn4'; 


const TABELA_VALORES = {
  "50":  1216.94,
  "40":  2008.12, 
  "150": 3650.83,
  "180": 4122.51,
  "200": 4867.77
};

function gerarMudanca() { executarGeracao(ID_MUDANCA, 'Mudança de UT', true); } 
function gerarRescisaoIniciativa() { executarGeracao(ID_RESCISAO_INICIATIVA, 'Rescisão (Iniciativa)', false); }
function gerarRescisaoNecessidade() { executarGeracao(ID_RESCISAO_NECESSIDADE, 'Rescisão (Necessidade)', false); }
function gerarAditivo() { executarGeracao(ID_ADITIVO, 'Aditivo de CH', true); } 
function gerarDeclaracaoAcumulo() { executarGeracao(ID_DECLARACAO_ACUMULO, 'Declaração de Acúmulo', false); }
function gerarAtualizacaoCadastral() { executarGeracao(ID_ATUALIZACAO_CADASTRO, 'Atualização Cadastral', false); }
function gerarDeclaracaoContrato() { executarGeracao(ID_DECLARACAO_CONTRATO, 'Declaração Contratual', false); }
function gerarFichaFuncional() { executarGeracao(ID_FICHA_FUNCIONAL, 'Ficha Funcional', false); }
 
function ehDataValida(texto) {
  if (!texto) return false;
  const t = String(texto).trim();
  if (t.length >= 8 && (t.includes('/') || t.includes('-')) && !/[a-zA-Z]/.test(t)) return true;
  return false;
}

// função principal
function executarGeracao(idDoModelo, nomePadrao, deveAtualizarBase) {
  const ui = SpreadsheetApp.getUi();
  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  // tela
  const aba = planilha.getSheetByName("AtualizaçõesDespachos");
  if (!aba) { ui.alert("Erro: Aba 'AtualizaçõesDespachos' não encontrada."); return; }
  
  const linha = aba.getActiveCell().getRow();
  if (linha <= 1) { ui.alert("⚠️ Selecione uma linha válida."); return; }

  // primeira coleta da aba de despachos
  const dados = {
    nome:          aba.getRange("F" + linha).getDisplayValue(),
    cpf:           aba.getRange("AC" + linha).getDisplayValue().trim(), 
    matricula:     aba.getRange("B" + linha).getDisplayValue(),
    numFuncional:  aba.getRange("C" + linha).getDisplayValue(),
    numContrato:   aba.getRange("D" + linha).getDisplayValue(),
    cargo:         aba.getRange("G" + linha).getDisplayValue(), 
    escola:        aba.getRange("N" + linha).getDisplayValue(),
    municipioEsc:  aba.getRange("AJ" + linha).getDisplayValue(), 
    dataInicio:    aba.getRange("M" + linha).getDisplayValue(), 
    cargaAtual:    aba.getRange("H" + linha).getDisplayValue(), 
    novaCarga:     aba.getRange("O" + linha).getDisplayValue(), 
    valor:         aba.getRange("AI" + linha).getDisplayValue(), 
    projeto:       aba.getRange("I" + linha).getDisplayValue(),
    disciplina:    aba.getRange("J" + linha).getDisplayValue(),
    formacao:      "",
    nivel:         aba.getRange("AG" + linha).getDisplayValue(),
    rg:            aba.getRange("AA" + linha).getDisplayValue(),
    orgao:         aba.getRange("AB" + linha).getDisplayValue(),
    municipioRes:  aba.getRange("AE" + linha).getDisplayValue(),
    estadoCivil:   aba.getRange("AD" + linha).getDisplayValue(),

//campos vazios
    dataNasc: "", endereco: "", numero: "", bairro: "", cep: "", ufEndereco: "PE", 
    dtEmissaoRG: "", telefone: "", naturalidade: "", pai: "", mae: "", 
    ctps: "", serieCtps: "", ufCtps: "", pis: "", titulo: "", zona: "", secao: "",
    reservista: "", serieReservista: "", regiaoReservista: "",
    agencia: "", conta: "", vigencia: "", email: "",
  };

  if (dados.cargo && dados.cargo.toUpperCase().includes("PROF")) dados.cargo = "PROFESSOR";
  
//cálculo de valor
  let valorParaBase = null; 
  
  if (dados.novaCarga && dados.novaCarga.trim() !== "") {
      const cargaLimpa = dados.novaCarga.toString().replace(/[^0-9]/g, "");
      
      if (TABELA_VALORES[cargaLimpa]) {
          const novoValorNum = TABELA_VALORES[cargaLimpa];
          const novoValorFormatado = novoValorNum.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
          
          dados.valor = novoValorFormatado; 
          valorParaBase = novoValorNum;     
          
          aba.getRange("AI" + linha).setValue(novoValorNum);
      }
  }
  
  if (!valorParaBase && dados.valor) {
      valorParaBase = parseFloat(dados.valor.toString().replace("R$", "").replace(/\./g, "").replace(",", ".").trim());
  }

  // prepara Busca
  const cpfBusca = parseInt(dados.cpf.replace(/[^0-9]/g, ""), 10);
  const nomeBuscaUpper = dados.nome.trim().toUpperCase();
  dados.cpf = dados.cpf.replace(/[^0-9]/g, "").padStart(11, "0");

  dados.valorExtenso = escreverPorExtenso(dados.valor);
  let cargaParaDoc = (dados.novaCarga && dados.novaCarga.trim() !== "") ? dados.novaCarga : dados.cargaAtual;

  // busca de aba contratos
  const abaContratos = planilha.getSheetByName("Contratos");

  if (abaContratos) {
    const dadosGerais = abaContratos.getDataRange().getDisplayValues(); 
    const cabecalho = dadosGerais[0];

    //mapeamento
    let indicesDataNasc = [17]; 
    let col = { Matricula: 0, Funcional: 1, CPF: 11 }; 
    
    for (let c = 0; c < cabecalho.length; c++) {
      const nome = String(cabecalho[c]).toUpperCase().trim();

      if (nome.includes("NASCIMENTO") && !nome.includes("FILHO") && !nome.includes("CONJUGE")) {
          if (c !== 17) indicesDataNasc.push(c);
      }
      if (nome === "NOME" || nome === "NOME DO SERVIDOR") col.Nome = c;
      if (nome.includes("CPF") && !nome.includes("VALIDAÇÃO")) col.CPF = c;
      if (nome.includes("MATRÍCULA") || nome.includes("MATRICULA")) col.Matricula = c;
      if (nome.includes("FUNCIONAL")) col.Funcional = c;
      if (nome === "RG" || nome === "IDENTIDADE") col.RG = c;
      if (nome.includes("ORGAO") || nome.includes("ÓRGÃO")) col.Orgao = c;
      if (nome.includes("EMISSÃO") && (nome.includes("RG") || nome.includes("IDENTIDADE"))) col.DtEmissaoRG = c;
      if (col.DtEmissaoRG === undefined && nome === "DATA DE EMISSÃO") col.DtEmissaoRG = c;
      if ((nome.includes("ENDEREÇO") || nome.includes("LOGRADOURO")) && !nome.includes("UF") && !nome.includes("MUNICÍPIO") && !nome.includes("MUNICIPIO")) {
          col.Endereco = c;
      }
      if (nome === "Nº" || nome === "NUMERO") col.Numero = c;
      if (nome.includes("BAIRRO")) col.Bairro = c;
      if (nome.includes("CEP")) col.CEP = c;
      if (nome.includes("MUNICÍPIO") || nome.includes("CIDADE")) col.MunicipioRes = c;
      if (nome.includes("UF ENDEREÇO") || nome.includes("UF")) col.UFEnd = c;
      if (nome.includes("PAI")) col.Pai = c;
      if (nome.includes("MÃE") || nome.includes("MAE")) col.Mae = c;
      if ((nome.includes("CURSO") || nome.includes("FORMAÇÃO")) && !nome.includes("HABILITAÇÃO") && !nome.includes("HABILITACAO")) {
          col.Curso = c;
      }
      if (nome.includes("NIVEL") || nome.includes("NÍVEL")) col.Nivel = c;
      if (nome.includes("DISCIPLINA")) col.Disciplina = c;
      if ((nome.includes("TITULO") || nome.includes("TÍTULO") || nome.includes("ELEITOR")) && !nome.includes("VALIDAÇÃO") && !nome.includes("VALIDACAO") && !nome.includes("ZONA") && !nome.includes("SEÇÃO")) {
          col.Titulo = c;
      }
      if (nome.includes("ZONA")) col.Zona = c;
      if (nome.includes("SEÇÃO") || nome.includes("SECAO")) col.Secao = c;
      if (nome.includes("PIS")) col.PIS = c;
      if (nome.includes("CTPS") && !nome.includes("SÉRIE") && !nome.includes("UF")) col.CTPS = c;
      if (nome.includes("SÉRIE CTPS")) col.SerieCTPS = c;
      if (nome.includes("UF CTPS")) col.UfCtps = c;
      if (nome.includes("RESERVISTA") && !nome.includes("SÉRIE") && !nome.includes("REGIÃO") && !nome.includes("VALIDAÇÃO") && !nome.includes("VALIDACAO")) {
          col.Reservista = c;
      }
      if (nome.includes("SÉRIE RESERVISTA")) col.SerieReservista = c;
      if (nome.includes("REGIÃO")) col.RegiaoReservista = c; 
      if (nome.includes("AGÊNCIA")) col.Agencia = c;
      if (nome.includes("CONTA")) col.Conta = c;
      if (nome.includes("NATURALIDADE")) col.Naturalidade = c;
      if (nome.includes("TELEFONE")) col.Telefone = c;
      if ((nome.includes("VIGÊNCIA") || nome.includes("VIGENCIA") || nome.includes("VENCIMENTO CONTRATO")) && !nome.includes("INICIO") && !nome.includes("INÍCIO")) {
        col.Vigencia = c;
      }
    }

    if (col.Nome === undefined) col.Nome = 10;
    if (col.Vigencia === undefined) col.Vigencia = 78;

    // busca de baixo pra cima
    let encontrou = false;

    for (let i = dadosGerais.length - 1; i >= 1; i--) {
      let linhaDados = dadosGerais[i];
      
      let cpfLinha = String(linhaDados[col.CPF]).replace(/[^0-9]/g, "");
      let cpfMatch = (parseInt(cpfLinha, 10) === cpfBusca);
      let nomeLinha = String(linhaDados[col.Nome]).trim().toUpperCase();
      let nomeMatch = (nomeLinha === nomeBuscaUpper);

      if (cpfMatch || nomeMatch) {
        
        let dataValidaEncontrada = null;
        for (let idxNasc of indicesDataNasc) {
            let valor = linhaDados[idxNasc];
            if (ehDataValida(valor)) {
                dataValidaEncontrada = valor;
                break; 
            }
        }

        if (dataValidaEncontrada || cpfMatch || nomeMatch) {
            encontrou = true;
            if (dataValidaEncontrada) dados.dataNasc = dataValidaEncontrada;

            // extração
            if ((!dados.matricula || dados.matricula == "") && linhaDados[col.Matricula]) dados.matricula = linhaDados[col.Matricula];
            if ((!dados.numFuncional || dados.numFuncional == "") && linhaDados[col.Funcional]) dados.numFuncional = linhaDados[col.Funcional];
            if (col.RG !== undefined) dados.rg = linhaDados[col.RG];
            if (col.Orgao !== undefined) dados.orgao = linhaDados[col.Orgao];
            if (col.DtEmissaoRG !== undefined) dados.dtEmissaoRG = linhaDados[col.DtEmissaoRG];
            if (col.Pai !== undefined) dados.pai = linhaDados[col.Pai];
            if (col.Mae !== undefined) dados.mae = linhaDados[col.Mae];
            if (col.Endereco !== undefined) dados.endereco = linhaDados[col.Endereco];
            if (col.Numero !== undefined) dados.numero = linhaDados[col.Numero];
            if (col.Bairro !== undefined) dados.bairro = linhaDados[col.Bairro];
            if (col.CEP !== undefined) dados.cep = linhaDados[col.CEP];
            if (col.MunicipioRes !== undefined) dados.municipioRes = linhaDados[col.MunicipioRes];
            if (col.UFEnd !== undefined) dados.ufEndereco = linhaDados[col.UFEnd];
            if (col.Telefone !== undefined) dados.telefone = linhaDados[col.Telefone];
            if (col.Naturalidade !== undefined) dados.naturalidade = linhaDados[col.Naturalidade];
            if ((!dados.formacao || dados.formacao == "") && col.Curso !== undefined) dados.formacao = linhaDados[col.Curso];
            if ((!dados.nivel || dados.nivel == "") && col.Nivel !== undefined) dados.nivel = linhaDados[col.Nivel];
            if ((!dados.disciplina || dados.disciplina == "") && col.Disciplina !== undefined) dados.disciplina = linhaDados[col.Disciplina];
            if (col.Titulo !== undefined) dados.titulo = linhaDados[col.Titulo];
            if (col.Zona !== undefined) dados.zona = linhaDados[col.Zona];
            if (col.Secao !== undefined) dados.secao = linhaDados[col.Secao];     
            if (col.PIS !== undefined) dados.pis = linhaDados[col.PIS];
            if (col.CTPS !== undefined) dados.ctps = linhaDados[col.CTPS];
            if (col.SerieCTPS !== undefined) dados.serieCtps = linhaDados[col.SerieCTPS];
            if (col.UfCtps !== undefined) dados.ufCtps = linhaDados[col.UfCtps];
            if (col.Reservista !== undefined) dados.reservista = linhaDados[col.Reservista];
            if (col.SerieReservista !== undefined) dados.serieReservista = linhaDados[col.SerieReservista];
            if (col.RegiaoReservista !== undefined) dados.regiaoReservista = linhaDados[col.RegiaoReservista];
            if (col.Agencia !== undefined) dados.agencia = linhaDados[col.Agencia];
            if (col.Conta !== undefined) dados.conta = linhaDados[col.Conta];
            if (col.Vigencia !== undefined) {
               let vig = linhaDados[col.Vigencia];
               if (Object.prototype.toString.call(vig) === '[object Date]') {
                  dados.vigencia = Utilities.formatDate(vig, "GMT-3", "dd/MM/yyyy");
               } else {
                  dados.vigencia = vig; 
               }
            }

            break; 
        }
      }
    }
    
    if (!encontrou) {
        ui.alert("⚠️ CPF/Nome não encontrado na base de dados.");
    }
  }

  if (!dados.formacao) dados.formacao = "NÃO INFORMADO";

  let mensagem = '📋 CONFIRA DADOS:\n\n' +
                 'Nome: ' + dados.nome + '\n' +
                 'CPF: ' + dados.cpf + '\n' +
                 'Vigência (Fim): ' + (dados.vigencia || "---") + '\n' +
                 'Valor Novo: ' + (dados.valor || "--");
  
  const resposta = ui.alert('Gerar Documento?', mensagem, ui.ButtonSet.YES_NO);
  if (resposta == ui.Button.NO) return;

  if (deveAtualizarBase) atualizarBaseDeDados(planilha, dados.cpf, dados.escola, dados.novaCarga, valorParaBase);

  // geração
  const dataObj = new Date();
  const meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];
  const dia = Utilities.formatDate(dataObj, "GMT-3", "dd");
  const mesExtenso = meses[parseInt(Utilities.formatDate(dataObj, "GMT-3", "M"))-1];
  const ano = Utilities.formatDate(dataObj,"GMT-3", "yyyy");
  const dataHojeExtenso = dia + " de " + mesExtenso + " de " + ano;

  const idPlanilha = planilha.getId();
  const folder = DriveApp.getFileById(idPlanilha).getParents().hasNext() ? DriveApp.getFileById(idPlanilha).getParents().next() : DriveApp.getRootFolder();
  
  const nomeFinal = nomePadrao + " - " + dados.nome;
  const copia = DriveApp.getFileById(idDoModelo).makeCopy(nomeFinal, folder);
  const doc = DocumentApp.openById(copia.getId());
  const corpo = doc.getBody();

  const replace = (tag, valor) => {
     let v = valor || " ";
     corpo.replaceText("\\{\\{\\s*" + tag + "\\s*\\}\\}", v);
     corpo.replaceText("\\{\\{\\s*" + tag.toUpperCase() + "\\s*\\}\\}", v);
  };

  replace("Data Atual", dataHojeExtenso);
  replace("Data ExtensoResc", dataHojeExtenso);
  replace("Data do Registro", dados.dataInicio);
  replace("Nome", dados.nome);
  replace("CPF", dados.cpf); 
  replace("RG", dados.rg);
  replace("Orgao", dados.orgao);
  replace("Identidade", dados.rg); 
  replace("Órgão Expedidor", dados.orgao); 
  replace("Data de Emissão", dados.dtEmissaoRG); 
  replace("UF", dados.ufEndereco); 
  replace("Estado Civil", dados.estadoCivil);
  replace("Data de Nascimento", dados.dataNasc);
  replace("Escola", dados.escola);
  replace("Escola Destino", dados.escola);
  replace("Município", dados.municipioEsc);
  replace("Municipio", dados.municipioEsc);
  replace("Municipio Residencia", dados.municipioRes);
  replace("Municipio Residencial", dados.municipioRes);
  replace("Curso", dados.formacao);
  replace("Formacao", dados.formacao);
  replace("Nivel", dados.nivel);
  replace("Matricula", dados.matricula);
  replace("Matrícula", dados.matricula);
  replace("Nº Matrícula", dados.matricula);
  replace("Numero Funcional", dados.numFuncional);
  replace("Número Funcional", dados.numFuncional);
  replace("Numero Contrato", dados.numContrato);
  replace("Nº Contrato", dados.numContrato);
  replace("Cargo", dados.cargo);
  replace("Função", dados.cargo);
  replace("Projeto", dados.projeto);
  replace("Disciplina", dados.disciplina);
  replace("Endereco", dados.endereco);
  replace("Endereço", dados.endereco);
  replace("Numero Residencial", dados.numero);
  replace("Bairro", dados.bairro);
  replace("CEP", dados.cep);
  replace("UF Endereco", dados.ufEndereco);
  replace("UF endereco", dados.ufEndereco); 
  replace("Telefone", dados.telefone);
  replace("Naturalidade", dados.naturalidade);
  replace("Titulo Eleitoral", dados.titulo);
  replace("Titulo", dados.titulo);
  replace("Zona", dados.zona);
  replace("Secao", dados.secao);
  replace("Seção", dados.secao);
  replace("Pasep", dados.pis);
  replace("PIS", dados.pis);
  replace("Pai", dados.pai);
  replace("Mae", dados.mae);
  replace("Mãe", dados.mae);
  replace("Agencia", dados.agencia);
  replace("Conta Corrente", dados.conta);
  replace("Reservista", dados.reservista);
  replace("Serie Reservista", dados.serieReservista);
  replace("Serie", dados.serieReservista); 
  replace("Regiao", dados.regiaoReservista); 
  replace("CTPS", dados.ctps);
  replace("Carteira de Trabalho", dados.ctps);
  replace("Serie CTPS", dados.serieCtps);
  replace("UF CTPS", dados.ufCtps);
  replace("Data Inicio", dados.dataInicio);
  replace("Carga Horaria", cargaParaDoc);
  replace("Valor", dados.valor);
  replace("Valor por extenso", dados.valorExtenso);
  replace("Nova Carga Horaria", dados.novaCarga);
  replace("Admissão", dados.admissao);
  replace("Vigencia Contrato", dados.vigencia);
  replace("Vigência Contrato", dados.vigencia);
  replace("E-Mail", dados.email);
  replace("aPartirde", dados.dataInicio);

  doc.saveAndClose();

  const url = doc.getUrl();
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family: sans-serif; padding: 10px;"><p>✅ <strong>' + nomePadrao + '</strong> gerado!</p><a href="' + url + '" target="_blank" style="background-color: #2196F3; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block;">📂 Abrir Documento</a></div>'
  ).setWidth(350).setHeight(250);
  ui.showModalDialog(html, "Concluído");
}

function atualizarBaseDeDados(planilha, cpfAlvo, novaEscola, novaCarga, novoValor) {
  const abaContratos = planilha.getSheetByName("Contratos");
  if (!abaContratos) return;
  const dadosGerais = abaContratos.getDataRange().getDisplayValues();
  const cabecalho = dadosGerais[0];
  let colCPF = -1, colEscola = -1, colCarga = -1, colValor = -1;
  
  for (let c = 0; c < cabecalho.length; c++) {
    const nome = String(cabecalho[c]).toUpperCase().trim();
    if (nome.includes("CPF") && !nome.includes("VALIDAÇÃO")) colCPF = c;
    if (nome === "ESCOLA DESTINO") colEscola = c;
    if (nome === "CARGA HORÁRIA" || nome === "H/A") colCarga = c;
    if (nome === "VALOR R$") colValor = c;
  }
  if (colCPF === -1) return;
  if (colValor === -1) colValor = 95;
  
  let cpfNumericoAlvo = parseInt(cpfAlvo.replace(/[^0-9]/g, ""), 10);

  for (let i = dadosGerais.length - 1; i >= 1; i--) {
    let cpfLinha = parseInt(String(dadosGerais[i][colCPF]).replace(/[^0-9]/g, ""), 10);
    if (cpfLinha === cpfNumericoAlvo) {
      if (colEscola > -1 && novaEscola && novaEscola.trim() !== "") abaContratos.getRange(i + 1, colEscola + 1).setValue(novaEscola);
      if (colCarga > -1 && novaCarga && novaCarga.toString().trim() !== "") abaContratos.getRange(i + 1, colCarga + 1).setValue(novaCarga);
      if (colValor > -1 && novoValor) abaContratos.getRange(i + 1, colValor + 1).setValue(novoValor);
      
      break; 
    }
  }
}

function escreverPorExtenso(valor) {
  if (!valor) return "__________________";
  let v = parseFloat(valor.toString().replace("R$", "").replace(/\./g, "").replace(",", ".").trim());
  if (isNaN(v)) return valor; 
  if (v === 0) return "zero reais";

  const unidade = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"];
  const dezena = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"];
  const dezena10 = ["dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"];
  const centena = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"];

  function converterGrupo(n) {
    let u = n % 10;
    let d = Math.floor((n % 100) / 10);
    let c = Math.floor(n / 100);
    let texto = "";

    if (n === 100) return "cem";

    if (c > 0) texto += centena[c];

    if (d > 0 || u > 0) {
        if (texto) texto += " e ";
        if (d === 1) {
            texto += dezena10[u];
        } else {
            if (d > 0) texto += dezena[d];
            if (u > 0) {
                if (d > 0) texto += " e ";
                texto += unidade[u];
            }
        }
    }
    return texto;
  }

  let inteiro = Math.floor(v);
  let centavos = Math.round((v - inteiro) * 100);
  let partes = [];
  let milhoes = Math.floor(inteiro / 1000000);
  let restoMilhao = inteiro % 1000000;
  let milhares = Math.floor(restoMilhao / 1000);
  let restoMil = restoMilhao % 1000;

  if (milhoes > 0) {
      let termo = converterGrupo(milhoes);
      partes.push(termo + (milhoes > 1 ? " milhões" : " milhão"));
  }

  if (milhares > 0) {
      let conector = "";
      if (milhoes > 0 && restoMil === 0) conector = " e "; 
      partes.push(conector + converterGrupo(milhares) + " mil");
  }

  if (restoMil > 0) {
      if ((milhares > 0 || milhoes > 0) && restoMil < 100) {
         partes.push("e");
      } else if (partes.length > 0) {
      }
      partes.push(converterGrupo(restoMil));
  }

  let extenso = partes.join(" ");
  if (inteiro > 0) extenso += (inteiro > 1 ? " reais" : " real");

  if (centavos > 0) {
      if (inteiro > 0) extenso += " e ";
      extenso += converterGrupo(centavos) + (centavos > 1 ? " centavos" : " centavo");
  }

  return extenso.charAt(0).toUpperCase() + extenso.slice(1);
}
