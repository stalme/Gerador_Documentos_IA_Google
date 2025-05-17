// --- CONFIGURAÇÕES DA PLANILHA E SCRIPT ---
const NOME_PLANILHA_PRINCIPAL = "Certificados";
const LINHA_INICIAL_DADOS = 2;
const NOME_INSTITUICAO = "[Nome da Sua Escola/Instituição Aqui]"; // Substitua pelo nome da sua instituição

// --- MAPEAMENTO DE COLUNAS (AJUSTE CONFORME SUA PLANILHA!) ---
const COLUNA_NOME_PARTICIPANTE = 1;
const COLUNA_NOME_EVENTO_ATIVIDADE = 2;
const COLUNA_MODELO_DOCUMENTO = 3;
const COLUNA_DATA_EVENTO = 4;
const COLUNA_CARGA_HORARIA = 5;
const COLUNA_DETALHES_ADICIONAIS = 6;
const COLUNA_LOCAL_ESPECIFICO = 7;
const COLUNA_HORA_INICIO_COMPARECIMENTO = 8;
const COLUNA_HORA_FIM_COMPARECIMENTO = 9;
const COLUNA_DISCIPLINA_PROJETO_VOLUNTARIO = 10;
const COLUNA_PERIODO_VOLUNTARIADO = 11;
const COLUNA_TITULO_PALESTRA_MINICURSO = 12;
const COLUNA_CPF_PARTICIPANTE = 13;
const COLUNA_SALVAR_COMO_MODELO_REUTILIZAVEL = 14; // Ex: Coluna N [Checkbox] - Se marcada, o texto da IA é salvo como modelo com placeholder de nome
const COLUNA_TEXTO_BASE_DOCUMENTO = 15;  // Ex: Coluna O
const COLUNA_USAR_TEXTO_BASE = 16;       // Ex: Coluna P [Checkbox]

// Colunas de Saída
const COLUNA_TEXTO_CERTIFICADO_GERADO = 17; // Ex: Coluna Q
const COLUNA_STATUS_GERACAO = 18;        // Ex: Coluna R
const COLUNA_CODIGO_VERIFICACAO = 19;     // Ex: Coluna S

/**
 * Adiciona um menu personalizado à planilha ao abrir.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Gerador de Documentos IA')
    .addItem('Gerar Texto para Linha(s) Selecionada(s)', 'gerarTextoParaLinhasSelecionadas')
    .addItem('Gerar Texto para Todas as Linhas Pendentes', 'gerarTextoParaTodasAsLinhasPendentes')
    .addToUi();
}

/**
 * Obtém a chave da API Gemini das Propriedades do Script.
 */
function getApiKey() {
  const userProperties = PropertiesService.getScriptProperties();
  const apiKey = userProperties.getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    SpreadsheetApp.getUi().alert('Chave de API não configurada!', 'Por favor, configure a GEMINI_API_KEY nas Propriedades do Script (Arquivo > Propriedades do projeto > Propriedades do script).', SpreadsheetApp.getUi().ButtonSet.OK);
    throw new Error('Chave de API GEMINI_API_KEY não configurada.');
  }
  return apiKey;
}

/**
 * Alternativa para ui.confirm usando ui.prompt.
 */
function solicitarConfirmacaoAlternativa(titulo, mensagemPrompt) {
  const ui = SpreadsheetApp.getUi();
  const respostaPrompt = ui.prompt(titulo, mensagemPrompt + "\n\nDigite 'SIM' para confirmar.", ui.ButtonSet.OK_CANCEL);
  if (respostaPrompt.getSelectedButton() == ui.Button.OK && respostaPrompt.getResponseText().trim().toUpperCase() === 'SIM') {
    return ui.Button.YES;
  } else {
    return ui.Button.NO;
  }
}

/**
 * Formata a data de entrada para ser usada no prompt.
 */
function formatarDataParaPrompt(dataInput) {
  if (!dataInput) {
    return "[Data não informada]";
  }
  if (dataInput instanceof Date && !isNaN(dataInput.getTime())) {
    return dataInput.toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' });
  }
  if (typeof dataInput === 'string') {
    dataInput = dataInput.trim();
    const parts = dataInput.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (parts) {
      const day = parseInt(parts[1], 10);
      const month = parseInt(parts[2], 10) - 1;
      const year = parseInt(parts[3], 10);
      const dateObj = new Date(year, month, day);
      if (dateObj.getFullYear() === year && dateObj.getMonth() === month && dateObj.getDate() === day) {
        return dateObj.toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' });
      } else {
        return dataInput;
      }
    } else {
      return dataInput;
    }
  }
  if (typeof dataInput === 'number') {
    const dateObjFromNum = new Date(dataInput);
    if (!isNaN(dateObjFromNum.getTime())) {
        return dateObjFromNum.toLocaleDateString('pt-BR', { timeZone: 'America/Sao_Paulo' });
    }
  }
  return String(dataInput);
}

/**
 * Constrói o prompt para a API Gemini.
 */
function construirPrompt(dadosLinha) {
  const modeloDocumento = dadosLinha[COLUNA_MODELO_DOCUMENTO - 1] || "Não especificado";
  const nomeParticipante = dadosLinha[COLUNA_NOME_PARTICIPANTE - 1] || "[Nome não informado]";
  const nomeEventoAtividade = dadosLinha[COLUNA_NOME_EVENTO_ATIVIDADE - 1] || "[Atividade não informada]";
  const dataEventoFinalParaPrompt = formatarDataParaPrompt(dadosLinha[COLUNA_DATA_EVENTO - 1]);
  const cargaHoraria = dadosLinha[COLUNA_CARGA_HORARIA - 1];
  const detalhesAdicionais = dadosLinha[COLUNA_DETALHES_ADICIONAIS - 1];

  let prompt = `Você é um assistente especializado em redigir documentos formais para instituições de ensino no Brasil, em português brasileiro.
Seu objetivo é gerar apenas o CORPO do texto para o documento solicitado, sem cabeçalhos, rodapés ou saudações finais (como "Atenciosamente,").
O texto deve ser direto e profissional.

Informações base para o documento:
- Nome do Indivíduo Principal: ${nomeParticipante}
- Modelo do Documento Solicitado: "${modeloDocumento}"
- Nome da Atividade Principal/Evento: ${nomeEventoAtividade}
- Data da Atividade Principal/Evento: ${dataEventoFinalParaPrompt}
- Instituição Emissora: ${NOME_INSTITUICAO}`;

  if (cargaHoraria) {
    prompt += `\n- Carga Horária (se aplicável ao modelo): ${cargaHoraria} horas`;
  }
  if (detalhesAdicionais) {
    prompt += `\n- Detalhes Adicionais Fornecidos: ${detalhesAdicionais}`;
  }
  const cpfParticipante = dadosLinha[COLUNA_CPF_PARTICIPANTE - 1];
  if (cpfParticipante) prompt += `\n- CPF do Indivíduo (se necessário no documento): ${cpfParticipante}`;

  prompt += "\n\n--- INSTRUÇÕES ESPECÍFICAS PARA O MODELO DE DOCUMENTO ---\n";

  switch (modeloDocumento) {
    case "Declaração de Comparecimento":
      const localComparecimento = dadosLinha[COLUNA_LOCAL_ESPECIFICO - 1] || nomeEventoAtividade;
      const horaInicio = dadosLinha[COLUNA_HORA_INICIO_COMPARECIMENTO - 1] || "[horário de início não informado]";
      const horaFim = dadosLinha[COLUNA_HORA_FIM_COMPARECIMENTO - 1] || "[horário de término não informado]";
      prompt += `Tipo: Declaração de Comparecimento.
Declare formalmente que ${nomeParticipante}${cpfParticipante ? `, portador(a) do CPF nº ${cpfParticipante},` : ''} esteve presente em ${localComparecimento} no dia ${dataEventoFinalParaPrompt}, no período das ${horaInicio} às ${horaFim}.
Se houver "Detalhes Adicionais" sobre o motivo do comparecimento (ex: "para participar da reunião X", "para realizar a prova Y"), inclua essa informação de forma concisa.
Finalize com uma frase como "Por ser verdade, firmamos a presente declaração." ou similar.`;
      break;
    case "Declaração Professor Voluntário":
      const disciplinaVoluntariado = dadosLinha[COLUNA_DISCIPLINA_PROJETO_VOLUNTARIO - 1] || "[disciplina/projeto não especificado]";
      const periodoVoluntariado = dadosLinha[COLUNA_PERIODO_VOLUNTARIADO - 1] || dataEventoFinalParaPrompt;
      prompt += `Tipo: Declaração de Atuação como Professor Voluntário.
Declare formalmente que ${nomeParticipante}${cpfParticipante ? `, portador(a) do CPF nº ${cpfParticipante},` : ''} atuou como professor(a) voluntário(a) na disciplina/projeto "${disciplinaVoluntariado}" nesta instituição (${NOME_INSTITUICAO}), durante o período de ${periodoVoluntariado}.
Mencione a carga horária total, se fornecida (${cargaHoraria ? cargaHoraria + ' horas' : 'carga horária não especificada'}).
Destaque a dedicação, o compromisso e a valiosa contribuição do(a) voluntário(a) para a comunidade acadêmica.
Finalize com uma frase como "Agradecemos imensamente pela colaboração." ou similar.`;
      break;
    case "Certificado Ouvinte Palestra":
      const tituloPalestra = dadosLinha[COLUNA_TITULO_PALESTRA_MINICURSO - 1] || nomeEventoAtividade;
      prompt += `Tipo: Certificado de Participação como Ouvinte em Palestra.
Certificamos que ${nomeParticipante}${cpfParticipante ? `, portador(a) do CPF nº ${cpfParticipante},` : ''} participou como ouvinte da palestra (ser for oficina, retirar o termo ouvinte) intitulada "${tituloPalestra}", realizada em ${dataEventoFinalParaPrompt}, com carga horária de ${cargaHoraria || 'X'} horas.
O evento foi promovido por ${NOME_INSTITUICAO}.
Destaque que a participação contribuiu para o seu desenvolvimento/aprimoramento.`;
      break;
    case "Certificado Ministrante Palestra":
      const tituloEspecificoPalestra = dadosLinha[COLUNA_TITULO_PALESTRA_MINICURSO - 1];
      const nomePrincipalDaAtividade = nomeEventoAtividade;
      let descricaoDaMinistracao;
      if (tituloEspecificoPalestra && tituloEspecificoPalestra.trim() !== "") {
        descricaoDaMinistracao = `ministrou a palestra intitulada "${tituloEspecificoPalestra}", realizada em ${dataEventoFinalParaPrompt}, como parte do evento "${nomePrincipalDaAtividade}"`;
      } else {
        descricaoDaMinistracao = `ministrou a palestra (ou atividade principal) "${nomePrincipalDaAtividade}", realizada em ${dataEventoFinalParaPrompt}`;
      }
      prompt += `Tipo: Certificado de Ministrante de Palestra.
Certificamos que ${nomeParticipante}${cpfParticipante ? `, portador(a) do CPF nº ${cpfParticipante},` : ''} ${descricaoDaMinistracao}, com carga horária de ${cargaHoraria || 'X'} horas.
A atividade foi promovida por ${NOME_INSTITUICAO}.
Agradeça pela valiosa contribuição e compartilhamento de conhecimento.`;
      break;
    default:
      prompt += `Modelo de documento "${modeloDocumento}" não reconhecido ou não possui instruções específicas detalhadas.
Por favor, gere um texto formal e apropriado para ${nomeParticipante} relacionado à atividade ${nomeEventoAtividade} em ${dataEventoFinalParaPrompt}, considerando os detalhes adicionais fornecidos.
Se for um certificado, use a terceira pessoa ("Certificamos que..."). Se for uma declaração, use a primeira pessoa do plural ("Declaramos que...").`;
      break;
  }
  prompt += "\n\n--- FIM DAS INSTRUÇÕES ESPECÍFICAS ---\nTexto do Documento (apenas o corpo):";
  return prompt;
}

/**
 * Chama a API Gemini
 */
function chamarApiGemini(prompt) {
  const apiKey = getApiKey();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${apiKey}`;
  const requestBody = {
    "contents": [{"parts": [{"text": prompt}]}],
    "generationConfig": {
      "temperature": 0.6, "topK": 1, "topP": 1, "maxOutputTokens": 800, "stopSequences": []
    },
    "safetySettings": [
      { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" }
    ]
  };
  const options = {'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(requestBody), 'muteHttpExceptions': true};

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      if (jsonResponse.candidates && jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts) {
        return jsonResponse.candidates[0].content.parts[0].text.trim();
      } else if (jsonResponse.promptFeedback && jsonResponse.promptFeedback.blockReason) {
        Logger.log(`Geração bloqueada pela API. Motivo: ${jsonResponse.promptFeedback.blockReason}.`);
        SpreadsheetApp.getUi().alert(`Geração de texto bloqueada. Motivo: ${jsonResponse.promptFeedback.blockReason}.`);
        return null;
      } else { Logger.log("Resposta da API não contém o texto esperado: " + responseBody); return null; }
    } else {
      Logger.log(`Erro API Gemini: ${responseCode} - ${responseBody}`);
      SpreadsheetApp.getUi().alert(`Erro API Gemini: ${responseCode}. Ver logs.`);
      return null;
    }
  } catch (error) {
    Logger.log(`Exceção API Gemini: ${error.toString()}`);
    SpreadsheetApp.getUi().alert(`Exceção API: ${error.message}. Ver logs.`);
    return null;
  }
}

/**
 * Substitui placeholders em um texto base com dados da linha.
 */
function substituirPlaceholders(textoBase, dadosLinha) {
  let textoProcessado = textoBase;

  const placeholders = {
    "{{NOME_PARTICIPANTE}}": dadosLinha[COLUNA_NOME_PARTICIPANTE - 1],
    "{{NOME_EVENTO_ATIVIDADE}}": dadosLinha[COLUNA_NOME_EVENTO_ATIVIDADE - 1],
    "{{MODELO_DOCUMENTO}}": dadosLinha[COLUNA_MODELO_DOCUMENTO - 1],
    "{{DATA_EVENTO}}": formatarDataParaPrompt(dadosLinha[COLUNA_DATA_EVENTO - 1]),
    "{{CARGA_HORARIA}}": dadosLinha[COLUNA_CARGA_HORARIA - 1] || "",
    "{{DETALHES_ADICIONAIS}}": dadosLinha[COLUNA_DETALHES_ADICIONAIS - 1] || "",
    "{{LOCAL_ESPECIFICO}}": dadosLinha[COLUNA_LOCAL_ESPECIFICO - 1] || "",
    "{{HORA_INICIO_COMPARECIMENTO}}": dadosLinha[COLUNA_HORA_INICIO_COMPARECIMENTO - 1] || "",
    "{{HORA_FIM_COMPARECIMENTO}}": dadosLinha[COLUNA_HORA_FIM_COMPARECIMENTO - 1] || "",
    "{{DISCIPLINA_PROJETO_VOLUNTARIO}}": dadosLinha[COLUNA_DISCIPLINA_PROJETO_VOLUNTARIO - 1] || "",
    "{{PERIODO_VOLUNTARIADO}}": dadosLinha[COLUNA_PERIODO_VOLUNTARIADO - 1] || "",
    "{{TITULO_PALESTRA_MINICURSO}}": dadosLinha[COLUNA_TITULO_PALESTRA_MINICURSO - 1] || "",
    "{{CPF_PARTICIPANTE}}": dadosLinha[COLUNA_CPF_PARTICIPANTE - 1] || "",
    "{{NOME_INSTITUICAO}}": NOME_INSTITUICAO
  };

  for (const placeholder in placeholders) {
    const valor = placeholders[placeholder] !== null && placeholders[placeholder] !== undefined ? placeholders[placeholder] : "";
    const regex = new RegExp(placeholder.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
    textoProcessado = textoProcessado.replace(regex, valor);
  }
  return textoProcessado;
}

/**
 * Processa uma única linha da planilha
 */
function processarLinha(numeroLinha, sheet) {
  const rangeLinha = sheet.getRange(numeroLinha, 1, 1, sheet.getMaxColumns());
  const dadosLinha = rangeLinha.getValues()[0];
  if (!dadosLinha[COLUNA_NOME_PARTICIPANTE - 1]) return;

  const salvarComoModeloReutilizavelCheckbox = dadosLinha[COLUNA_SALVAR_COMO_MODELO_REUTILIZAVEL - 1] === true; // Usa a constante RENOMEADA
  const usarTextoBaseCheckbox = dadosLinha[COLUNA_USAR_TEXTO_BASE - 1] === true;
  let textoBaseArmazenado = dadosLinha[COLUNA_TEXTO_BASE_DOCUMENTO - 1];
  let textoFinalParaCertificado = "";
  let statusFinal = "";
  let codigoVerificacao = Utilities.getUuid();

  if (usarTextoBaseCheckbox && textoBaseArmazenado && typeof textoBaseArmazenado === 'string' && textoBaseArmazenado.trim() !== "") {
    Logger.log(`Linha ${numeroLinha}: Usando texto base armazenado.`);
    textoFinalParaCertificado = substituirPlaceholders(textoBaseArmazenado, dadosLinha);
    statusFinal = "Texto Gerado (Base)";
  } else {
    Logger.log(`Linha ${numeroLinha}: Gerando novo texto com IA.`);
    const promptParaIA = construirPrompt(dadosLinha);
    if (!promptParaIA) {
      statusFinal = "Erro: Falha ao construir prompt";
      sheet.getRange(numeroLinha, COLUNA_STATUS_GERACAO).setValue(statusFinal);
      sheet.getRange(numeroLinha, COLUNA_CODIGO_VERIFICACAO).setValue("");
      Logger.log(`Linha ${numeroLinha}: ${statusFinal}`);
      return;
    }

    const textoGeradoPelaIA = chamarApiGemini(promptParaIA);

    if (textoGeradoPelaIA) {
      textoFinalParaCertificado = textoGeradoPelaIA; // Este é o texto com o nome do participante atual
      statusFinal = "Texto Gerado (IA)";

      // Se "Salvar Como Modelo Reutilizável?" estiver marcado
      if (salvarComoModeloReutilizavelCheckbox) {
        const nomeDoParticipanteAtual = dadosLinha[COLUNA_NOME_PARTICIPANTE - 1];
        let textoModeloComPlaceholder = textoGeradoPelaIA; // Começa com o texto gerado pela IA

        if (nomeDoParticipanteAtual && nomeDoParticipanteAtual.trim() !== "") {
          // Escapa caracteres especiais no nome para usar em RegExp de forma segura
          const nomeEscapadoRegex = nomeDoParticipanteAtual.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
          const regexNome = new RegExp(nomeEscapadoRegex, 'g'); // 'g' para substituir todas as ocorrências
          // Substitui o nome do participante atual pelo placeholder no texto que será salvo como modelo
          textoModeloComPlaceholder = textoGeradoPelaIA.replace(regexNome, '{{NOME_PARTICIPANTE}}');
        }
        // Salva o texto JÁ COM O PLACEHOLDER na coluna de texto base
        sheet.getRange(numeroLinha, COLUNA_TEXTO_BASE_DOCUMENTO).setValue(textoModeloComPlaceholder);
        Logger.log(`Linha ${numeroLinha}: Novo modelo reutilizável (com placeholder de nome) salvo no Texto Base.`);
      }
    } else {
      statusFinal = "Erro: Falha na geração pela IA";
      codigoVerificacao = "";
    }
  }
  sheet.getRange(numeroLinha, COLUNA_TEXTO_CERTIFICADO_GERADO).setValue(textoFinalParaCertificado);
  sheet.getRange(numeroLinha, COLUNA_STATUS_GERACAO).setValue(statusFinal);
  sheet.getRange(numeroLinha, COLUNA_CODIGO_VERIFICACAO).setValue(codigoVerificacao);
  Logger.log(`Linha ${numeroLinha}: Status final - ${statusFinal}. Código: ${codigoVerificacao}. Texto (início): ${textoFinalParaCertificado.substring(0,100)}...`);
}

/**
 * Gera o texto para a(s) linha(s) atualmente selecionada(s) pelo usuário.
 */
function gerarTextoParaLinhasSelecionadas() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_PLANILHA_PRINCIPAL);
  if (!sheet) { ui.alert("Planilha '" + NOME_PLANILHA_PRINCIPAL + "' não encontrada!"); return; }
  const rangeSelecionado = sheet.getActiveRange();
  if (!rangeSelecionado) { ui.alert("Por favor, selecione pelo menos uma linha para processar."); return; }

  const linhaInicial = rangeSelecionado.getRow();
  const numLinhas = rangeSelecionado.getNumRows();
  if (linhaInicial < LINHA_INICIAL_DADOS) { ui.alert("Selecione linhas a partir da linha " + LINHA_INICIAL_DADOS + "."); return; }

  const mensagemDeConfirmacao = `Gerar texto e código de verificação para ${numLinhas} linha(s) selecionada(s) (da ${linhaInicial} até ${linhaInicial + numLinhas - 1})?`;
  const confirmacaoDoUsuario = solicitarConfirmacaoAlternativa('Confirmação', mensagemDeConfirmacao);

  if (confirmacaoDoUsuario === ui.Button.YES) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Processando linhas selecionadas...', 'Aguarde', -1);
    for (let i = 0; i < numLinhas; i++) {
      const linhaAtual = linhaInicial + i;
      try { processarLinha(linhaAtual, sheet); } catch (e) {
        Logger.log(`Erro ao processar linha ${linhaAtual}: ${e.toString()}`);
        sheet.getRange(linhaAtual, COLUNA_STATUS_GERACAO).setValue("Erro no Script");
      }
      SpreadsheetApp.flush();
    }
    SpreadsheetApp.getActiveSpreadsheet().toast('Processamento concluído!', 'Sucesso', 5);
    ui.alert("Processamento concluído!");
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('Operação cancelada.', 'Cancelado', 5);
  }
}

/**
 * Gera texto para todas as linhas pendentes.
 */
function gerarTextoParaTodasAsLinhasPendentes() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOME_PLANILHA_PRINCIPAL);
  if (!sheet) { ui.alert("Planilha '" + NOME_PLANILHA_PRINCIPAL + "' não encontrada!"); return; }

  const ultimaLinhaComConteudo = sheet.getLastRow();
  if (ultimaLinhaComConteudo < LINHA_INICIAL_DADOS) { ui.alert("Não há dados para processar."); return; }

  const mensagemDeConfirmacao = `Gerar texto e código de verificação para todas as linhas pendentes (da linha ${LINHA_INICIAL_DADOS} até ${ultimaLinhaComConteudo})?`;
  const confirmacaoDoUsuario = solicitarConfirmacaoAlternativa('Confirmação', mensagemDeConfirmacao);

  if (confirmacaoDoUsuario === ui.Button.YES) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Processando todas as linhas pendentes...', 'Aguarde', -1);
    let linhasProcessadasCount = 0;
    for (let i = LINHA_INICIAL_DADOS; i <= ultimaLinhaComConteudo; i++) {
      const nomeParticipante = sheet.getRange(i, COLUNA_NOME_PARTICIPANTE).getValue();
      const textoJaGerado = sheet.getRange(i, COLUNA_TEXTO_CERTIFICADO_GERADO).getValue();
      const statusAtual = sheet.getRange(i, COLUNA_STATUS_GERACAO).getValue();
      if (nomeParticipante && (textoJaGerado === "" || (statusAtual && statusAtual.toString().toLowerCase().includes("erro")))) {
        try {
          processarLinha(i, sheet);
          linhasProcessadasCount++;
        } catch (e) {
          Logger.log(`Erro ao processar linha ${i} em lote: ${e.toString()}`);
          sheet.getRange(i, COLUNA_STATUS_GERACAO).setValue("Erro Crítico no Script");
        }
        SpreadsheetApp.flush();
        if (linhasProcessadasCount > 0 && linhasProcessadasCount % 10 === 0 && (!statusAtual || !statusAtual.toString().includes("Base"))) {
            Utilities.sleep(1000);
        }
      }
    }
    SpreadsheetApp.getActiveSpreadsheet().toast('Processamento concluído!', 'Sucesso', 5);
    if (linhasProcessadasCount > 0) { ui.alert(`Processamento de ${linhasProcessadasCount} linha(s) pendente(s) concluído!`); }
    else { ui.alert("Nenhuma linha pendente encontrada."); }
  } else {
     SpreadsheetApp.getActiveSpreadsheet().toast('Operação cancelada.', 'Cancelado', 5);
  }
}
