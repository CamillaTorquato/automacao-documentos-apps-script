/*******************************************************
 * CONFIGURAÇÃO
 *******************************************************/

const PASTA_AAAAMM_ID = 'COLE_AQUI_O_ID_DA_PASTA';
const NOME_ABA = 'Lançamentos';


/*******************************************************
 * COLUNAS
 *******************************************************/

const COL_DATA = 1;       // A
const COL_CNPJ = 2;       // B
const COL_NOME = 4;       // D
const COL_VALOR = 6;      // F
const COL_FATURA = 7;     // G

const COL_NF = 10;        // J
const COL_BOL = 11;       // K
const COL_STATUS = 17;    // Q


/*******************************************************
 * MENU
 *******************************************************/

function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu('📎 Documentos')
    .addItem('Processar um dia', 'processarDia')
    .addToUi();
}


/*******************************************************
 * PROCESSAR UM DIA
 *******************************************************/

function processarDia() {

  const ui = SpreadsheetApp.getUi();

  const resposta = ui.prompt(
    'Processar documentos',
    'Digite a data da pasta.\n\nExemplo: 28.11.2025',
    ui.ButtonSet.OK_CANCEL
  );

  if (
    resposta.getSelectedButton() !== ui.Button.OK
  ) {
    return;
  }

  const dataPasta =
    resposta.getResponseText().trim();

  if (
    !/^\d{2}\.\d{2}\.\d{4}$/.test(dataPasta)
  ) {

    ui.alert(
      'Data inválida.\n\nUse DD.MM.AAAA.'
    );

    return;
  }

  executarProcessamento(dataPasta);
}


/*******************************************************
 * PROCESSAMENTO PRINCIPAL
 *******************************************************/

function executarProcessamento(dataPasta) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const aba =
    ss.getSheetByName(NOME_ABA);

  if (!aba) {

    SpreadsheetApp.getUi().alert(
      'Não encontrei a aba "' +
      NOME_ABA +
      '".'
    );

    return;
  }


  /*****************************************************
   * PASTA 202511
   *****************************************************/

  let pasta202511;

  try {

    pasta202511 =
      DriveApp.getFolderById(
        PASTA_202511_ID
      );

  } catch (erro) {

    SpreadsheetApp.getUi().alert(
      'Erro ao acessar a pasta 202511:\n\n' +
      erro.message
    );

    return;
  }


  /*****************************************************
   * PASTA DO DIA
   *****************************************************/

  const pastas =
    pasta202511.getFoldersByName(
      dataPasta
    );

  if (!pastas.hasNext()) {

    SpreadsheetApp.getUi().alert(
      'Não encontrei a pasta:\n\n' +
      dataPasta
    );

    return;
  }

  const pastaDia =
    pastas.next();


  /*****************************************************
   * CARREGA PDFs
   *****************************************************/

  const arquivos = [];

  const iterator =
    pastaDia.getFiles();

  while (iterator.hasNext()) {

    const arquivo =
      iterator.next();

    if (
      arquivo.getMimeType() !==
      MimeType.PDF
    ) {
      continue;
    }

    const nome =
      arquivo.getName();

    arquivos.push({

      id: arquivo.getId(),

      nome: nome,

      nomeBusca:
        normalizarTexto(nome),

      url:
        arquivo.getUrl(),

      tipo:
        identificarTipoDocumento(nome)
    });
  }


  if (arquivos.length === 0) {

    SpreadsheetApp.getUi().alert(
      'A pasta ' +
      dataPasta +
      ' não possui PDFs.'
    );

    return;
  }


  /*****************************************************
   * DADOS DA PLANILHA
   *****************************************************/

  const ultimaLinha =
    aba.getLastRow();

  if (ultimaLinha < 2) {

    SpreadsheetApp.getUi().alert(
      'Não existem lançamentos.'
    );

    return;
  }

  const dados =
    aba
      .getRange(
        2,
        1,
        ultimaLinha - 1,
        17
      )
      .getValues();


  /*****************************************************
   * CONTADORES
   *****************************************************/

  let linhasProcessadas = 0;
  let nfEncontradas = 0;
  let bolEncontrados = 0;
  let nfNaoEncontradas = 0;
  let bolNaoEncontrados = 0;
  let casosConferencia = 0;


  /*****************************************************
   * PROCESSA LINHAS
   *****************************************************/

  for (
    let i = 0;
    i < dados.length;
    i++
  ) {

    const linha =
      i + 2;


    const data =
      dados[i][COL_DATA - 1];

    const cnpj =
      dados[i][COL_CNPJ - 1];

    const nome =
      dados[i][COL_NOME - 1];

    const valor =
      dados[i][COL_VALOR - 1];

    const numeroFatura =
      dados[i][COL_FATURA - 1];


    /***************************************************
     * DATA
     ***************************************************/

    if (
      !dataCorresponde(
        data,
        dataPasta
      )
    ) {
      continue;
    }

    linhasProcessadas++;


    /***************************************************
     * CONTEÚDO ATUAL
     ***************************************************/

    const nfAtual =
      aba
        .getRange(
          linha,
          COL_NF
        )
        .getValue();

    const bolAtual =
      aba
        .getRange(
          linha,
          COL_BOL
        )
        .getValue();


    /***************************************************
     * BUSCA
     ***************************************************/

    const resultado =
      encontrarDocumentosSeguros(
        arquivos,
        cnpj,
        nome,
        valor,
        numeroFatura
      );


    /***************************************************
     * NF → J
     ***************************************************/

    if (
      resultado.nf.status ===
      'ENCONTRADO'
    ) {

      if (!nfAtual) {

        inserirLink(
          aba,
          linha,
          COL_NF,
          resultado.nf.arquivo.nome,
          resultado.nf.arquivo.url
        );

        nfEncontradas++;
      }

    } else if (
      resultado.nf.status ===
      'MULTIPLOS'
    ) {

      registrarStatusSeVazio(
        aba,
        linha,
        'CONFERIR NF - múltiplos documentos'
      );

      casosConferencia++;

    } else {

      nfNaoEncontradas++;

      registrarStatusSeVazio(
        aba,
        linha,
        'NF não encontrada'
      );
    }


    /***************************************************
     * BOL → K
     ***************************************************/

    if (
      resultado.bol.status ===
      'ENCONTRADO'
    ) {

      if (!bolAtual) {

        inserirLink(
          aba,
          linha,
          COL_BOL,
          resultado.bol.arquivo.nome,
          resultado.bol.arquivo.url
        );

        bolEncontrados++;
      }

    } else if (
      resultado.bol.status ===
      'MULTIPLOS'
    ) {

      registrarStatusSeVazio(
        aba,
        linha,
        'CONFERIR BOL - múltiplos documentos'
      );

      casosConferencia++;

    } else {

      bolNaoEncontrados++;

      registrarStatusSeVazio(
        aba,
        linha,
        'BOL não encontrado'
      );
    }
  }


  /*****************************************************
   * RESUMO
   *****************************************************/

  SpreadsheetApp.getUi().alert(

    'Processamento concluído!\n\n' +

    'Pasta: ' +
    dataPasta +

    '\n\nLinhas processadas: ' +
    linhasProcessadas +

    '\n\nNF colocadas em J: ' +
    nfEncontradas +

    '\nBOL colocados em K: ' +
    bolEncontrados +

    '\n\nNF não encontradas: ' +
    nfNaoEncontradas +

    '\nBOL não encontrados: ' +
    bolNaoEncontrados +

    '\n\nCasos para conferência: ' +
    casosConferencia
  );
}


/*******************************************************
 * BUSCA SEGURA
 *******************************************************/

function encontrarDocumentosSeguros(
  arquivos,
  cnpj,
  nome,
  valor,
  numeroFatura
) {

  const cnpjBusca =
    normalizarCNPJ(cnpj);

  const nomeBusca =
    normalizarTexto(nome);

  const valorBusca =
    normalizarValor(valor);

  const faturaBusca =
    normalizarNumero(numeroFatura);


  /*****************************************************
   * SEPARA NF E BOL
   *
   * IMPORTANTE:
   *
   * FAT/FATURA entra nos DOIS grupos.
   *****************************************************/

  const arquivosNF =
    arquivos.filter(
      function(arquivo) {

        return (
          arquivo.tipo === 'NF' ||
          arquivo.tipo === 'FAT'
        );
      }
    );

  const arquivosBOL =
    arquivos.filter(
      function(arquivo) {

        return (
          arquivo.tipo === 'BOL' ||
          arquivo.tipo === 'FAT'
        );
      }
    );


  return {

    nf:
      procurarDocumentoSeguro(
        arquivosNF,
        cnpjBusca,
        nomeBusca,
        valorBusca,
        faturaBusca
      ),

    bol:
      procurarDocumentoSeguro(
        arquivosBOL,
        cnpjBusca,
        nomeBusca,
        valorBusca,
        faturaBusca
      )
  };
}


/*******************************************************
 * REGRA DE CORRESPONDÊNCIA
 *******************************************************/

function procurarDocumentoSeguro(
  arquivos,
  cnpj,
  nome,
  valor,
  fatura
) {


  /*****************************************************
   * SE EXISTE FATURA
   *****************************************************/

  if (fatura) {

    const candidatos =
      arquivos.filter(
        function(arquivo) {

          return arquivo.nomeBusca
            .includes(fatura);
        }
      );


    if (
      candidatos.length === 0
    ) {

      return {
        status:
          'NAO_ENCONTRADO'
      };
    }


    if (
      candidatos.length === 1
    ) {

      return {
        status:
          'ENCONTRADO',

        arquivo:
          candidatos[0]
      };
    }


    /*
     * Desempate por nome + valor
     */

    const refinados =
      candidatos.filter(
        function(arquivo) {

          return (
            nome &&
            valor &&
            arquivo.nomeBusca.includes(nome) &&
            valorContidoNoTexto(
              arquivo.nomeBusca,
              valor
            )
          );
        }
      );


    if (
      refinados.length === 1
    ) {

      return {
        status:
          'ENCONTRADO',

        arquivo:
          refinados[0]
      };
    }


    return {
      status:
        'MULTIPLOS'
    };
  }


  /*****************************************************
   * SEM FATURA:
   *
   * NOME + VALOR SÃO OBRIGATÓRIOS
   *****************************************************/

  if (
    !nome ||
    !valor
  ) {

    return {
      status:
        'NAO_ENCONTRADO'
    };
  }


  const candidatos =
    arquivos.filter(
      function(arquivo) {

        const nomeBate =
          arquivo.nomeBusca
            .includes(nome);

        const valorBate =
          valorContidoNoTexto(
            arquivo.nomeBusca,
            valor
          );

        return (
          nomeBate &&
          valorBate
        );
      }
    );


  /*****************************************************
   * NENHUM CANDIDATO
   *****************************************************/

  if (
    candidatos.length === 0
  ) {

    return {
      status:
        'NAO_ENCONTRADO'
    };
  }


  /*****************************************************
   * UM ÚNICO CANDIDATO
   *****************************************************/

  if (
    candidatos.length === 1
  ) {

    return {
      status:
        'ENCONTRADO',

      arquivo:
        candidatos[0]
    };
  }


  /*****************************************************
   * VÁRIOS CANDIDATOS
   *
   * TENTA CNPJ
   *****************************************************/

  if (cnpj) {

    const candidatosCNPJ =
      candidatos.filter(
        function(arquivo) {

          return arquivo.nomeBusca
            .includes(cnpj);
        }
      );


    if (
      candidatosCNPJ.length === 1
    ) {

      return {
        status:
          'ENCONTRADO',

        arquivo:
          candidatosCNPJ[0]
      };
    }
  }


  /*****************************************************
   * NÃO FOI POSSÍVEL DETERMINAR
   *****************************************************/

  return {
    status:
      'MULTIPLOS'
  };
}


/*******************************************************
 * IDENTIFICA NF, BOL OU FAT
 *******************************************************/

function identificarTipoDocumento(nome) {

  const texto =
    normalizarTexto(nome);


  /*
   * FATURA / FAT
   *
   * Deve ser identificado ANTES de NF/BOL,
   * pois o mesmo arquivo será usado nas
   * duas colunas.
   */

  if (
    texto.includes('FATURA') ||
    texto.includes('FAT')
  ) {

    return 'FAT';
  }


  /*
   * BOLETO / BOL
   */

  if (
    texto.includes('BOLETO') ||
    texto.includes('BOL')
  ) {

    return 'BOL';
  }


  /*
   * NOTA FISCAL / NF
   */

  if (
    texto.includes('NOTAFISCAL') ||
    texto.includes('NF')
  ) {

    return 'NF';
  }


  return 'OUTRO';
}


/*******************************************************
 * INSERE LINK
 *******************************************************/

function inserirLink(
  aba,
  linha,
  coluna,
  nomeArquivo,
  url
) {

  const richText =
    SpreadsheetApp
      .newRichTextValue()
      .setText(
        '📄 ' + nomeArquivo
      )
      .setLinkUrl(url)
      .build();


  aba
    .getRange(
      linha,
      coluna
    )
    .setRichTextValue(
      richText
    );
}


/*******************************************************
 * STATUS — NUNCA SOBRESCREVE Q
 *******************************************************/

function registrarStatusSeVazio(
  aba,
  linha,
  mensagem
) {

  const celula =
    aba.getRange(
      linha,
      COL_STATUS
    );

  const atual =
    celula.getValue();


  if (
    atual === '' ||
    atual === null
  ) {

    celula.setValue(
      mensagem
    );
  }
}


/*******************************************************
 * NORMALIZA TEXTO
 *******************************************************/

function normalizarTexto(valor) {

  if (
    valor === null ||
    valor === undefined
  ) {

    return '';
  }


  return String(valor)
    .normalize('NFD')
    .replace(
      /[\u0300-\u036f]/g,
      ''
    )
    .toUpperCase()
    .replace(
      /[^A-Z0-9]/g,
      ''
    );
}


/*******************************************************
 * NORMALIZA CNPJ
 *******************************************************/

function normalizarCNPJ(valor) {

  if (!valor) {
    return '';
  }

  return String(valor)
    .replace(
      /\D/g,
      ''
    );
}


/*******************************************************
 * NORMALIZA FATURA
 *******************************************************/

function normalizarNumero(valor) {

  if (
    valor === null ||
    valor === undefined ||
    valor === ''
  ) {

    return '';
  }

  return String(valor)
    .replace(
      /\D/g,
      ''
    );
}


/*******************************************************
 * NORMALIZA VALOR — CORRIGIDA
 *******************************************************/

function normalizarValor(valor) {

  if (
    valor === null ||
    valor === undefined ||
    valor === ''
  ) {

    return '';
  }


  /*
   * Se o Sheets armazenar como número:
   *
   * -31454.5
   *
   * representa:
   *
   * (31.454,50)
   */

  if (
    typeof valor === 'number'
  ) {

    if (
      isNaN(valor)
    ) {

      return '';
    }


    return Math.abs(valor)
      .toFixed(2)
      .replace(
        '.',
        ''
      );
  }


  /*
   * Caso seja texto.
   */

  let texto =
    String(valor)
      .trim();


  texto =
    texto
      .replace(
        /R\$/gi,
        ''
      )
      .replace(
        /\s/g,
        ''
      )
      .replace(
        /\(/g,
        ''
      )
      .replace(
        /\)/g,
        ''
      )
      .replace(
        /-/g,
        ''
      );


  /*
   * Formato brasileiro:
   *
   * 31.454,50
   *
   * →
   *
   * 31454.50
   */

  texto =
    texto.replace(
      /\./g,
      ''
    );


  texto =
    texto.replace(
      ',',
      '.'
    );


  const numero =
    parseFloat(texto);


  if (
    isNaN(numero)
  ) {

    return '';
  }


  return Math.abs(numero)
    .toFixed(2)
    .replace(
      '.',
      ''
    );
}


/*******************************************************
 * COMPARA VALOR
 *******************************************************/

function valorContidoNoTexto(
  texto,
  valor
) {

  if (
    !texto ||
    !valor
  ) {

    return false;
  }


  const textoNormalizado =
    normalizarTexto(texto);


  if (
    textoNormalizado.includes(valor)
  ) {

    return true;
  }


  /*
   * Também verifica outras formas.
   */

  if (
    valor.length >= 3
  ) {

    const parteInteira =
      valor.substring(
        0,
        valor.length - 2
      );

    const centavos =
      valor.substring(
        valor.length - 2
      );


    const formas = [

      parteInteira +
      ',' +
      centavos,

      parteInteira +
      '.' +
      centavos,

      inserirSeparadorMilhar(
        parteInteira
      ) +
      ',' +
      centavos,

      inserirSeparadorMilhar(
        parteInteira
      ) +
      '.' +
      centavos
    ];


    for (
      const forma of formas
    ) {

      if (
        textoNormalizado.includes(
          normalizarTexto(forma)
        )
      ) {

        return true;
      }
    }
  }


  return false;
}


/*******************************************************
 * SEPARADOR DE MILHAR
 *******************************************************/

function inserirSeparadorMilhar(
  numero
) {

  return numero.replace(
    /\B(?=(\d{3})+(?!\d))/g,
    '.'
  );
}


/*******************************************************
 * COMPARA DATA
 *******************************************************/

function dataCorresponde(
  dataCelula,
  dataPasta
) {

  if (!dataCelula) {
    return false;
  }


  /*
   * Data real do Sheets
   */

  if (
    Object.prototype.toString
      .call(dataCelula) ===
    '[object Date]'
  ) {

    const dia =
      String(
        dataCelula.getDate()
      ).padStart(
        2,
        '0'
      );

    const mes =
      String(
        dataCelula.getMonth() + 1
      ).padStart(
        2,
        '0'
      );

    const ano =
      dataCelula.getFullYear();


    return (
      dia +
      '.' +
      mes +
      '.' +
      ano ===
      dataPasta
    );
  }


  /*
   * Data como texto
   */

  const texto =
    String(dataCelula)
      .trim()
      .replace(
        /\//g,
        '.'
      )
      .replace(
        /-/g,
        '.'
      );


  return (
    texto ===
    dataPasta
  );
}
