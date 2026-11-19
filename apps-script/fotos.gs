/**
 * Recebe fotos enviadas pela página /fotos do convite e salva numa pasta do Google Drive.
 *
 * COMO CONFIGURAR
 * 1. Crie uma pasta no Google Drive (ex.: "Fotos - 15 anos Milena").
 * 2. Abra a pasta e copie o código que aparece no link depois de /folders/
 * 3. Cole esse código em PASTA_ID abaixo.
 * 4. Implantar > Nova implantação > Tipo: App da Web
 *      Executar como: Eu
 *      Quem pode acessar: Qualquer pessoa
 * 5. Copie a URL gerada (termina em /exec) e cole em URL_DO_SCRIPT no fotos/index.html.
 */

const PASTA_ID = 'COLE_AQUI_O_ID_DA_PASTA';

// Proteção simples contra envios indevidos. Use o MESMO texto no fotos/index.html.
const CHAVE_EVENTO = 'Milena15';

const TAMANHO_MAXIMO_MB = 15;
const TIPOS_ACEITOS = ['image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif', 'image/gif'];

function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);

    if (dados.chave !== CHAVE_EVENTO) {
      return resposta({ ok: false, erro: 'chave inválida' });
    }
    if (TIPOS_ACEITOS.indexOf(dados.tipo) === -1) {
      return resposta({ ok: false, erro: 'tipo de arquivo não aceito' });
    }

    const bytes = Utilities.base64Decode(dados.base64);
    if (bytes.length > TAMANHO_MAXIMO_MB * 1024 * 1024) {
      return resposta({ ok: false, erro: 'arquivo muito grande' });
    }

    const nome = montarNome(dados.convidado, dados.nomeArquivo, dados.tipo);
    const arquivo = DriveApp.getFolderById(PASTA_ID)
      .createFile(Utilities.newBlob(bytes, dados.tipo, nome));

    return resposta({ ok: true, id: arquivo.getId() });
  } catch (erro) {
    return resposta({ ok: false, erro: String(erro) });
  }
}

// Permite abrir a URL no navegador só para testar se está no ar.
function doGet() {
  return resposta({ ok: true, mensagem: 'Script de fotos funcionando.' });
}

function montarNome(convidado, nomeOriginal, tipo) {
  const data = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd_HH-mm-ss');
  const quem = limpar(convidado) || 'convidado';
  const extensao = (tipo.split('/')[1] || 'jpg').replace('jpeg', 'jpg');
  const sorteio = Math.random().toString(36).slice(2, 6);
  return data + '_' + quem + '_' + sorteio + '.' + extensao;
}

function limpar(texto) {
  return String(texto || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-zA-Z0-9 ]/g, '')
    .trim().replace(/\s+/g, '-')
    .slice(0, 40);
}

function resposta(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}
