// ============================================
// popup.js - Com verificação de XLSX
// ============================================

document.addEventListener('DOMContentLoaded', function() {
  
  const btnExportar = document.getElementById('btnExportar');
  const status = document.getElementById('status');
  const contador = document.getElementById('contador');
  
  const DOMINIO = 'adekz.jawplasticos.com.br';

  // Defina aqui as colunas que você quer no Excel (vazio = pega todas)
  // Ex: ['código', 'descrição', 'quantidade', 'saldo']
  const COLUNAS_DESEJADAS = [];
  
  // VERIFICA se XLSX está carregado
  if (typeof XLSX === 'undefined') {
    status.textContent = '❌ Erro: Biblioteca XLSX não carregada. Recarregue a extensão.';
    btnExportar.disabled = true;
    return;
  }
  
  async function atualizarContador() {
    const abas = await chrome.tabs.query({ currentWindow: true });
    const abasSistema = abas.filter(aba => aba.url && aba.url.includes(DOMINIO));
    contador.textContent = `${abasSistema.length} aba(s) do Adekz aberta(s)`;
  }
  
  atualizarContador();
  
  // Gera um nome válido pra aba do Excel, só adicionando sufixo numérico
  // se o nome já tiver sido usado antes (evita duplicidade real, sem
  // poluir nomes que já são únicos).
  function nomeValidoExcel(titulo, nomesUsados) {
    let nome = titulo || 'Aba';
    nome = nome.replace(/[\\/*?:[\]]/g, '').trim();
    if (nome.length > 28) nome = nome.substring(0, 28);

    let nomeFinal = nome;
    let contador = 1;
    while (nomesUsados.has(nomeFinal.toLowerCase())) {
      contador++;
      const sufixo = `_${contador}`;
      nomeFinal = (nome.substring(0, 31 - sufixo.length) + sufixo);
    }
    nomesUsados.add(nomeFinal.toLowerCase());
    return nomeFinal;
  }
  
  btnExportar.addEventListener('click', async function() {
    
    // Verifica novamente antes de usar
    if (typeof XLSX === 'undefined') {
      status.textContent = '❌ Erro: XLSX não disponível.';
      return;
    }
    
    btnExportar.disabled = true;
    status.textContent = '🔍 Buscando abas do Adekz...';
    
    try {
      const todasAbas = await chrome.tabs.query({ currentWindow: true });
      const abasDoSistema = todasAbas.filter(aba => 
        aba.url && aba.url.includes(DOMINIO)
      );
      
      if (abasDoSistema.length === 0) {
        status.textContent = '❌ Nenhuma aba do Adekz encontrada.';
        btnExportar.disabled = false;
        return;
      }
      
      status.textContent = `📑 ${abasDoSistema.length} aba(s) encontrada(s)`;
      
      const workbook = XLSX.utils.book_new();
      let abasExportadas = 0;
      const nomesUsados = new Set();
      
      for (let i = 0; i < abasDoSistema.length; i++) {
        const aba = abasDoSistema[i];
        
        status.textContent = `⏳ ${i + 1}/${abasDoSistema.length}: ${aba.title}`;
        
        try {
          const resultado = await chrome.scripting.executeScript({
            target: { tabId: aba.id },
            func: extrairTabelaDaPagina,
            args: [COLUNAS_DESEJADAS]
          });
          
          const { dados, nomeAba } = resultado[0]?.result || {};
          
          if (dados && dados.length > 0) {
            const worksheet = XLSX.utils.aoa_to_sheet(dados);
            const nome = nomeValidoExcel(nomeAba || aba.title, nomesUsados);
            XLSX.utils.book_append_sheet(workbook, worksheet, nome);
            abasExportadas++;
          }
          
        } catch (erro) {
          console.error(`Erro em ${aba.title}:`, erro);
        }
      }
      
      if (abasExportadas > 0) {
        const excelBinario = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBinario], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        
        await chrome.downloads.download({
          url: url,
          filename: 'Exportação_Sistema.xlsx',
          saveAs: true
        });
        
        URL.revokeObjectURL(url);
        status.textContent = `✅ ${abasExportadas} aba(s) exportadas!`;
        
      } else {
        status.textContent = '❌ Nenhuma aba com dados de tabela encontrada.';
      }
      
    } catch (erroGeral) {
      status.textContent = `❌ Erro: ${erroGeral.message}`;
    }
    
    btnExportar.disabled = false;
    atualizarContador();
  });
});

// Esta função roda DENTRO da página do Adekz (via chrome.scripting.executeScript),
// não tem acesso ao escopo acima.
function extrairTabelaDaPagina(colunasDesejadas) {
  // Pega o valor "de verdade" da célula, seja texto ou input/select/textarea
  function valorCelula(celula) {
    const campo = celula.querySelector('input, select, textarea');
    if (campo) {
      if (campo.tagName === 'SELECT') {
        return campo.options[campo.selectedIndex]?.text.trim() || '';
      }
      return (campo.value || '').trim();
    }
    return celula.innerText.trim();
  }

  const tabelas = document.querySelectorAll('table');
  if (tabelas.length === 0) return { dados: [], nomeAba: document.title };

  let maior = tabelas[0];
  let max = 0;
  tabelas.forEach(t => {
    const n = t.querySelectorAll('tr').length;
    if (n > max) { max = n; maior = t; }
  });

  const linhas = Array.from(maior.querySelectorAll('tr'));
  if (linhas.length === 0) return { dados: [], nomeAba: document.title };

  // Cabeçalho: primeira linha com <th>, ou primeira linha mesmo
  const cabecalhoTr = linhas.find(tr => tr.querySelector('th')) || linhas[0];
  const cabecalho = Array.from(cabecalhoTr.querySelectorAll('th, td')).map(c => c.innerText.trim());

  // Se colunasDesejadas foi passado, acha os índices correspondentes
  let indicesFiltrados = null;
  if (colunasDesejadas && colunasDesejadas.length > 0) {
    indicesFiltrados = colunasDesejadas
      .map(nomeCol => cabecalho.findIndex(h => h.toLowerCase().includes(nomeCol.toLowerCase())))
      .filter(idx => idx !== -1);
  }

  const dados = [];
  linhas.forEach(tr => {
    const celulas = Array.from(tr.querySelectorAll('th, td'));
    let textos = celulas.map(valorCelula);
    if (indicesFiltrados) {
      textos = indicesFiltrados.map(idx => textos[idx] ?? '');
    }
    if (textos.some(t => t !== '')) dados.push(textos);
  });

  // Nome da aba: pega o ÚLTIMO <h1> válido da página (ignorando
  // placeholders tipo "Carregando dados..."), que é o título real da tela.
  const candidatosH1 = Array.from(document.querySelectorAll('h1'))
    .map(el => el.innerText.trim())
    .filter(t => t && !t.toLowerCase().includes('carregando'));

  const nomeAba = candidatosH1.length > 0
    ? candidatosH1[candidatosH1.length - 1]
    : document.title;

  return { dados, nomeAba };
}