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
  
  function nomeValidoExcel(titulo, index) {
    let nome = titulo || 'Aba';
    nome = nome.replace(/[\\/*?:[\]]/g, '').trim();
    if (nome.length > 25) nome = nome.substring(0, 22) + '...';
    return `${nome}_${index + 1}`.substring(0, 31);
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
            const nome = nomeValidoExcel(nomeAba || aba.title, i);
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

  // Nome da aba: tenta achar um título de verdade na página
  const heading = document.querySelector('h1, h2, .page-title, .titulo-pagina, .breadcrumb-item.active');
  const nomeAba = heading ? heading.innerText.trim() : document.title;

  return { dados, nomeAba };
}