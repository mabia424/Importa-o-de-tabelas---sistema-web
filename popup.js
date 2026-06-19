// ============================================
// popup.js - Com verificação de XLSX
// ============================================

document.addEventListener('DOMContentLoaded', function() {
  
  const btnExportar = document.getElementById('btnExportar');
  const status = document.getElementById('status');
  const contador = document.getElementById('contador');
  
  const DOMINIO = 'adekz.jawplasticos.com.br';
  
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
            func: extrairTabelaDaPagina
          });
          
          const dados = resultado[0]?.result;
          
          if (dados && dados.length > 0) {
            const worksheet = XLSX.utils.aoa_to_sheet(dados);
            const nomeAba = nomeValidoExcel(aba.title, i);
            XLSX.utils.book_append_sheet(workbook, worksheet, nomeAba);
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

function extrairTabelaDaPagina() {
  const tabelas = document.querySelectorAll('table');
  if (tabelas.length === 0) return [];
  
  let maior = tabelas[0];
  let max = 0;
  
  tabelas.forEach(t => {
    const n = t.querySelectorAll('tr').length;
    if (n > max) { max = n; maior = t; }
  });
  
  const dados = [];
  maior.querySelectorAll('tr').forEach(tr => {
    const textos = Array.from(tr.querySelectorAll('th, td')).map(c => c.innerText.trim());
    if (textos.some(t => t !== '')) dados.push(textos);
  });
  
  return dados;
}