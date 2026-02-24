/* script-fusionado.js
   Unificación: empleados únicos que sirven para PDV y Domingos.
   Domingos lee/actualiza el mismo array 'empregados' guardado en localStorage.
*/

/* =========================
   Utilidades para Domingos
   ========================= */

function dom_domingosDoMes(monthIndex, year){
  const arr = [];
  // buscamos todos los domingos entre 1 y 31 (algunos meses 30/31)
  for(let d=1; d<=31; d++){
    const dt = new Date(year, monthIndex, d);
    if(dt.getMonth() !== monthIndex) continue;
    if(dt.getDay() === 0) arr.push(new Date(dt.getFullYear(), dt.getMonth(), dt.getDate()));
  }
  return arr;
}

function dom_semanasEntre(primeiro, outro){
  const msWeek = 7 * 24 * 3600 * 1000;
  const utcP = Date.UTC(primeiro.getFullYear(), primeiro.getMonth(), primeiro.getDate());
  const utcO = Date.UTC(outro.getFullYear(), outro.getMonth(), outro.getDate());
  return Math.floor((utcO - utcP) / msWeek);
}

function dom_isFolga(primeiroDomingoStr, domingoDate){
  if(!primeiroDomingoStr) return false;
  const parts = primeiroDomingoStr.split('-').map(Number);
  const base = new Date(parts[0], parts[1]-1, parts[2]);
  if(isNaN(base.getTime())) return false;
  const diff = dom_semanasEntre(base, domingoDate);
  if(isNaN(diff)) return false;
  const idx = ((diff % 4) + 4) % 4;
  return idx === 3;
}

/* =========================
   Storage y datos globales
   ========================= */

let empregados = JSON.parse(localStorage.getItem('condor_empregados') || '[]');
let atribuicoes = JSON.parse(localStorage.getItem('condor_atribuicoes') || '[]');
let turnoAtual = 'manha';

function salvarDados() {
  // Normalizamos estructura para compatibilidad: aseguramos keys que usa DOM
  empregados = empregados.map(e => ({
    nome: e.nome || '',
    tipo: e.tipo || 'condor',
    livre: e.livre || '',
    turno: e.turno || '09:00',
    categoriaDom: e.categoriaDom || '',         // categoria para domingos (empaquetador/caja/fiscal)
    primeiroDomingo: e.primeiroDomingo || ''    // YYYY-MM-DD o ''
  }));
  localStorage.setItem('condor_empregados', JSON.stringify(empregados));
  localStorage.setItem('condor_atribuicoes', JSON.stringify(atribuicoes));
}

/* Inicialización DOMContentLoaded */
document.addEventListener('DOMContentLoaded', () => {
  preencherPDV();
  configurarSeletorHora();
  renderEmpregados();
  renderAtribuicoes();
  renderFolgas();
  mostrar('principal');

  // eventos PDV
  const dataInput = document.getElementById('data');
  if (dataInput) {
    dataInput.addEventListener('change', () => {
      renderEmpregadosDisponiveis();
      renderAtribuicoes();
    });
  }

  // inicializamos eventos del módulo Domingos
  dom_initEvents();
  // render inicial domingos
  dom_renderPlanilha();
});

/* ====== Mostrar secciones (global) ====== */
function mostrar(id) {
  document.querySelectorAll('.section').forEach(s => s.style.display = 'none');
  const el = document.getElementById(id);
  if (el) el.style.display = 'block';
  if (id === 'folgas') renderFolgas();
}

/* ====== Hora selector (flatpickr) ====== */
function configurarSeletorHora() {
  flatpickr("#empTurno", {
    enableTime: true,
    noCalendar: true,
    dateFormat: "H:i",
    time_24hr: true,
    minuteIncrement: 15
  });
}

function abrirSeletorHora() {
  const el = document.getElementById('empTurno');
  if (el && el._flatpickr) el._flatpickr.open();
  else if(el) el.focus();
}

/* ====== EMPREGADOS (PDV + Domingo fields) ====== */
function adicionarEmpregado() {
  const nome = document.getElementById('empNome').value.trim();
  const tipo = document.getElementById('empTipo').value;
  const livre = document.getElementById('empLivre').value;
  const turno = document.getElementById('empTurno').value.trim();
  const categoriaDom = document.getElementById('empCategoriaDom').value || '';
  const primeiroDom = document.getElementById('empPrimeiroDom').value || '';

  if (!nome || !turno) { alert('Preencha nome e horário.'); return; }
  if (!/^\d{2}:\d{2}$/.test(turno)) { alert('Formato de horário inválido.'); return; }

  // si existe, actualizamos; si no, agregamos
  const idx = empregados.findIndex(e => e.nome.toLowerCase() === nome.toLowerCase());
  if(idx >= 0){
    const existente = empregados[idx];
    existente.tipo = tipo;
    existente.livre = livre;
    existente.turno = turno;
    existente.categoriaDom = categoriaDom;
    existente.primeiroDomingo = primeiroDom;
    empregados[idx] = existente;
  } else {
    empregados.push({ nome, tipo, livre, turno, categoriaDom, primeiroDomingo: primeiroDom });
  }

  salvarDados();
  renderEmpregados();
  limparFormularioEmp();
  renderFolgas();
  dom_renderPlanilha();
}

function limparFormularioEmp() {
  const elNome = document.getElementById('empNome');
  const elTurno = document.getElementById('empTurno');
  const elCatDom = document.getElementById('empCategoriaDom');
  const elPrimDom = document.getElementById('empPrimeiroDom');
  if(elNome) elNome.value = '';
  if(elTurno) elTurno.value = '';
  if(elCatDom) elCatDom.value = '';
  if(elPrimDom) elPrimDom.value = '';
}

function renderEmpregados() {
  const ulManha = document.getElementById('listaManha');
  const ulTarde = document.getElementById('listaTarde');
  if(!ulManha || !ulTarde) return;
  ulManha.innerHTML = '';
  ulTarde.innerHTML = '';

  const manha = empregados.filter(e => parseInt(e.turno.split(':')[0]) < 12);
  const tarde = empregados.filter(e => parseInt(e.turno.split(':')[0]) >= 12);

  const criarLi = e => `
    <div>
      <strong>${e.nome}</strong> <small>(${e.tipo})</small><br>
      <small>Início: ${e.turno} | Folga: ${e.livre}</small><br>
      <small class="small">Domingos: ${e.categoriaDom || '—'} ${e.primeiroDomingo ? '| 1º: ' + e.primeiroDomingo : ''}</small>
    </div>
    <div>
      <button class="sec" onclick="editarEmpregado('${e.nome}')">Editar</button>
      <button onclick="removerEmpregado('${e.nome}')">Remover</button>
    </div>`;

  manha.forEach(e => {
    const li = document.createElement('li');
    li.innerHTML = criarLi(e);
    ulManha.appendChild(li);
  });

  tarde.forEach(e => {
    const li = document.createElement('li');
    li.innerHTML = criarLi(e);
    ulTarde.appendChild(li);
  });

  renderEmpregadosDisponiveis();
}

function removerEmpregado(nome) {
  if (!confirm(`Remover ${nome}?`)) return;
  empregados = empregados.filter(e => e.nome !== nome);
  atribuicoes = atribuicoes.filter(a => a.colaborador !== nome);
  salvarDados();
  renderEmpregados();
  renderAtribuicoes();
  renderFolgas();
  dom_renderPlanilha();
}

function editarEmpregado(nome) {
  const e = empregados.find(x => x.nome === nome);
  if (!e) return;

  const novoNome = prompt('Editar nome:', e.nome)?.trim() || e.nome;
  const novoTurno = prompt('Editar turno (HH:MM):', e.turno)?.trim() || e.turno;
  const novoLivre = prompt('Editar folga (ex: segunda):', e.livre)?.trim().toLowerCase() || e.livre;
  const novoTipo = prompt('Editar tipo (condor/terceirizado):', e.tipo)?.trim().toLowerCase() || e.tipo;
  const novaCategoriaDom = prompt('Categoria Domingos (empaquetador/caja/fiscal ou vazio):', e.categoriaDom || '')?.trim() || '';
  const novoPrimeiroDom = prompt('1º domingo trabalhado (YYYY-MM-DD ou vazio):', e.primeiroDomingo || '')?.trim() || '';

  // actualizar referencias en atribuicoes
  atribuicoes.forEach(a => {
    if (a.colaborador === e.nome) a.colaborador = novoNome;
  });

  e.nome = novoNome;
  e.turno = novoTurno;
  e.livre = novoLivre;
  e.tipo = novoTipo;
  e.categoriaDom = novaCategoriaDom;
  e.primeiroDomingo = novoPrimeiroDom;

  salvarDados();
  renderEmpregados();
  renderAtribuicoes();
  renderFolgas();
  dom_renderPlanilha();
}

/* ====== PDV: preencher PDV list ====== */
function preencherPDV() {
  const sel = document.getElementById('pdvNumero');
  if(!sel) return;
  sel.innerHTML = '';
  for (let i = 101; i <= 123; i++) {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = `PDV ${i}`;
    sel.appendChild(opt);
  }
}

/* ====== Turno ====== */
function setarTurno(t) {
  turnoAtual = t;
  const btnM = document.getElementById('btnManha');
  const btnT = document.getElementById('btnTarde');
  if(btnM) btnM.classList.toggle('active', t === 'manha');
  if(btnT) btnT.classList.toggle('active', t === 'tarde');
  renderEmpregadosDisponiveis();
  renderAtribuicoes();
}

/* ====== Disponibilidade ====== */
function renderEmpregadosDisponiveis() {
  const sel = document.getElementById('empDisponivel');
  if(!sel) return;
  sel.innerHTML = '';

  const data = document.getElementById('data')?.value;
  let dia = '';
  if (data) {
    dia = new Date(data).toLocaleDateString('pt-BR', { weekday: 'long' }).toLowerCase();
  }

  const disponiveis = empregados.filter(e => {
    const hora = parseInt(e.turno.split(':')[0]);
    const turno = hora < 12 ? 'manha' : 'tarde';
    if (e.livre === dia) return false;
    return turno === turnoAtual;
  });

  disponiveis.forEach(e => {
    const opt = document.createElement('option');
    opt.value = e.nome;
    opt.textContent = `${e.nome} (${e.tipo}, ${e.turno})`;
    sel.appendChild(opt);
  });

  if (disponiveis.length === 0) {
    const opt = document.createElement('option');
    opt.textContent = 'Nenhum colaborador disponível';
    sel.appendChild(opt);
  }
}

/* ====== Atribuições ====== */
function atribuir() {
  const data = document.getElementById('data')?.value;
  const pdv = document.getElementById('pdvNumero')?.value;
  const colaborador = document.getElementById('empDisponivel')?.value;
  const emp = empregados.find(e => e.nome === colaborador);

  if (!data || !pdv || !emp) { alert('Preencha todos os campos.'); return; }

  const existe = atribuicoes.find(a => a.data === data && a.pdv === pdv && a.turno === turnoAtual);
  if (existe) { alert('Este PDV já está atribuído neste turno.'); return; }

  atribuicoes.push({
    data,
    pdv,
    colaborador: emp.nome,
    tipo: emp.tipo,
    turno: turnoAtual
  });

  salvarDados();
  renderAtribuicoes();
}

function renderAtribuicoes() {
  const tbody = document.querySelector('#tabelaAtribuicoes tbody');
  if(!tbody) return;
  tbody.innerHTML = '';

  const dataSelecionada = document.getElementById('data')?.value;
  const atribuicoesFiltradas = dataSelecionada ? atribuicoes.filter(a => a.data === dataSelecionada) : atribuicoes;

  atribuicoesFiltradas
    .sort((a, b) => String(a.pdv).localeCompare(String(b.pdv)))
    .forEach(a => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${a.data}</td>
        <td>${a.turno}</td>
        <td>${a.pdv}</td>
        <td>${a.colaborador}</td>
        <td>${a.tipo}</td>
        <td><button onclick="removerAtribuicao('${a.data}','${a.pdv}','${a.turno}')">❌</button></td>`;
      tbody.appendChild(tr);
    });
}

function removerAtribuicao(data, pdv, turno) {
  atribuicoes = atribuicoes.filter(a => !(a.data === data && a.pdv === pdv && a.turno === turno));
  salvarDados();
  renderAtribuicoes();
}

/* ====== Folgas (PDV) ====== */
function renderFolgas() {
  const cont = document.getElementById('folgasContainer');
  cont.innerHTML = '';

  const dias = ['segunda','terca','quarta','quinta','sexta','sabado','domingo'];

  function criarBloco(titulo, filtroFn, turnoKey){
    const wrapper = document.createElement('div');
    wrapper.classList.add('folga-wrapper');

    wrapper.innerHTML = `
      <h2 style="text-align:center;">${titulo}</h2>
      <div class="folgas-grid" id="grid-${turnoKey}"></div>
      <div style="text-align:center;">
        <button onclick="baixarFolgasTurno('${turnoKey}')">Baixar ${titulo}</button>
      </div>
      <hr>
    `;

    cont.appendChild(wrapper);

    const grid = wrapper.querySelector(`#grid-${turnoKey}`);

    dias.forEach(dia => {
      const card = document.createElement('div');
      card.classList.add('folga-card');

      const lista = empregados.filter(e =>
        e.livre === dia && filtroFn(e)
      );

      card.innerHTML = `
        <h4>${dia.charAt(0).toUpperCase() + dia.slice(1)}</h4>
        <ul>${lista.map(e => `<li>${e.nome} (${e.turno})</li>`).join('') || '<li>Ninguém</li>'}</ul>
      `;

      grid.appendChild(card);
    });
  }

  // EMPLEADOS NORMALES
  criarBloco("Turno MANHÃ", e => e.tipo !== "terceirizado" && parseInt(e.turno.split(':')[0]) < 12, "manha");
  criarBloco("Turno TARDE", e => e.tipo !== "terceirizado" && parseInt(e.turno.split(':')[0]) >= 12, "tarde");

  // TERCIARIZADOS
  criarBloco("Terceirizados MANHÃ", e => e.tipo === "terceirizado" && parseInt(e.turno.split(':')[0]) < 12, "tman");
  criarBloco("Terceirizados TARDE", e => e.tipo === "terceirizado" && parseInt(e.turno.split(':')[0]) >= 12, "ttar");
}

/* =========================
   DOMINGOS: UI y acciones
   Ahora trabajan con 'empregados' (unificado)
   ========================= */

const domBtnAdicionar = () => document.getElementById('domBtnAdicionar');
const domNomeInput = () => document.getElementById('domNome');
const domCatInput = () => document.getElementById('domCategoria');
const domPrimeiroInput = () => document.getElementById('domPrimeiro');
const domTabs = () => document.querySelectorAll('#domingo .tab');
const domMesSelect = () => document.getElementById('domMes');
const domAnoInput = () => document.getElementById('domAno');
const domBtnCarregarEl = () => document.getElementById('domBtnCarregar');
const domBtnExportarEl = () => document.getElementById('domBtnExportar');
const domTabela = () => document.getElementById('domTabelaPlanilha');
const domTitulo = () => document.getElementById('domTituloPlanilha');

let dom_categoriaAtiva = 'empaquetador';
if(domMesSelect()) domMesSelect().value = '10';
if(domAnoInput()) domAnoInput().value = '2025';

function dom_initEvents(){
  const btn = domBtnAdicionar();
  if(btn) btn.addEventListener('click', dom_adicionar);

  const tabs = domTabs();
  if(tabs) tabs.forEach(t => t.addEventListener('click', ()=>{
    tabs.forEach(x=>x.classList.remove('active'));
    t.classList.add('active');
    dom_categoriaAtiva = t.dataset.cat;
    dom_renderPlanilha();
  }));

  const car = domBtnCarregarEl();
  if(car) car.addEventListener('click', dom_renderPlanilha);
  const exp = domBtnExportarEl();
  if(exp) exp.addEventListener('click', dom_exportarExcel);
}

function dom_adicionar(){
  const nome = domNomeInput().value.trim();
  const categoria = domCatInput().value;
  const primeiroDom = domPrimeiroInput().value;
  if(!nome || !primeiroDom){ alert('Complete nome e 1º domingo'); return; }
  const [y,m,d] = primeiroDom.split('-').map(Number);
  const dt = new Date(y,m-1,d);
  if(dt.getDay() !== 0){ alert('A data escolhida não é domingo.'); return; }

  const idx = empregados.findIndex(e => e.nome.toLowerCase() === nome.toLowerCase());
  if(idx >= 0){
    // atualiza empleado existente
    empregados[idx].categoriaDom = categoria;
    empregados[idx].primeiroDomingo = primeiroDom;
  } else {
    // cria mínimo (recomendado cadastrar por formulario principal)
    empregados.push({
      nome,
      tipo: 'condor',
      livre: '',
      turno: '09:00',
      categoriaDom: categoria,
      primeiroDomingo: primeiroDom
    });
  }

  salvarDados();
  domNomeInput().value = ''; domPrimeiroInput().value = '';
  renderEmpregados();
  dom_renderPlanilha();
}

function dom_editarEmpregado(nome){
  // abrimos prompts pero actualizamos el registro en empregados
  const emp = empregados.find(e=>e.nome===nome);
  if(!emp) return;
  const novoNome = prompt('Nome:', emp.nome);
  if(novoNome === null) return;
  const novaCategoria = prompt('Categoria (empaquetador/caja/fiscal):', emp.categoriaDom || '');
  if(novaCategoria === null) return;
  const novaData = prompt('1º domingo trabalhado (YYYY-MM-DD):', emp.primeiroDomingo || '');
  if(novaData === null) return;
  const [y,m,d] = novaData.split('-').map(Number);
  const dt = new Date(y,m-1,d);
  if(dt.getDay() !== 0){ alert('Data não é domingo. Não salvo.'); return; }
  emp.nome = novoNome.trim();
  emp.categoriaDom = novaCategoria.trim();
  emp.primeiroDomingo = novaData;
  salvarDados();
  renderEmpregados();
  dom_renderPlanilha();
}
function dom_removerEmpregado(nome){
  if(!confirm('Remover funcionário?')) return;
  empregados = empregados.filter(e=>e.nome!==nome);
  salvarDados();
  renderEmpregados();
  dom_renderPlanilha();
}

function dom_cellTH(txt){ const th = document.createElement('th'); th.textContent = txt; return th; }

function dom_renderPlanilha(){
  if(!domTabela()) return;
  const mes = parseInt(domMesSelect().value,10);
  const ano = parseInt(domAnoInput().value,10);
  const cat = dom_categoriaAtiva;
  const nomeCat = cat === 'caja' ? 'Operador de Caixa' : (cat === 'empaquetador' ? 'Empacotador' : 'Fiscal');

  const dataRef = new Date(ano, mes, 1);
  const mesNome = dataRef.toLocaleString('pt-BR',{month:'long', year:'numeric'});
  domTitulo().textContent = `Planilha — ${nomeCat} — ${mesNome}`;

  const lista = empregados.filter(e=>e.categoriaDom === cat);

  const tabela = domTabela();
  tabela.innerHTML = '';
  const thead = document.createElement('thead');
  const trh = document.createElement('tr');
  trh.appendChild(dom_cellTH('Funcionário'));
  // mostramos máximo 31 dias (el render de DOM funciona con meses reales)
  for(let d=1; d<=31; d++) trh.appendChild(dom_cellTH(String(d)));
  trh.appendChild(dom_cellTH('Ações'));
  thead.appendChild(trh);
  tabela.appendChild(thead);

  const tbody = document.createElement('tbody');

  if(lista.length === 0){
    const trEmpty = document.createElement('tr');
    const td = document.createElement('td');
    td.colSpan = 33;
    td.textContent = 'Não há funcionários nesta categoria.';
    td.style.textAlign = 'center';
    trEmpty.appendChild(td);
    tbody.appendChild(trEmpty);
    tabela.appendChild(tbody);
    return;
  }

  const domingos = dom_domingosDoMes(mes, ano);
  for(const emp of lista){
    const tr = document.createElement('tr');
    const tdNome = document.createElement('td');
    tdNome.classList.add('employee-cell');
    tdNome.textContent = emp.nome;
    tr.appendChild(tdNome);

    const mapa = {};
    for(const sd of domingos){
      mapa[sd.getDate()] = { folga: dom_isFolga(emp.primeiroDomingo, sd) };
    }

    // 1..31
    for(let d=1; d<=31; d++){
      const td = document.createElement('td');
      if(mapa[d]){
        if(mapa[d].folga){
          td.textContent = 'F';
          td.classList.add('folga');
        } else {
          td.textContent = 'TB';
          td.classList.add('trabalha');
        }
      } else {
        td.textContent = '';
      }
      tr.appendChild(td);
    }

    const tdAcoes = document.createElement('td');
    const btnE = document.createElement('button');
    btnE.textContent = 'Editar';
    btnE.style.marginRight = '6px';
    btnE.onclick = ()=> dom_editarEmpregado(emp.nome);
    const btnX = document.createElement('button');
    btnX.textContent = 'Excluir';
    btnX.onclick = ()=> dom_removerEmpregado(emp.nome);
    tdAcoes.appendChild(btnE);
    tdAcoes.appendChild(btnX);
    tr.appendChild(tdAcoes);

    tbody.appendChild(tr);
  }

  tabela.appendChild(tbody);
}

/* Exportar XLSX con SheetJS para Domingos (usa empleados unificados) */
async function dom_exportarExcel() {
  const mes = parseInt(domMesSelect().value,10);
  const ano = parseInt(domAnoInput().value,10);
  const cat = dom_categoriaAtiva;
  const nomeCat = cat === 'caja' ? 'Operador Caixa' :
                  cat === 'empaquetador' ? 'Empacotador' :
                  'Fiscal';

  const lista = empregados.filter(e => e.categoriaDom === cat);

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Domingos");

  // Título grande
  sheet.mergeCells("A1:AG1");
  sheet.getCell("A1").value = `PLANILHA DE DOMINGOS — ${nomeCat.toUpperCase()} — ${mes+1}/${ano}`;
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getCell("A1").font = { bold: true, size: 16 };

  // Header
  const header = ["Funcionário"];
  for (let d = 1; d <= 31; d++) header.push(d);
  sheet.addRow(header);

  // pinta header
  sheet.getRow(2).eachCell(c=>{
    c.font = { bold: true };
    c.alignment = { horizontal: "center" };
    c.border = { top:{style:'thin'}, right:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'} };
  });

  // datos
  for(const emp of lista){
    const row = [emp.nome];
    const domingos = dom_domingosDoMes(mes, ano);
    const mapa = {};
    for(const sd of domingos) mapa[sd.getDate()] = dom_isFolga(emp.primeiroDomingo, sd);

    for(let d = 1; d <= 31; d++){
      if(mapa[d] === undefined) row.push("");
      else row.push(mapa[d] ? "F" : "TB");
    }
    sheet.addRow(row);
  }

  // Anchos de columna
  sheet.getColumn(1).width = 30;
  for (let i = 2; i <= 32; i++) sheet.getColumn(i).width = 4;

  // bordes y colores
  sheet.eachRow((row, rowIdx)=>{
    if(rowIdx < 2) return;
    row.eachCell((cell)=>{
      cell.alignment = { horizontal: "center" };
      cell.border = { top:{style:'thin'}, right:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'} };

      if(cell.value === "F") cell.fill = { type: "pattern", pattern: "solid", fgColor:{ argb:"FFECECEC" } };
      if(cell.value === "TB") cell.fill = { type: "pattern", pattern: "solid", fgColor:{ argb:"FFDFF2E6" } };
    });
  });

  // descargar
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer]);
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `Domingos_${nomeCat}_${mes+1}-${ano}.xlsx`;
  link.click();
}
// ============================
// EXPORTAR FOLGAS PARA EXCEL (con formato tipo tarjetas)
// ============================
async function baixarFolgasTurno(turnoKey) {

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(turnoKey.toUpperCase());

  const titulo = {
    manha: "Turno MANHÃ",
    tarde: "Turno TARDE",
    tman: "Terceirizados MANHÃ",
    ttar: "Terceirizados TARDE"
  }[turnoKey];

  const estiloHeader = {
    font: { bold: true, size: 14, color: { argb: "FFFFFFFF" } },
    alignment: { horizontal: "center", vertical: "middle" },
    fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E2A78" } }
  };

  const estiloCard = {
    border: {
      top: { style: "thin" }, right: { style: "thin" },
      bottom: { style: "thin" }, left: { style: "thin" }
    },
    alignment: { vertical: "middle" }
  };

  sheet.mergeCells("A1:H1");
  sheet.getCell("A1").value = titulo;
  sheet.getCell("A1").font = { bold: true, size: 18 };
  sheet.getCell("A1").alignment = { horizontal: "center" };

  const dias = ["segunda","terca","quarta","quinta","sexta","sabado","domingo"];

  dias.forEach((dia, i) => {
    const col = i + 1;

    sheet.getCell(2, col).value = dia.toUpperCase();
    Object.assign(sheet.getCell(2, col), estiloHeader);

    const lista = empregados.filter(e => {
      const hora = parseInt(e.turno.split(":")[0]);
      const manha = hora < 12;
      const tarde = hora >= 12;

      if (turnoKey === "manha") return e.tipo !== "terceirizado" && manha && e.livre === dia;
      if (turnoKey === "tarde") return e.tipo !== "terceirizado" && tarde && e.livre === dia;
      if (turnoKey === "tman")  return e.tipo === "terceirizado" && manha && e.livre === dia;
      if (turnoKey === "ttar")  return e.tipo === "terceirizado" && tarde && e.livre === dia;
    });

    let row = 3;
    if (lista.length === 0) {
      sheet.getCell(row, col).value = "Nenhum";
      Object.assign(sheet.getCell(row, col), estiloCard);
    } else {
      lista.forEach(e => {
        sheet.getCell(row, col).value = `${e.nome} (${e.turno})`;
        Object.assign(sheet.getCell(row, col), estiloCard);
        row++;
      });
    }

    sheet.getColumn(col).width = 22;
  });

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer]);
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${titulo}.xlsx`;
  link.click();
}

/* ============================
   DESCARGAR TODOS LOS TURNOS
   ============================ */
async function baixarFolgasTodas() {
  const workbook = new ExcelJS.Workbook();

  async function agregarHoja(turnoKey, titulo) {
    const sheet = workbook.addWorksheet(titulo.toUpperCase());

    const estiloHeader = {
      font: { bold: true, size: 14, color: { argb: "FFFFFFFF" } },
      alignment: { horizontal: "center", vertical: "middle" },
      fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FF1E2A78" } }
    };

    const estiloCard = {
      border: {
        top: { style: "thin" }, right: { style: "thin" },
        bottom: { style: "thin" }, left: { style: "thin" }
      },
      alignment: { vertical: "middle" }
    };

    const dias = ["segunda","terca","quarta","quinta","sexta","sabado","domingo"];

    sheet.mergeCells("A1:H1");
    sheet.getCell("A1").value = titulo;
    sheet.getCell("A1").font = { bold: true, size: 18 };
    sheet.getCell("A1").alignment = { horizontal: "center" };

    dias.forEach((dia, i) => {
      const col = i + 1;
      sheet.getCell(2, col).value = dia.toUpperCase();
      Object.assign(sheet.getCell(2, col), estiloHeader);

      const lista = empregados.filter(e => {
        const hora = parseInt(e.turno.split(":")[0]);
        const manha = hora < 12;
        const tarde = hora >= 12;

        if (turnoKey === "manha") return e.tipo !== "terceirizado" && manha && e.livre === dia;
        if (turnoKey === "tarde") return e.tipo !== "terceirizado" && tarde && e.livre === dia;
        if (turnoKey === "tman")  return e.tipo === "terceirizado" && manha && e.livre === dia;
        if (turnoKey === "ttar")  return e.tipo === "terceirizado" && tarde && e.livre === dia;
      });

      let row = 3;
      if (lista.length === 0) {
        sheet.getCell(row, col).value = "Nenhum";
        Object.assign(sheet.getCell(row, col), estiloCard);
      } else {
        lista.forEach(e => {
          sheet.getCell(row, col).value = `${e.nome} (${e.turno})`;
          Object.assign(sheet.getCell(row, col), estiloCard);
          row++;
        });
      }

      sheet.getColumn(col).width = 22;
    });
  }

  await agregarHoja("manha", "Turno MANHÃ");
  await agregarHoja("tarde", "Turno TARDE");
  await agregarHoja("tman", "Terceirizados MANHÃ");
  await agregarHoja("ttar", "Terceirizados TARDE");

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer]);
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "Folgas_COMPLETO.xlsx";
  link.click();
}
// ============================
// EXPORTAR ASIGNACIONES DIARIAS
// ============================
async function baixarExcel(turnoEscolhido) {
  const dataSelecionada = document.getElementById('data')?.value;
  if (!dataSelecionada) { alert("Selecione uma data."); return; }

  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet(turnoEscolhido === 'manha' ? "Manhã" : "Tarde");

  sheet.mergeCells("A1:E1");
  sheet.getCell("A1").value = turnoEscolhido === "manha" ? "7:00 – 11:00" : "12:00 – 15:30";
  sheet.getCell("A1").font = { bold: true, size: 14 };
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };

  sheet.columns = [
    { header: "DATA", width: 12 },
    { header: "PDV", width: 6 },
    { header: "NOME", width: 26 },
    { header: "TIPO", width: 15 },
    { header: "F/S", width: 15 }
  ];

  const header = sheet.getRow(2);
  header.values = ["DATA","PDV","NOME","TIPO","F/S"];
  header.eachCell(c=>{
    c.fill = { type:"pattern", pattern:"solid", fgColor:{ argb:"FFCCCCCC" }};
    c.font = { bold:true };
    c.alignment = { horizontal:"center" };
    c.border = { top:{style:'thin'}, bottom:{style:'thin'}, left:{style:'thin'}, right:{style:'thin'} };
  });

  const atribs = atribuicoes.filter(a => a.data === dataSelecionada);

  const condor = [];
  const terce = [];

  atribs.forEach(a => {
    const emp = empregados.find(e => e.nome === a.colaborador);
    if (!emp) return;

    const hora = parseInt(emp.turno.split(":")[0]);
    const turnoEmp = hora < 12 ? "manha" : "tarde";
    if (turnoEmp !== turnoEscolhido) return;

    const destino = emp.tipo === "terceirizado" ? terce : condor;
    destino.push([
      a.data,
      a.pdv,
      emp.nome,
      emp.tipo.toUpperCase(),
      emp.livre?.toUpperCase() || ""
    ]);
  });

  function addTitulo(texto){
    const row = sheet.addRow([texto]);
    sheet.mergeCells(`A${row.number}:E${row.number}`);
    row.getCell(1).fill = { type:"pattern", pattern:"solid", fgColor:{ argb:"FFEFEFEF" }};
    row.getCell(1).font = { bold:true };
    row.getCell(1).alignment = { horizontal:"left" };
  }

  if (condor.length) {
    addTitulo("CONDOR");
    condor.forEach(l => sheet.addRow(l));
  }

  if (terce.length) {
    addTitulo("TERCEIRIZADOS");
    terce.forEach(l => sheet.addRow(l));
  }

  sheet.eachRow(r=>{
    r.eachCell(c=>{
      c.border = { top:{style:"thin"}, left:{style:"thin"}, bottom:{style:"thin"}, right:{style:"thin"} };
      c.alignment = { horizontal:"center" };
    });
  });

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer]);
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `ESCALA_${turnoEscolhido}_${dataSelecionada}.xlsx`;
  link.click();
}



/* FIN */
