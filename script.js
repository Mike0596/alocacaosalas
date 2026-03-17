// =========================================================
// BANCO DE DADOS DE ALTA PERFORMANCE (IndexedDB) E GERADOR DE ID
// =========================================================
const dbName = "DashboardSalasDB";
const storeName = "dadosSalas";
let masterIdCounter = Date.now() % 100000;

function gerarIdUnico() {
    return "ID-" + (masterIdCounter++).toString(36).toUpperCase();
}

function initDB(callback) {
    let request = indexedDB.open(dbName, 1);
    request.onupgradeneeded = function(e) {
        let db = e.target.result;
        if (!db.objectStoreNames.contains(storeName)) {
            db.createObjectStore(storeName, { keyPath: "id" });
        }
    };
    request.onsuccess = function(e) { callback(e.target.result); };
    request.onerror = function(e) { console.error("Erro no Banco de Dados:", e); };
}

function salvarDadosDB(dadosArray) {
    initDB(function(db) {
        let tx = db.transaction(storeName, "readwrite");
        let store = tx.objectStore(storeName);
        store.put({ id: 1, payload: dadosArray });
    });
}

function carregarDadosDB(callback) {
    initDB(function(db) {
        let tx = db.transaction(storeName, "readonly");
        let store = tx.objectStore(storeName);
        let request = store.get(1);
        request.onsuccess = function() {
            callback(request.result ? request.result.payload : []);
        };
    });
}

function salvarManutencaoDB() {
    initDB(function(db) {
        let tx = db.transaction(storeName, "readwrite");
        let store = tx.objectStore(storeName);
        store.put({ id: 2, bloqueadas: salasEmManutencao, projetorRuim: salasProjetorRuim, turmasOcultas: turmasOcultas });
    });
}

function carregarManutencaoDB(callback) {
    initDB(function(db) {
        let tx = db.transaction(storeName, "readonly");
        let store = tx.objectStore(storeName);
        let request = store.get(2);
        request.onsuccess = function() {
            if (request.result) {
                salasEmManutencao = request.result.bloqueadas || [];
                salasProjetorRuim = request.result.projetorRuim || [];
                turmasOcultas = request.result.turmasOcultas || []; 
            }
            callback();
        };
    });
}

function limparDB() {
    initDB(function(db) {
        let tx = db.transaction(storeName, "readwrite");
        // O comando 'delete(1)' apaga APENAS a gaveta de dados importados (ID 1)
        // Isso protege as Configurações (ID 3), Mapa (ID 2) e Consolidados (ID 4)
        tx.objectStore(storeName).delete(1); 
    });
}

// =========================================================
// BANCO DE DADOS DAS CONFIGURAÇÕES E REGRAS
// =========================================================
let regrasAlocacao = { 
    manterProf: true, 
    prioridadeCalouros: true, 
    margemFolga: true, 
    manterProfSemestre: false, 
    ordemBlocos: "E, G, F, C",
    travaSaude: true,    // Nova
    travaEngArq: true,   
    prioridadeTamanho: true    // Nova
};

function salvarConfigDB() {
    initDB(function(db) {
        let tx = db.transaction(storeName, "readwrite");
        let store = tx.objectStore(storeName);
        store.put({ id: 3, salasSalvas: salas, blocosSalvos: blocosCadastrados, regras: regrasAlocacao });
    });
}

function carregarConfigDB(callback) {
    initDB(function(db) {
        let tx = db.transaction(storeName, "readonly");
        let store = tx.objectStore(storeName);
        let request = store.get(3);
        request.onsuccess = function() {
            if (request.result) {
                if (request.result.salasSalvas && request.result.salasSalvas.length > 0) {
                    salas = request.result.salasSalvas;
                }
                if (request.result.blocosSalvos && request.result.blocosSalvos.length > 0) {
                    blocosCadastrados = request.result.blocosSalvos;
                }
                if (request.result.regras) {
                    regrasAlocacao = request.result.regras;
                }
            }
            callback();
        };
    });
}

// Inicialização Sequencial do Sistema
window.addEventListener('DOMContentLoaded', () => {
    construirFiltrosAvancados();
    carregarConsolidadoDB(); 
    carregarConfigPdfDB(); // <-- Nova linha para puxar a imagem e configurações do PDF!
    carregarConfigDB(function() {
        carregarManutencaoDB(function() {
            popularSalasManutencao();
            carregarDadosDB(function(d) {
                if (d && d.length > 0) {
                    let mudou = false;
                    d.forEach(linha => {
                        if(!linha.idUnico) { linha.idUnico = gerarIdUnico(); mudou = true; }
                    });
                    dados = d;
                    if(mudou) salvarDadosDB(dados);
                    atualizarFiltrosDinamicos();
                    executarProcessamento(); 
                } else {
                    ocultarLoading();
                }
                
                if (typeof renderizarMapaSalas === "function") renderizarMapaSalas();
                renderizarConfiguracoes();
            });
        });
    });
});

// =========================================================
// LÓGICA DE DADOS GLOBAIS E DEFINIÇÕES
// =========================================================
let dados = [];
let windowDadosResultantes = []; 
let renderId = 0; 
let chartDiasInstancia = null;
let chartCursosInstancia = null;

let salasEmManutencao = []; 
let salasProjetorRuim = [];
let salasSelecionadas = [];
let turmasOcultas = [];
// INSIRA ESTE BLOCO NO SEU JS:
let dadosConsolidados = {};

function salvarConsolidadoDB() {
    initDB(function(db) {
        let tx = db.transaction(storeName, "readwrite");
        tx.objectStore(storeName).put({ id: 4, consolidado: dadosConsolidados });
    });
}

function carregarConsolidadoDB() {
    initDB(function(db) {
        let tx = db.transaction(storeName, "readonly");
        let req = tx.objectStore(storeName).get(4);
        req.onsuccess = function() {
            if (req.result && req.result.consolidado) dadosConsolidados = req.result.consolidado;
        };
    });
}

let filtrosSelecao = {
    dia: 'TODOS', turno: 'TODOS', periodo: 'TODOS', curso: 'TODOS', disciplina: 'TODOS',
    turma: 'TODOS', prof: 'TODOS', alunos: 'TODOS', sala: 'TODOS', capacidade: 'TODOS'
};

let salas = [
    {name:"C-111", cap:40}, {name:"C-306", cap:40}, {name:"C-301", cap:42}, {name:"C-109", cap:45},
    {name:"C-308", cap:46}, {name:"C-304", cap:47}, {name:"C-219", cap:48}, {name:"C-305", cap:49},
    {name:"C-211", cap:51}, {name:"C-212", cap:53}, {name:"C-218", cap:53}, {name:"C-209", cap:54},
    {name:"C-210", cap:56}, {name:"C-112", cap:58}, {name:"C-208", cap:58}, {name:"E-101", cap:60},
    {name:"E-102", cap:60}, {name:"E-103", cap:60}, {name:"E-104", cap:60}, {name:"C-203", cap:60},
    {name:"G-101", cap:61}, {name:"G-104", cap:61}, {name:"C-204", cap:61}, {name:"C-216", cap:61},
    {name:"G-102", cap:62}, {name:"G-103", cap:63}, {name:"G-107", cap:63}, {name:"C-220", cap:64},
    {name:"C-307", cap:75}, {name:"G-105", cap:80}, {name:"G-106", cap:81}, {name:"C-302", cap:95},
    {name:"C-206", cap:102},{name:"C-105", cap:105},{name:"C-214", cap:120}
];

let blocosCadastrados = [...new Set(salas.map(s => s.name.charAt(0).toUpperCase()))].sort();
salas.forEach(s => { if (!s.tipo) s.tipo = "Sala comum"; });

// =========================================================
// CONFIGURAÇÕES: BLOCOS, SALAS E REGRAS
// =========================================================

function renderizarConfiguracoes() {
    // 1. Sincronizar Checkboxes e Ordem com a Memória
    const r = regrasAlocacao;
    if (document.getElementById('chkManterProf')) {
        document.getElementById('chkManterProf').checked = !!r.manterProf;
        document.getElementById('chkProfSemestre').checked = !!r.manterProfSemestre;
        document.getElementById('chkPrioridadeCalouro').checked = !!r.prioridadeCalouros;
        document.getElementById('chkMargemFolga').checked = !!r.margemFolga;
        document.getElementById('chkPrioridadeTamanho').checked = !!r.prioridadeTamanho;
        document.getElementById('chkTravaSaude').checked = !!r.travaSaude;
        document.getElementById('chkTravaEngArq').checked = !!r.travaEngArq;
        
        document.getElementById('inputOrdemBlocos').value = r.ordemBlocos || "";
        document.getElementById('ordemAtualExibicao').innerText = r.ordemBlocos || "";
    }

    // 2. Renderizar Blocos
    let listaB = document.getElementById('listaBlocos');
    if (listaB) {
        listaB.innerHTML = '';
        blocosCadastrados.forEach(bloco => {
            listaB.innerHTML += `
                <li style="display: flex; justify-content: space-between; align-items: center; padding: 6px; border-bottom: 1px solid #ddd;">
                    <strong>Bloco ${bloco}</strong>
                    <div>
                        <button onclick="editarBloco('${bloco}')" style="background: #f59e0b; padding: 4px 8px; font-size: 10px; margin-right: 5px;">✏️ Editar</button>
                        <button onclick="removerBloco('${bloco}')" style="background: #dc3545; padding: 4px 8px; font-size: 10px;">❌</button>
                    </div>
                </li>
            `;
        });
    }

    let selectBloco = document.getElementById('novaSalaBloco');
    if (selectBloco) {
        let blocoSelecionado = selectBloco.value; 
        selectBloco.innerHTML = '<option value="" disabled>Bloco...</option>';
        blocosCadastrados.forEach(bloco => {
            let selected = (bloco === blocoSelecionado) ? 'selected' : '';
            selectBloco.innerHTML += `<option value="${bloco}" ${selected}>Bloco ${bloco}</option>`;
        });
        if (!blocoSelecionado) selectBloco.selectedIndex = 0;
    }

    // 3. Renderizar Salas
    let listaS = document.getElementById('listaSalasConfig');
    if (listaS) {
        listaS.innerHTML = '';
        let salasOrdenadas = [...salas].sort((a, b) => a.name.localeCompare(b.name));
        salasOrdenadas.forEach(sala => {
            listaS.innerHTML += `
                <li style="display: flex; justify-content: space-between; align-items: center; padding: 6px; border-bottom: 1px solid #ddd;">
                    <span><strong>${sala.name}</strong> (Cap: ${sala.cap}) <br> <em style="color:#888; font-size: 10px;">${sala.tipo}</em></span>
                    <div>
                        <button onclick="editarSala('${sala.name}')" style="background: #f59e0b; padding: 4px 8px; font-size: 10px; margin-right: 5px;">✏️ Editar</button>
                        <button onclick="removerSala('${sala.name}')" style="background: #dc3545; padding: 4px 8px; font-size: 10px;">❌</button>
                    </div>
                </li>
            `;
        });
    }
}

function atualizarRegrasAlocacao() {
    regrasAlocacao.manterProf = document.getElementById('chkManterProf').checked;
    regrasAlocacao.prioridadeCalouros = document.getElementById('chkPrioridadeCalouro').checked;
    regrasAlocacao.margemFolga = document.getElementById('chkMargemFolga').checked;
    regrasAlocacao.manterProfSemestre = document.getElementById('chkProfSemestre').checked;
    regrasAlocacao.prioridadeTamanho = document.getElementById('chkPrioridadeTamanho').checked;
    
    // Captura das novas travas
    if(document.getElementById('chkTravaSaude')) regrasAlocacao.travaSaude = document.getElementById('chkTravaSaude').checked;
    if(document.getElementById('chkTravaEngArq')) regrasAlocacao.travaEngArq = document.getElementById('chkTravaEngArq').checked;
    
    let ordemTemp = document.getElementById('inputOrdemBlocos').value.toUpperCase();
    if(ordemTemp) regrasAlocacao.ordemBlocos = ordemTemp;

    salvarConfigDB();
    executarProcessamento(); 
}

function salvarOrdemBlocos() {
    let inputOrdem = document.getElementById('inputOrdemBlocos').value.toUpperCase();
    let ordemDigitada = inputOrdem.split(',').map(b => b.trim()).filter(b => b);
    let blocosFaltando = blocosCadastrados.filter(b => !ordemDigitada.includes(b));
    
    if (blocosFaltando.length > 0) {
        alert(`⚠️ AÇÃO BLOQUEADA: Você precisa incluir TODOS os blocos existentes na ordem de prioridade.\n\nBlocos esquecidos: ${blocosFaltando.join(', ')}`);
        document.getElementById('inputOrdemBlocos').value = regrasAlocacao.ordemBlocos || blocosCadastrados.join(', ');
        return;
    }

    regrasAlocacao.ordemBlocos = ordemDigitada.join(', ');
    salvarConfigDB();
    executarProcessamento();
    
    let spanOrdemAtual = document.getElementById('ordemAtualExibicao');
    if (spanOrdemAtual) spanOrdemAtual.innerText = regrasAlocacao.ordemBlocos;
    
    alert('✅ Ordem de blocos salva e aplicada com sucesso!');
}

function adicionarBloco() {
    let input = document.getElementById('novoBlocoNome');
    let nome = input.value.trim().toUpperCase();

    if (!nome) { alert("Digite a letra do bloco!"); return; }
    if (blocosCadastrados.includes(nome)) { alert("Este bloco já existe!"); return; }

    blocosCadastrados.push(nome);
    blocosCadastrados.sort();
    input.value = '';
    
    salvarConfigDB();
    renderizarConfiguracoes();
    renderizarMapaSalas(); 
    alert(`✅ Bloco ${nome} Adicionado com sucesso!`);
}

function editarBloco(nomeAntigo) {
    let novoNome = prompt(`Digite a nova letra para o bloco ${nomeAntigo}:`, nomeAntigo);
    if (!novoNome) return;
    novoNome = novoNome.trim().toUpperCase();
    
    if (novoNome === nomeAntigo) return;
    if (blocosCadastrados.includes(novoNome)) { alert("Já existe um bloco com essa letra!"); return; }

    blocosCadastrados = blocosCadastrados.filter(b => b !== nomeAntigo);
    blocosCadastrados.push(novoNome);
    blocosCadastrados.sort();

    salas.forEach(s => {
        if (s.name.startsWith(nomeAntigo)) {
            s.name = novoNome + s.name.substring(nomeAntigo.length);
        }
    });

    salvarConfigDB();
    renderizarConfiguracoes();
    renderizarMapaSalas();
    alert(`✅ Bloco alterado de ${nomeAntigo} para ${novoNome}! As salas associadas foram atualizadas.`);
}

function removerBloco(nome) {
    let salasVinculadas = salas.filter(s => s.name.startsWith(nome));
    if (salasVinculadas.length > 0) {
        alert(`❌ Não é possível remover o Bloco ${nome}. Existem ${salasVinculadas.length} sala(s) dentro dele. Remova ou edite as salas primeiro.`);
        return;
    }

    if (confirm(`Tem certeza que deseja remover o Bloco ${nome}?`)) {
        blocosCadastrados = blocosCadastrados.filter(b => b !== nome);
        salvarConfigDB();
        renderizarConfiguracoes();
        renderizarMapaSalas();
        alert(`🗑️ Bloco ${nome} removido!`);
    }
}

let salaSendoEditada = null;

function editarSala(nomeAntigo) {
    let sala = salas.find(s => s.name === nomeAntigo);
    if (!sala) return;

    let partes = sala.name.split('-');
    let bloco = partes[0] ? partes[0].toUpperCase() : sala.name.charAt(0).toUpperCase();
    let numero = partes[1] ? partes[1] : sala.name.substring(1);

    document.getElementById('novaSalaBloco').value = bloco;
    document.getElementById('novaSalaNumero').value = numero;
    document.getElementById('novaSalaCapacidade').value = sala.cap;
    document.getElementById('novaSalaTipo').value = sala.tipo || "Sala comum";

    salaSendoEditada = nomeAntigo;

    let btnAdd = document.getElementById('btnAdicionarSala');
    let btnCancel = document.getElementById('btnCancelarEdicaoSala');
    
    if (btnAdd) {
        btnAdd.innerHTML = "💾 Salvar Alterações";
        btnAdd.style.background = "#f59e0b";
    }
    if (btnCancel) btnCancel.style.display = "block";

    document.getElementById('novaSalaNumero').focus();
}

function salvarFormularioSala() {
    let bloco = document.getElementById('novaSalaBloco').value;
    let numero = document.getElementById('novaSalaNumero').value.trim().toUpperCase();
    let cap = parseInt(document.getElementById('novaSalaCapacidade').value);
    let tipo = document.getElementById('novaSalaTipo').value;

    if (!bloco) { alert("Selecione um Bloco primeiro!"); return; }
    if (!numero || isNaN(cap) || cap <= 0 || !tipo) { alert("Preencha todos os dados da sala!"); return; }

    let nomeCompleto = bloco + "-" + numero;

    if (salaSendoEditada) {
        if (nomeCompleto !== salaSendoEditada && salas.some(s => s.name === nomeCompleto)) {
            alert(`Já existe outra sala cadastrada como ${nomeCompleto}!`); return;
        }

        let salaObj = salas.find(s => s.name === salaSendoEditada);
        if (salaObj) {
            salaObj.name = nomeCompleto;
            salaObj.cap = cap;
            salaObj.tipo = tipo;
        }

        alert(`✅ Sala ${nomeCompleto} atualizada com sucesso!`);
        cancelarEdicaoSala(); 
    } else {
        if (salas.some(s => s.name === nomeCompleto)) {
            alert(`Já existe uma sala cadastrada como ${nomeCompleto}!`); return;
        }

        salas.push({ name: nomeCompleto, cap: cap, tipo: tipo });
        alert(`✅ Sala ${nomeCompleto} Adicionada com sucesso ao Bloco ${bloco}!`);
        
        document.getElementById('novaSalaNumero').value = '';
        document.getElementById('novaSalaCapacidade').value = '';
        document.getElementById('novaSalaTipo').value = '';
    }

    salvarConfigDB(); 
    renderizarConfiguracoes();
    renderizarMapaSalas(); 
    atualizarFiltrosDinamicos(); 
}

function cancelarEdicaoSala() {
    salaSendoEditada = null;
    
    document.getElementById('novaSalaNumero').value = '';
    document.getElementById('novaSalaCapacidade').value = '';
    document.getElementById('novaSalaTipo').value = '';
    document.getElementById('novaSalaBloco').selectedIndex = 0;

    let btnAdd = document.getElementById('btnAdicionarSala');
    let btnCancel = document.getElementById('btnCancelarEdicaoSala');
    
    if (btnAdd) {
        btnAdd.innerHTML = "Adicionar Sala";
        btnAdd.style.background = "#1fbf75"; 
    }
    if (btnCancel) btnCancel.style.display = "none";
}

function removerSala(nome) {
    if (confirm(`Tem certeza que deseja remover a sala ${nome}?`)) {
        salas = salas.filter(s => s.name !== nome);
        salvarConfigDB();
        renderizarConfiguracoes();
        renderizarMapaSalas();
        atualizarFiltrosDinamicos();
        alert(`🗑️ Sala ${nome} removida!`);
    }
}

// =========================================================
// FUNÇÕES GERAIS DO SISTEMA
// =========================================================

let linhaEditando = null;

function mostrarLoading() {
    document.body.style.cursor = 'wait';
    let overlay = document.getElementById('loadingOverlay');
    if(overlay) overlay.style.display = 'flex';
}

function ocultarLoading() {
    document.body.style.cursor = 'default';
    let overlay = document.getElementById('loadingOverlay');
    if(overlay) overlay.style.display = 'none';
}

function copiarParaAreaDeTransferencia(texto) {
    navigator.clipboard.writeText(texto).then(() => {
    }).catch(err => {
        console.error('Erro ao copiar', err);
    });
}

function buscarLinhaEdicao() {
    let idBusca = document.getElementById('inputIdEdicao').value.trim().toUpperCase();
    linhaEditando = dados.find(d => (d.idUnico || '').toUpperCase() === idBusca);
    
    let painel = document.getElementById('painelEdicao');

    if (linhaEditando) {
        // Cria listas únicas do que já existe na base para o autocompletar (datalist)
        let cursos = [...new Set(dados.map(d => d.curso).filter(Boolean))].sort();
        let profs = [...new Set(dados.map(d => d.prof).filter(Boolean))].sort();
        let disciplinas = [...new Set(dados.map(d => d.nomDisc).filter(Boolean))].sort();

        // Constrói o HTML dinâmico com Inputs + Datalists (permite digitar ou selecionar)
        painel.innerHTML = `
            <p style="font-size:14px; margin-top:0; font-weight:bold; color:#1d7bd8; border-bottom:1px solid #ccc; padding-bottom:5px;">
                Editando ID: ${linhaEditando.idUnico}
            </p>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 15px;">
                <div class="form-group">
                    <label>Curso:</label>
                    <input type="text" id="edCurso" value="${linhaEditando.curso || ''}" list="dlCursos" placeholder="Digite ou selecione">
                    <datalist id="dlCursos">${cursos.map(c => `<option value="${c}">`).join('')}</datalist>
                </div>
                <div class="form-group">
                    <label>Disciplina:</label>
                    <input type="text" id="edDisc" value="${linhaEditando.nomDisc || ''}" list="dlDisc" placeholder="Digite ou selecione">
                    <datalist id="dlDisc">${disciplinas.map(d => `<option value="${d}">`).join('')}</datalist>
                </div>
                <div class="form-group">
                    <label>Turma(s) (Separe por vírgula):</label>
                    <input type="text" id="edTurma" value="${linhaEditando.turma || ''}" placeholder="Ex: 1001, 1002">
                </div>
                <div class="form-group">
                    <label>Professor:</label>
                    <input type="text" id="edProf" value="${linhaEditando.prof || ''}" list="dlProfs" placeholder="Digite ou selecione">
                    <datalist id="dlProfs">${profs.map(p => `<option value="${p}">`).join('')}</datalist>
                </div>
                <div class="form-group">
                    <label>Qtd. Alunos:</label>
                    <input type="number" id="edAlunos" value="${linhaEditando.alunos || 0}">
                </div>
                <div class="form-group">
                    <label>Forçar Sala:</label>
                    <select id="edSala">
                        <option value="">-- MODO AUTOMÁTICO --</option>
                        ${salas.map(s => `<option value="${s.name}" ${linhaEditando.salaManual === s.name ? 'selected' : ''}>${s.name} (Cap: ${s.cap})</option>`).join('')}
                    </select>
                </div>
            </div>
            <button onclick="salvarLinhaEdicao()" style="background:#1fbf75; width:100%; padding:12px;">💾 Salvar Todas as Alterações</button>
        `;
        painel.style.display = 'block';
    } else {
        alert("ID da aula não encontrado na base de dados.");
        painel.style.display = 'none';
    }
}

function salvarLinhaEdicao() {
    if (!linhaEditando) return;
    
    // Coleta todos os valores editados e atualiza o objeto na memória
    linhaEditando.curso = document.getElementById('edCurso').value.trim().toUpperCase();
    linhaEditando.nomDisc = document.getElementById('edDisc').value.trim().toUpperCase();
    linhaEditando.turma = document.getElementById('edTurma').value.trim().toUpperCase();
    linhaEditando.prof = document.getElementById('edProf').value.trim().toUpperCase();
    linhaEditando.alunos = parseInt(document.getElementById('edAlunos').value) || 0;
    linhaEditando.salaManual = document.getElementById('edSala').value;

    // Salva no banco de dados IndexedDB
    salvarDadosDB(dados);
    
    alert("Todas as alterações foram salvas na memória fixa!");
    
    // Limpa a tela e recalcula
    document.getElementById('painelEdicao').style.display = 'none';
    document.getElementById('inputIdEdicao').value = '';
    
    atualizarFiltrosDinamicos(); // Atualiza os filtros superiores caso tenha criado um professor/curso novo
    executarProcessamento();
}



function popularSalasManutencao() {}

function alternarManutencaoMapa(nomeSala) {
    if (salasEmManutencao.includes(nomeSala)) {
        salasEmManutencao = salasEmManutencao.filter(n => n !== nomeSala);
        salasProjetorRuim.push(nomeSala);
    } else if (salasProjetorRuim.includes(nomeSala)) {
        salasProjetorRuim = salasProjetorRuim.filter(n => n !== nomeSala);
    } else {
        salasEmManutencao.push(nomeSala);
    }
    salvarManutencaoDB();
    renderizarMapaSalas(); // <-- ESTA É A LINHA NOVA: Atualiza o mapa instantaneamente na tela
    executarProcessamento(); 
}

function toggleTurmaBloqueada(turmaNome, event) {
    if (event) event.stopPropagation();

    let tNorm = normalizarTexto(turmaNome);
    let itemClicado = event ? event.currentTarget : null;
    
    if (turmasOcultas.includes(tNorm)) {
        turmasOcultas = turmasOcultas.filter(t => t !== tNorm);
        if (itemClicado) {
            itemClicado.style.color = '#334155';
            itemClicado.style.textDecoration = 'none';
            itemClicado.style.fontWeight = 'normal';
            itemClicado.innerText = turmaNome; 
        }
    } else {
        turmasOcultas.push(tNorm);
        if (itemClicado) {
            itemClicado.style.color = '#dc3545';
            itemClicado.style.textDecoration = 'line-through';
            itemClicado.style.fontWeight = 'bold';
            itemClicado.innerText = '❌ Ocultada: ' + turmaNome; 
        }
    }
}

function confirmarBloqueioTurmas(event) {
    if (event) event.stopPropagation();
    
    let body = document.getElementById('body-turmaBloqueada');
    if (body) body.style.display = 'none';
    
    let lblT = document.getElementById('lbl-turmaBloqueada');
    if (lblT) lblT.innerText = turmasOcultas.length > 0 ? `${turmasOcultas.length} Ocultadas` : 'Nenhuma';
    
    salvarManutencaoDB();
    atualizarFiltrosDinamicos();
    executarProcessamento();
}

function extrairNumeroPeriodo(texto) {
    let numStr = String(texto).replace(/[^0-9]/g, '');
    let num = parseInt(numStr);
    return isNaN(num) ? 0 : num;
}

function normalizarTexto(texto) {
    if(!texto) return "";
    return String(texto).trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function classificarTurno(horaIni) {
    if (!horaIni) return "";
    let h = parseInt(horaIni.substring(0, 2));
    if (h < 12) return "Manhã";
    if (h < 18) return "Tarde";
    return "Noite";
}

function ehPrimeiroPeriodo(row) {
    let t = normalizarTexto(row.turma);
    return t.startsWith('10') || t.includes('CALOURO'); 
}

// =========================================================
// O CÉREBRO: ALGORITMO E PROCESSAMENTO
// =========================================================

function getPrioridadeSala(nomeSala, isPrimeiroPeriodo) {
    let bloco = nomeSala.charAt(0).toUpperCase();
    let numStr = nomeSala.replace(/[^0-9]/g, '');
    let numero = numStr ? parseInt(numStr) : 0;

    if (numero >= 300) return 99;

    let ordemStr = regrasAlocacao.ordemBlocos || "E, G, F, C";
    let ordemArray = ordemStr.split(',').map(b => b.trim().toUpperCase());

    let prioridadeBase = ordemArray.indexOf(bloco);
    if (prioridadeBase === -1) prioridadeBase = 50; 

    let desempateAndar = (numero < 200) ? 0 : 0.5;

    return prioridadeBase + desempateAndar; 
}

function temConflitoHorario(intervalosOcupados, horaIniNova, horaFimNova) {
    if (!horaIniNova || !horaFimNova) return false;
    for (let intervalo of intervalosOcupados) {
        if (horaIniNova < intervalo.fim && horaFimNova > intervalo.inicio) return true; 
    }
    return false; 
}

function escolherSalaInteligente(row, mapaDeOcupacao, historicoProfs, historicoProfsGlobal, variar = false) {
    let dia = normalizarTexto(row.diaSemana);
    let qtdAlunos = row.alunos;
    let prof = normalizarTexto(row.prof);
    let isPrimPer = regrasAlocacao.prioridadeCalouros ? ehPrimeiroPeriodo(row) : false;

    // --- NOVAS VARIÁVEIS DE SAÚDE ---
    const cursosSaude = ["PSICOLOGIA", "FISIOTERAPIA", "ODONTOLOGIA", "NUTRICAO", "ENFERMAGEM", "FARMACIA", "BIOMEDICINA", "EDUCACAO FISICA - BACHARELADO"];
    let cursoAtual = normalizarTexto(row.curso);
    let ehCursoSaude = cursosSaude.includes(cursoAtual);
    // --- NOVAS VARIÁVEIS DE ENGENHARIA/ARQUITETURA ---
    const cursosEngArq = ["ARQUITETURA E URBANISMO", "ENGENHARIA CIVIL"];
    let ehCursoEngArq = cursosEngArq.includes(cursoAtual);
    
    if (!mapaDeOcupacao[dia]) mapaDeOcupacao[dia] = {};
    if (!historicoProfs[dia]) historicoProfs[dia] = {};

    // Se o professor faltou, devolve imediatamente o status sem ocupar sala
    if (row.salaManual === "FALTOU") {
        return { name: "FALTOU", cap: 0 };
    }

    if (row.salaManual && row.salaManual !== "") {
        let salaForce = salas.find(s => s.name === row.salaManual);
        if (salaForce) {
            if (!mapaDeOcupacao[dia][salaForce.name]) mapaDeOcupacao[dia][salaForce.name] = [];
            mapaDeOcupacao[dia][salaForce.name].push({ inicio: row.horaIni, fim: row.horaFim });
            return salaForce;
        }
    }

    let salaAnterior = null;
    if (regrasAlocacao.manterProfSemestre && prof !== "") {
        salaAnterior = historicoProfsGlobal[prof];
    } else if (regrasAlocacao.manterProf && prof !== "") {
        salaAnterior = historicoProfs[dia][prof];
    }

    if (salaAnterior && !salasEmManutencao.includes(salaAnterior) && (!variar || Math.random() > 0.5)) {
        let salaObj = salas.find(s => s.name === salaAnterior);
        // --- PARTE 2: TRAVA DE SEGURANÇA PARA LABORATÓRIOS (SAÚDE E ENG) ---
        if (salaObj) {
            if (salaObj.tipo === "Laboratório de Saúde" && !ehCursoSaude) salaObj = null;
            else if (salaObj.tipo === "Laboratório de Engenharia/Arquitetura" && !ehCursoEngArq) salaObj = null;
        }
        if (salaObj && qtdAlunos <= (salaObj.cap * 1.05)) {
            let horarios = mapaDeOcupacao[dia][salaAnterior] || [];
            if (!temConflitoHorario(horarios, row.horaIni, row.horaFim)) {
                mapaDeOcupacao[dia][salaAnterior].push({ inicio: row.horaIni, fim: row.horaFim });
                return salaObj; 
            }
        }
    }

    let pass1Folga = [];   
    let pass2Cheio = [];   
    let pass3Estouro = []; 

    for (let sala of salas) {
        if (salasEmManutencao.includes(sala.name)) continue;

        // --- REGRA 1.1: Restrição Exclusiva do Laboratório de Eng/Arq ---
        if (sala.tipo === "Laboratório de Engenharia/Arquitetura" && !ehCursoEngArq) {
            continue; // Se a sala for Lab de Eng/Arq e o curso não for da área, pula esta sala
        }

        // --- REGRA 1: Restrição Exclusiva do Laboratório ---
        if (sala.tipo === "Laboratório de Saúde" && !ehCursoSaude) {
            continue; // Se a sala for Lab e o curso não for da saúde, ignora esta sala e vai pra próxima
        }

        let horariosDaSala = mapaDeOcupacao[dia][sala.name] || [];
        if (!temConflitoHorario(horariosDaSala, row.horaIni, row.horaFim)) {
            let cap = sala.cap;
            
            if (qtdAlunos <= (cap * 1.05)) {
                let limiteFolga = regrasAlocacao.margemFolga ? (isPrimPer ? 0.80 : 0.95) : 1.0; 
                
                if (qtdAlunos <= (cap * limiteFolga)) pass1Folga.push(sala);
                else if (qtdAlunos <= cap) pass2Cheio.push(sala);
                else pass3Estouro.push(sala);
            }
        }
    }

    let salasDisponiveis = pass1Folga.length > 0 ? pass1Folga : (pass2Cheio.length > 0 ? pass2Cheio : pass3Estouro);

    if (salasDisponiveis.length > 0) {
        salasDisponiveis.sort((a, b) => {

            // --- REGRA 2.1: Preferência Absoluta para Laboratórios de Eng/Arq ---
            if (ehCursoEngArq) {
                let aEhLabEng = a.tipo === "Laboratório de Engenharia/Arquitetura";
                let bEhLabEng = b.tipo === "Laboratório de Engenharia/Arquitetura";
                if (aEhLabEng && !bEhLabEng) return -1; // Sala 'a' (Lab Eng) vai pro topo
                if (!aEhLabEng && bEhLabEng) return 1;  // Sala 'b' (Lab Eng) vai pro topo
            }

            // --- REGRA 2: Preferência Absoluta para Laboratórios ---
            if (ehCursoSaude) {
                let aEhLab = a.tipo === "Laboratório de Saúde";
                let bEhLab = b.tipo === "Laboratório de Saúde";
                if (aEhLab && !bEhLab) return -1; // Sala 'a' (Lab) vai pro topo da lista
                if (!aEhLab && bEhLab) return 1;  // Sala 'b' (Lab) vai pro topo da lista
            }
            let prioA = getPrioridadeSala(a.name, isPrimPer);
            let prioB = getPrioridadeSala(b.name, isPrimPer);
            
            if (variar) {
                prioA += Math.random() * 0.5;
                prioB += Math.random() * 0.5;
            }

            if (prioA !== prioB) return prioA - prioB;
            return a.cap - b.cap;
        });

        let salaEscolhida = salasDisponiveis[0];
        
        if (!mapaDeOcupacao[dia][salaEscolhida.name]) mapaDeOcupacao[dia][salaEscolhida.name] = [];
        mapaDeOcupacao[dia][salaEscolhida.name].push({ inicio: row.horaIni, fim: row.horaFim });
        
        if (prof !== "") {
            historicoProfs[dia][prof] = salaEscolhida.name;
            historicoProfsGlobal[prof] = salaEscolhida.name;
        }
        
        return salaEscolhida;
    }

    return {name: "SEM SALA", cap: 0};
}

function processar(variar = false) {
    if (dados.length === 0) return;

    let dIni = document.getElementById("dataInicio").value;
    let dFim = document.getElementById("dataFim").value;
    let fs = filtrosSelecao;

    let dadosParaProcessar = dados.filter(row => {
        if (dIni && dFim && row.dataIni) {
            let dataNormal = converterDataComparacao(row.dataIni);
            if (dataNormal < dIni || dataNormal > dFim) return false;
        }
        if (fs.dia !== 'TODOS' && normalizarTexto(row.diaSemana) !== normalizarTexto(fs.dia)) return false;
        if (fs.turno !== 'TODOS') {
            let tClassificado = normalizarTexto(classificarTurno(row.horaIni));
            if (fs.turno === 'MANHÃ/TARDE') {
                if (tClassificado !== 'MANHA' && tClassificado !== 'TARDE') return false;
            } else {
                if (tClassificado !== normalizarTexto(fs.turno)) return false;
            }
        }

        if (fs.periodo !== 'TODOS' && normalizarTexto(row.periodoAcad) !== normalizarTexto(fs.periodo)) return false;
        if (fs.curso !== 'TODOS' && normalizarTexto(row.curso) !== normalizarTexto(fs.curso)) return false;
        let mat = row.codDisc ? `${row.codDisc.trim()} - ${row.nomDisc.trim()}` : row.nomDisc.trim();
        if (fs.disciplina !== 'TODOS' && normalizarTexto(mat) !== normalizarTexto(fs.disciplina)) return false;
        if (fs.prof !== 'TODOS' && normalizarTexto(row.prof) !== normalizarTexto(fs.prof)) return false;
        if (fs.alunos !== 'TODOS' && String(row.alunos) !== fs.alunos) return false;
        
        if (turmasOcultas.includes(normalizarTexto(row.turma))) return false;
        
        return true;
    });

    let todosOrdem = [...dadosParaProcessar];
    todosOrdem.sort((a, b) => {
        if (a.diaSemana !== b.diaSemana) return a.diaSemana.localeCompare(b.diaSemana);
        if (a.horaIni !== b.horaIni) return a.horaIni.localeCompare(b.horaIni);
        
        // SE A REGRA ESTIVER ATIVA: Turmas com mais alunos escolhem a sala primeiro no mesmo horário!
        if (regrasAlocacao.prioridadeTamanho) return b.alunos - a.alunos;
        
        return 0;
    });

    let ocupacaoMap = {}; 
    let historicoProfs = {}; 
    let historicoProfsGlobal = {};
    
    windowDadosResultantes = [];
    
    for (let i = 0; i < todosOrdem.length; i++) {
        let row = todosOrdem[i];
        let sala = escolherSalaInteligente(row, ocupacaoMap, historicoProfs, historicoProfsGlobal, variar);
        let perc = sala.cap ? parseFloat(((row.alunos / sala.cap) * 100).toFixed(1)) : 0;
        
        windowDadosResultantes.push({
            ...row,
            salaSugerida: sala.name,
            capacidade: sala.cap,
            percUso: perc
        });
    }

    aplicarFiltrosTabela(); 
}

function executarProcessamento() {
    mostrarLoading(); 
    
    let sp1 = document.getElementById('spinnerGrafico1');
    let sp2 = document.getElementById('spinnerGrafico2');
    if(sp1) sp1.style.display = 'flex';
    if(sp2) sp2.style.display = 'flex';

    setTimeout(() => {
        processar(false); 
        ocultarLoading();
    }, 600);

    setTimeout(() => {
        atualizarGraficos(); 
        if(sp1) sp1.style.display = 'none';
        if(sp2) sp2.style.display = 'none';
    }, 2400);
}

function executarRecalculo() {
    mostrarLoading(); 
    let sp1 = document.getElementById('spinnerGrafico1');
    let sp2 = document.getElementById('spinnerGrafico2');
    if(sp1) sp1.style.display = 'flex';
    if(sp2) sp2.style.display = 'flex';

    setTimeout(() => {
        processar(true); 
        ocultarLoading();
    }, 500);

    setTimeout(() => {
        atualizarGraficos(); 
        if(sp1) sp1.style.display = 'none';
        if(sp2) sp2.style.display = 'none';
    }, 2000);
}

// =========================================================
// FILTROS E EXIBIÇÃO DE TABELA
// =========================================================

function construirFiltrosAvancados() {
    let container = document.getElementById('containerHubBase');
    if (!container) return;

    container.innerHTML = '';
    container.style.display = 'flex';
    container.style.flexDirection = 'column';
    container.style.gap = '15px';

    // LINHA 1: Filtros Visíveis e Botão Processar
    let divFiltros = document.createElement('div');
    divFiltros.style.display = 'flex';
    divFiltros.style.gap = '15px';
    divFiltros.style.alignItems = 'center';
    divFiltros.style.width = '100%';
    divFiltros.style.flexWrap = 'wrap';

    let grpData = document.createElement('div');
    grpData.className = 'input-group';
    grpData.innerHTML = `
        <label style="font-size:12px;">Dia da Semana:</label>
        <div class="custom-dd" id="dd-dia">
            <div class="custom-dd-header" onclick="alternarListaSuspensa('dia')"><span id="lbl-dia">TODOS</span> <span>▼</span></div>
            <div class="custom-dd-body" id="body-dia">
                <input type="text" placeholder="Pesquisar..." onkeyup="filtrarListaSuspensa('dia', this.value)">
                <div class="custom-dd-list" id="list-dia"></div>
            </div>
        </div>
        <input type="hidden" id="dataInicio" value="">
        <input type="hidden" id="dataFim" value="">
    `;
    divFiltros.appendChild(grpData);

    let grpTurno = document.createElement('div');
    grpTurno.className = 'input-group';
    grpTurno.innerHTML = `
        <label style="font-size:12px;">Turno:</label>
        <div class="custom-dd" id="dd-turno">
            <div class="custom-dd-header" onclick="alternarListaSuspensa('turno')"><span id="lbl-turno">Todos</span> <span>▼</span></div>
            <div class="custom-dd-body" id="body-turno">
                <input type="text" placeholder="Pesquisar..." onkeyup="filtrarListaSuspensa('turno', this.value)">
                <div class="custom-dd-list" id="list-turno"></div>
            </div>
        </div>
    `;
    divFiltros.appendChild(grpTurno);

    let btnProcessar = document.createElement('button');
    btnProcessar.id = 'btnProcessar';
    btnProcessar.innerHTML = '▶️ Processar Dados';
    btnProcessar.style.background = '#1d7bd8'; // Força cor base
    btnProcessar.onclick = executarProcessamento; 
    divFiltros.appendChild(btnProcessar);
    
    let contadorAulas = document.createElement('div');
    contadorAulas.id = 'contadorAulasFound';
    contadorAulas.style.cssText = 'background: #e2e8f0; color: #1e293b; padding: 8px 15px; border-radius: 6px; font-weight: bold; font-size: 13px; border: 1px solid #cbd5e1; margin-left: 5px;';
    contadorAulas.innerHTML = '0 AULAS ENCONTRADAS';
    divFiltros.appendChild(contadorAulas);

    // LINHA 2: Botões de Ação
    let divBotoes = document.createElement('div');
    divBotoes.style.display = 'flex';
    divBotoes.style.gap = '10px';
    divBotoes.style.alignItems = 'center';
    divBotoes.style.width = '100%';
    divBotoes.style.flexWrap = 'wrap';
    divBotoes.style.borderTop = '1px solid #e0e0e0';
    divBotoes.style.paddingTop = '15px';

    let btnRecalcular = document.createElement('button');
    btnRecalcular.innerText = '🔄 Nova Opção';
    btnRecalcular.style.background = '#8b5cf6';
    btnRecalcular.onclick = executarRecalculo;

    let btnExportar = document.createElement('button');
    btnExportar.innerText = '⬇️ CSV';
    btnExportar.style.background = '#10b981';
    btnExportar.onclick = exportarCSV;

    let btnLimpar = document.createElement('button');
    btnLimpar.innerText = '🗑️ Limpar';
    btnLimpar.style.background = '#dc3545';
    btnLimpar.onclick = limparMemoria;

    let btnMaisFiltros = document.createElement('button');
    btnMaisFiltros.id = 'btnMaisFiltros';
    btnMaisFiltros.innerText = '🔽 Mais Filtros';
    btnMaisFiltros.style.background = 'transparent';
    btnMaisFiltros.style.color = '#1d7bd8';
    btnMaisFiltros.style.border = '1px solid #1d7bd8';
    btnMaisFiltros.onclick = toggleMaisFiltros;

    divBotoes.appendChild(btnRecalcular);
    divBotoes.appendChild(btnExportar);
    divBotoes.appendChild(btnLimpar);
    divBotoes.appendChild(btnMaisFiltros);

    // LINHA 3: Filtros Ocultos
    let divOcultos = document.createElement('div');
    divOcultos.id = 'filtrosOcultosContainer';

    const camposOcultos = [
        { id: 'periodo', label: 'Período:' }, 
        { id: 'curso', label: 'Curso:' },
        { id: 'disciplina', label: 'Matéria:' },
        { id: 'prof', label: 'Professor:' },
        { id: 'alunos', label: 'Alunos:' },
        { id: 'sala', label: 'Sala Sug.:' },
        { id: 'capacidade', label: 'Capac.:' }
    ];

    camposOcultos.forEach(f => {
        let div = document.createElement('div');
        div.className = 'input-group';
        div.innerHTML = `
            <label style="font-size:12px;">${f.label}</label>
            <div class="custom-dd" id="dd-${f.id}">
                <div class="custom-dd-header" onclick="alternarListaSuspensa('${f.id}')"><span id="lbl-${f.id}">Todos</span> <span>▼</span></div>
                <div class="custom-dd-body" id="body-${f.id}">
                    <input type="text" placeholder="Pesquisar..." onkeyup="filtrarListaSuspensa('${f.id}', this.value)">
                    <div class="custom-dd-list" id="list-${f.id}"></div>
                </div>
            </div>
        `;
        divOcultos.appendChild(div);
    });

    let divTurmaBloq = document.createElement('div');
    divTurmaBloq.className = 'input-group';
    divTurmaBloq.innerHTML = `
        <label style="font-size:12px; color:#dc3545;">Turma Bloqueada:</label>
        <div class="custom-dd" id="dd-turmaBloqueada">
            <div class="custom-dd-header" onclick="alternarListaSuspensa('turmaBloqueada')"><span id="lbl-turmaBloqueada">Nenhuma</span> <span>▼</span></div>
            <div class="custom-dd-body" id="body-turmaBloqueada">
                <input type="text" placeholder="Pesquisar..." onkeyup="filtrarListaSuspensa('turmaBloqueada', this.value)">
                <div class="custom-dd-list" id="list-turmaBloqueada"></div>
            </div>
        </div>
    `;
    divOcultos.appendChild(divTurmaBloq);

    // ADICIONA TUDO AO CONTAINER
    container.appendChild(divFiltros);
    container.appendChild(divBotoes);
    container.appendChild(divOcultos);

    document.getElementById('dataInicio').onchange = () => { 
        document.getElementById("dataFim").value = document.getElementById("dataInicio").value;
        sincronizarDataComDia(); 
    };

    document.addEventListener('click', (e) => {
        if (!e.target.closest('.custom-dd')) {
            document.querySelectorAll('.custom-dd-body').forEach(b => b.style.display = 'none');
        }
    });

    prepararCabecalhoTabela();
}

function toggleMaisFiltros() {
    let container = document.getElementById('filtrosOcultosContainer');
    let btn = document.getElementById('btnMaisFiltros');
    if (container.classList.contains('expandido')) {
        container.classList.remove('expandido');
        btn.innerText = '🔽 Mais Filtros';
    } else {
        container.classList.add('expandido');
        btn.innerText = '🔼 Menos Filtros';
    }
}

function sincronizarDataComDia() {
    let dIni = document.getElementById("dataInicio").value;
    if (dIni) {
        let p = dIni.split('-');
        let dateObj = new Date(p[0], p[1] - 1, p[2]);
        let diasVisuais = ['DOMINGO', 'SEGUNDA', 'TERÇA', 'QUARTA', 'QUINTA', 'SEXTA', 'SÁBADO'];
        let diasReais = ['DOMINGO', 'SEGUNDA', 'TERCA', 'QUARTA', 'QUINTA', 'SEXTA', 'SABADO'];
        let idx = dateObj.getDay();
        selecionarOpcao('dia', diasReais[idx], diasVisuais[idx]);
    }
}

function alternarListaSuspensa(id) {
    let body = document.getElementById(`body-${id}`);
    let estaVisivel = body.style.display === 'block';
    document.querySelectorAll('.custom-dd-body').forEach(b => b.style.display = 'none');
    if (!estaVisivel) body.style.display = 'block';
}

function filtrarListaSuspensa(id, termo) {
    let pesquisa = normalizarTexto(termo);
    let itens = document.querySelectorAll(`#list-${id} .custom-dd-option`);
    itens.forEach(item => {
        if (item.innerText === 'Todos' || normalizarTexto(item.innerText).includes(pesquisa)) {
            item.style.display = 'block';
        } else {
            item.style.display = 'none';
        }
    });
}

function selecionarOpcao(id, valorReal, textoVisivel) {
    filtrosSelecao[id] = valorReal;
    let lbl = document.getElementById(`lbl-${id}`);
    if(lbl) lbl.innerText = textoVisivel;
    
    let body = document.getElementById(`body-${id}`);
    if(body) body.style.display = 'none';
    
    let inp = document.querySelector(`#body-${id} input`);
    if(inp) inp.value = "";
    
    filtrarListaSuspensa(id, "");
}

function atualizarFiltrosDinamicos() {
    let sets = { dia: new Set(), periodo: new Set(), curso: new Set(), disciplina: new Set(), turma: new Set(), prof: new Set(), alunos: new Set() };
    
    sets.dia.add('SEGUNDA'); sets.dia.add('TERÇA'); sets.dia.add('QUARTA'); sets.dia.add('QUINTA'); sets.dia.add('SEXTA'); sets.dia.add('SÁBADO');

    dados.forEach(d => {
        if(d.diaSemana) sets.dia.add(d.diaSemana.toUpperCase());
        if(d.curso) sets.curso.add(d.curso.toUpperCase());
        if(d.nomDisc) {
            let m = d.codDisc ? `${d.codDisc.trim()} - ${d.nomDisc.trim()}` : d.nomDisc.trim();
            sets.disciplina.add(m.toUpperCase());
        }
        if(d.turma) sets.turma.add(d.turma.toUpperCase());
        if(d.prof) sets.prof.add(d.prof.toUpperCase());
        if(d.alunos) sets.alunos.add(String(d.alunos));
        if(d.periodoAcad) sets.periodo.add(d.periodoAcad.toUpperCase());
    });

    let salasSet = new Set(), capSet = new Set();
    salas.forEach(s => { salasSet.add(s.name); capSet.add(String(s.cap)); });
    salasSet.add('SEM SALA');

    function popularLista(id, conj) {
        let div = document.getElementById(`list-${id}`);
        if (!div) return;
        let html = `<div class="custom-dd-option" onclick="selecionarOpcao('${id}', 'TODOS', 'Todos')"><strong>Todos</strong></div>`;
        Array.from(conj).sort().forEach(val => {
            if (val && val.trim() !== '') {
                let vEscape = val.replace(/'/g, "\\'");
                html += `<div class="custom-dd-option" onclick="selecionarOpcao('${id}', '${vEscape}', '${vEscape}')">${val}</div>`;
            }
        });
        div.innerHTML = html;
        filtrosSelecao[id] = 'TODOS';
        let lbl = document.getElementById(`lbl-${id}`);
        if(lbl) lbl.innerText = 'Todos';
    }

    let divDia = document.getElementById('list-dia');
    if (divDia) {
        let htmlDia = `<div class="custom-dd-option" onclick="selecionarOpcao('dia', 'TODOS', 'Todos')"><strong>Todos</strong></div>`;
        
        let diasNomes = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
        let hojeIdx = new Date().getDay();
        let nomeHoje = diasNomes[hojeIdx];
        let valHoje = nomeHoje.toUpperCase(); 
        
        htmlDia += `<div class="custom-dd-option" style="color: #1d7bd8; font-weight: bold; background: #f0f8ff; border-radius: 4px; margin-bottom: 5px;" onclick="selecionarOpcao('dia', '${valHoje}', 'Hoje (${nomeHoje})')">📅 Hoje (${nomeHoje})</div>`;
        
        let ordemCorreta = ["SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SÁBADO", "DOMINGO"];
        ordemCorreta.forEach(diaStr => {
            if (sets.dia.has(diaStr)) {
                let nomeVisual = diaStr.charAt(0) + diaStr.slice(1).toLowerCase();
                htmlDia += `<div class="custom-dd-option" onclick="selecionarOpcao('dia', '${diaStr}', '${nomeVisual}')">${nomeVisual}</div>`;
            }
        });
        
        divDia.innerHTML = htmlDia;
        filtrosSelecao['dia'] = 'TODOS';
        let lbl = document.getElementById('lbl-dia');
        if(lbl) lbl.innerText = 'Todos';
    }
    popularLista('periodo', sets.periodo);
    popularLista('curso', sets.curso);
    popularLista('disciplina', sets.disciplina);
    popularLista('prof', sets.prof);
    popularLista('alunos', sets.alunos);
    popularLista('sala', salasSet);
    popularLista('capacidade', capSet);

    let divTurmaB = document.getElementById('list-turmaBloqueada');
    if (divTurmaB) {
        let htmlT = `
        <div style="display: flex; gap: 5px; padding: 5px; border-bottom: 2px solid #eee; position: sticky; top: 0; background: white; z-index: 2;">
            <button onclick="turmasOcultas=[]; salvarManutencaoDB(); atualizarFiltrosDinamicos(); executarProcessamento();" style="flex: 1; padding: 6px; font-size: 11px; background: #e2e8f0; color: #475569; border-radius: 4px; border: 1px solid #cbd5e1; cursor: pointer;">🔄 Limpar</button>
            <button onclick="confirmarBloqueioTurmas(event)" style="flex: 1; padding: 6px; font-size: 11px; background: #1d7bd8; color: white; border-radius: 4px; border: none; cursor: pointer;">✅ OK</button>
        </div>`;

        Array.from(sets.turma).sort().forEach(val => {
            if (val && val.trim() !== '') {
                let vEscape = val.replace(/'/g, "\\'");
                let isBlocked = turmasOcultas.includes(normalizarTexto(val));
                
                let css = isBlocked ? "color: #dc3545; text-decoration: line-through; font-weight: bold;" : "color: #334155;";
                let txt = isBlocked ? `❌ ${val}` : val;
                
                htmlT += `<div class="custom-dd-option" style="${css}" onclick="toggleTurmaBloqueada('${vEscape}', event)">${txt}</div>`;
            }
        });
        divTurmaB.innerHTML = htmlT;
        
        let lblT = document.getElementById('lbl-turmaBloqueada');
        if (lblT) lblT.innerText = turmasOcultas.length > 0 ? `${turmasOcultas.length} Ocultadas` : 'Nenhuma';
    }


    let divTurno = document.getElementById(`list-turno`);
    if(divTurno) {
        divTurno.innerHTML = `
            <div class="custom-dd-option" onclick="selecionarOpcao('turno', 'TODOS', 'Todos')"><strong>Todos</strong></div>
            <div class="custom-dd-option" onclick="selecionarOpcao('turno', 'MANHÃ', 'Manhã')">Manhã</div>
            <div class="custom-dd-option" onclick="selecionarOpcao('turno', 'TARDE', 'Tarde')">Tarde</div>
            <div class="custom-dd-option" onclick="selecionarOpcao('turno', 'NOITE', 'Noite')">Noite</div>
            <div class="custom-dd-option" onclick="selecionarOpcao('turno', 'MANHÃ/TARDE', 'Manhã/Tarde')">Manhã/Tarde</div>
        `;
        filtrosSelecao['turno'] = 'TODOS';
        let lblTurno = document.getElementById(`lbl-turno`);
        if (lblTurno) lblTurno.innerText = 'Todos';
    }
}

function prepararCabecalhoTabela() {
    // Busca exclusivamente o cabeçalho da tabela principal (ignorando a tabela do Consolidado)
    let theadTr = document.getElementById("resultado").closest("table").querySelector("thead tr");
    if (theadTr && !document.getElementById('sortPercUso')) {
        
        if (!document.getElementById('th-id')) {
            let thId = document.createElement("th");
            thId.id = "th-id";
            thId.innerHTML = `<button onclick="abrirModalNovaAula()" style="background:#10b981; color:white; border:none; border-radius:4px; width:24px; height:24px; cursor:pointer; padding:0; display:flex; align-items:center; justify-content:center; box-shadow: 0 2px 4px rgba(16,185,129,0.3);" title="Adicionar Nova Aula">
    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>
</button>`;
            thId.style.width = "30px";
            thId.style.textAlign = "center";
            theadTr.insertBefore(thId, theadTr.firstChild);
        }

        let ths = theadTr.querySelectorAll("th");
        
        function injetarFiltroOrdenacaoNum(th, id, labelBase) {
            th.innerHTML = `
                <select id="${id}" style="background: transparent; color: inherit; font-weight: 600; border: none; font-size: 11px; cursor: pointer; outline: none; text-transform: uppercase; width:100%;" onchange="aplicarFiltrosTabela('${id}')">
                    <option value="" style="color:#334155;">${labelBase} ↕</option>
                    <option value="asc" style="color:#334155;">MENOR ↑</option>
                    <option value="desc" style="color:#334155;">MAIOR ↓</option>
                </select>
            `;
        }
        
        injetarFiltroOrdenacaoNum(ths[3], 'sortPeriodo', 'PERÍODO');
        injetarFiltroOrdenacaoNum(ths[7], 'sortAlunos', 'ALUNOS');
        injetarFiltroOrdenacaoNum(ths[9], 'sortCapacidade', 'CAPAC.');
        injetarFiltroOrdenacaoNum(ths[10], 'sortPercUso', '% USO');
    }
}

function aplicarFiltrosTabela(fonteModificada) {
    if (!windowDadosResultantes || windowDadosResultantes.length === 0) return;
    
    if (fonteModificada) {
        ['sortPeriodo', 'sortAlunos', 'sortCapacidade', 'sortPercUso'].forEach(id => {
            if (id !== fonteModificada && document.getElementById(id)) {
                document.getElementById(id).value = "";
            }
        });
    }

    let valPeriodo = document.getElementById('sortPeriodo') ? document.getElementById('sortPeriodo').value : "";
    let valAlunos = document.getElementById('sortAlunos') ? document.getElementById('sortAlunos').value : "";
    let valCap = document.getElementById('sortCapacidade') ? document.getElementById('sortCapacidade').value : "";
    let valPerc = document.getElementById('sortPercUso') ? document.getElementById('sortPercUso').value : "";
    
    let fs = filtrosSelecao;
    
    let filtrado = windowDadosResultantes.filter(row => {
        if (salasSelecionadas.length > 0 && !salasSelecionadas.includes(normalizarTexto(row.salaSugerida))) return false;
        if (fs.capacidade !== 'TODOS' && String(row.capacidade) !== fs.capacidade) return false;
        return true;
    });

    if (valPeriodo === 'asc') filtrado.sort((a, b) => extrairNumeroPeriodo(a.periodoAcad) - extrairNumeroPeriodo(b.periodoAcad));
    else if (valPeriodo === 'desc') filtrado.sort((a, b) => extrairNumeroPeriodo(b.periodoAcad) - extrairNumeroPeriodo(a.periodoAcad));
    else if (valAlunos === 'asc') filtrado.sort((a, b) => a.alunos - b.alunos);
    else if (valAlunos === 'desc') filtrado.sort((a, b) => b.alunos - a.alunos);
    else if (valCap === 'asc') filtrado.sort((a, b) => a.capacidade - b.capacidade);
    else if (valCap === 'desc') filtrado.sort((a, b) => b.capacidade - a.capacidade);
    else if (valPerc === 'asc') filtrado.sort((a, b) => a.percUso - b.percUso);
    else if (valPerc === 'desc') filtrado.sort((a, b) => b.percUso - a.percUso);

    let contador = document.getElementById('contadorAulasFound');
    if (contador) contador.innerHTML = `${filtrado.length} AULAS ENCONTRADAS`;

    renderizarTabela(filtrado);
    renderizarMapaSalas();
}

function renderizarTabela(dadosParaRenderizar) {
    let tbody = document.getElementById("resultado");
    tbody.innerHTML = "";
    
    renderId++; 
    let myRenderId = renderId; 
    let index = 0;
    let tamanhoLote = 500; 

    function cascata() {
        if (myRenderId !== renderId) return; 

        let fragment = document.createDocumentFragment();
        let fimLote = Math.min(index + tamanhoLote, dadosParaRenderizar.length);

        for (; index < fimLote; index++) {
            let row = dadosParaRenderizar[index];
            let tr = document.createElement("tr");
            let classeSemSala = row.salaSugerida === "SEM SALA" ? "warning" : "";
            
            // Pinta a linha inteira de vermelho claro se faltou
            if (row.salaSugerida === "FALTOU") {
                tr.style.backgroundColor = "#fee2e2"; 
            }
            
            let horarioExibicao = row.horaIni;
            if(horarioExibicao.length > 5) horarioExibicao = horarioExibicao.substring(0, 5);

            let datasStr = "-";
            if (row.dataIni && row.dataFimAloc) {
                datasStr = `<br><span style="font-size:10px; color:#64748b; font-weight:normal;">${formatarDataVisual(row.dataIni)} a ${formatarDataVisual(row.dataFimAloc)}</span>`;
            } else if (row.dataIni) {
                datasStr = `<br><span style="font-size:10px; color:#64748b; font-weight:normal;">A partir de ${formatarDataVisual(row.dataIni)}</span>`;
            }

            let nomeTurno = classificarTurno(row.horaIni);
            let turnoStr = `<span style="font-size:10px; color:#64748b; font-weight:normal; text-transform:uppercase;">${nomeTurno}</span><br>`;

            let widthBar = Math.min(row.percUso, 100);
            let styleBarraVermelha = row.percUso > 100 ? 'background: linear-gradient(90deg, #ef4444, #b91c1c);' : '';

            let avisoProjetor = salasProjetorRuim.includes(row.salaSugerida) ? `<span style="font-size:10px; color:#d97706; font-weight:bold; display:block; line-height:1;">⚠️ Proj. Ruim</span>` : '';
            let avisoForcado = row.salaManual ? `<span style="font-size:10px; color:#1d7bd8; font-weight:bold; display:block; line-height:1;">🔒 Forçada</span>` : '';

           let discUnificadaStr = `
                <div style="font-weight: 600; color: #333; font-size: 11px;">
                    ${row.codDisc || ''} ${row.turma || ''} ${row.nomDisc || ''}
                </div>
            `;

            tr.innerHTML = `
                <td style="text-align: center;">
                    <button title="Editar Aula" onclick="abrirEdicaoDireta('${row.idUnico}')" style="background:#f59e0b; color:white; border:none; border-radius:4px; padding:4px 6px; cursor:pointer; display:inline-flex; align-items:center; justify-content:center; box-shadow: 0 2px 4px rgba(245,158,11,0.3); transition: transform 0.1s;" onmouseover="this.style.transform='scale(1.05)'" onmouseout="this.style.transform='scale(1)'">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path></svg>
                    </button>
                </td>
                <td><strong>${row.diaSemana}</strong>${datasStr}</td>
                <td>${turnoStr}${horarioExibicao}</td>
                <td style="font-weight: 600; color: #475569;">${row.periodoAcad || '-'}</td>
                <td>${row.curso}</td>
                <td>${discUnificadaStr}</td>
                <td>${row.prof}</td>
                <td>${row.alunos}</td>
                <td class="${classeSemSala}">${avisoForcado}${avisoProjetor}${row.salaSugerida}</td>
                <td>${row.capacidade}</td>
                <td>
                    <div class="progress">
                        <div class="progress-bar" style="width:${widthBar}%; ${styleBarraVermelha}"></div>
                    </div>
                    ${row.percUso}%
                </td>
            `;
            fragment.appendChild(tr);
        }

        tbody.appendChild(fragment);

        if (index < dadosParaRenderizar.length) {
            requestAnimationFrame(cascata);
        }
    }
    cascata();
}

function atualizarGraficos() {
    let base = windowDadosResultantes.length > 0 ? windowDadosResultantes : dados;
    if (base.length === 0) return;

    let dIni = document.getElementById("dataInicio").value;
    let dFim = document.getElementById("dataFim").value;
    let fs = filtrosSelecao;

    let filtrado = base.filter(row => {
        if (dIni && dFim && row.dataIni) {
            let dataA = converterDataComparacao(row.dataIni);
            if (dataA < dIni || dataA > dFim) return false;
        }
        if (fs.dia !== 'TODOS' && normalizarTexto(row.diaSemana) !== normalizarTexto(fs.dia)) return false;
        
        if (fs.turno !== 'TODOS') {
            let tClass = normalizarTexto(classificarTurno(row.horaIni));
            if (fs.turno === 'MANHÃ/NOITE') {
                if (tClass !== 'MANHA' && tClass !== 'NOITE') return false;
            } else {
                if (tClass !== normalizarTexto(fs.turno)) return false;
            }
        }

        if (fs.periodo !== 'TODOS' && normalizarTexto(row.periodoAcad) !== normalizarTexto(fs.periodo)) return false;
        if (fs.curso !== 'TODOS' && normalizarTexto(row.curso) !== normalizarTexto(fs.curso)) return false;
        let mat = row.codDisc ? `${row.codDisc.trim()} - ${row.nomDisc.trim()}` : row.nomDisc.trim();
        if (fs.disciplina !== 'TODOS' && normalizarTexto(mat) !== normalizarTexto(fs.disciplina)) return false;
        if (fs.prof !== 'TODOS' && normalizarTexto(row.prof) !== normalizarTexto(fs.prof)) return false;
        if (salasSelecionadas.length > 0 && (!row.salaSugerida || !salasSelecionadas.includes(normalizarTexto(row.salaSugerida)))) return false;
        if (fs.capacidade !== 'TODOS' && row.capacidade && String(row.capacidade) !== fs.capacidade) return false;
        
        return true;
    });

    let alunosPorDia = { "SEGUNDA": 0, "TERCA": 0, "QUARTA": 0, "QUINTA": 0, "SEXTA": 0, "SABADO": 0 };
    let alunosPorCurso = {};

    filtrado.forEach(d => {
        let diaNormal = normalizarTexto(d.diaSemana);
        if (alunosPorDia[diaNormal] !== undefined) alunosPorDia[diaNormal] += d.alunos;
        let cNormal = d.curso ? d.curso : "OUTROS";
        alunosPorCurso[cNormal] = (alunosPorCurso[cNormal] || 0) + d.alunos;
    });

    let dadosDias = [alunosPorDia["SEGUNDA"], alunosPorDia["TERCA"], alunosPorDia["QUARTA"], alunosPorDia["QUINTA"], alunosPorDia["SEXTA"], alunosPorDia["SABADO"]];
    let cursosOrdenados = Object.entries(alunosPorCurso).sort((a, b) => b[1] - a[1]).slice(0, 5);
    let labelsCursos = cursosOrdenados.map(c => c[0]);
    let dadosCursos = cursosOrdenados.map(c => c[1]);

    if(labelsCursos.length === 0) { labelsCursos = ["Sem Dados"]; dadosCursos = [0]; }

    if (chartDiasInstancia) {
        chartDiasInstancia.data.datasets[0].data = dadosDias;
        chartDiasInstancia.update();
    } else {
        let ctxDias = document.getElementById('graficoDias').getContext('2d');
        chartDiasInstancia = new Chart(ctxDias, {
            type: 'bar',
            data: { labels: ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado"], datasets: [{ label: 'Total de Alunos', data: dadosDias, backgroundColor: '#1d7bd8', borderRadius: 4 }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { title: { display: true, text: 'Volume de Alunos por Dia' } } }
        });
    }

    if (chartCursosInstancia) {
        chartCursosInstancia.data.labels = labelsCursos;
        chartCursosInstancia.data.datasets[0].data = dadosCursos;
        chartCursosInstancia.update(); 
    } else {
        let ctxCursos = document.getElementById('graficoCursos').getContext('2d');
        chartCursosInstancia = new Chart(ctxCursos, {
            type: 'doughnut',
            data: { labels: labelsCursos, datasets: [{ data: dadosCursos, backgroundColor: ['#1d7bd8', '#19b5c8', '#1fbf75', '#f6a821', '#e53935'] }] },
            options: { responsive: true, maintainAspectRatio: false, plugins: { title: { display: true, text: 'Top 5 Cursos ' } } }
        });
    }
}

// =========================================================
// FUNÇÕES DE ARQUIVO E EXPORTAÇÃO
// =========================================================

function exportarCSV() {
    if (!windowDadosResultantes || windowDadosResultantes.length === 0) {
        alert("Não há dados processados para exportar. Clique em 'Processar Dados' primeiro.");
        return;
    }

    const headers = ["dia", "Curso", "Horário", "Período", "Disciplina", "Professor(a)", "Salas"];
    let csvContent = "\uFEFF"; 
    csvContent += headers.join(";") + "\n";

    windowDadosResultantes.forEach(row => {
        const dia = row.diaSemana || "";
        const curso = row.curso || "";
        const horario = `${(row.horaIni || "").substring(0, 5)} às ${(row.horaFim || "").substring(0, 5)}`;
        const periodo = row.periodoAcad || "";
        const disciplina = row.nomDisc || "";
        const professor = row.prof || "";
        const sala = row.salaSugerida.replace("🔒 Forçada", "").replace("⚠️ Proj. Ruim", "").trim();

        const linhaMapeada = [
            `"${dia}"`, `"${curso}"`, `"${horario}"`, `"${periodo}"`, `"${disciplina}"`, `"${professor}"`, `"${sala}"`
        ];
        csvContent += linhaMapeada.join(";") + "\n";
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", "planejamento_salas_estacio.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function importarCopiado() {
    let texto = document.getElementById("textoExcel").value.trim();
    if (!texto) { alert("Cole os dados da planilha!"); return; }

    let btn = document.getElementById("btnImportar");
    let txtOriginal = btn.innerHTML;
    btn.innerHTML = "⏳ Importando...";
    btn.disabled = true;

    setTimeout(() => {
        let linhas = texto.split('\n');
        if (linhas.length < 2) { alert("Copie os cabeçalhos das colunas."); btn.innerHTML = txtOriginal; btn.disabled = false; return; }

        function encontrarIndice(cabs, chaves) {
            for (let i = 0; i < cabs.length; i++) {
                let cLimpo = cabs[i].replace(/_/g, " ");
                for (let k of chaves) { if (cLimpo.includes(k)) return i; }
            }
            return -1;
        }

        let cabs = linhas[0].split('\t').map(col => col.trim().toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, ""));
        
        let iDia = encontrarIndice(cabs, ["DIA DA SEMAN"]);
        let iCurso = encontrarIndice(cabs, ["NOM CURSO"]);
        let iHoraIni = encontrarIndice(cabs, ["MINDEHH INI"]);
        let iHoraFim = encontrarIndice(cabs, ["MAXDEHH FIM"]);
        let iCodDisc = encontrarIndice(cabs, ["COD DISCIPLIN"]);
        let iNomDisc = encontrarIndice(cabs, ["NOM DISCIPLINA"]);
        let iTurma = encontrarIndice(cabs, ["COD TURM"]);
        let iProf = encontrarIndice(cabs, ["NOM PROFESSOR"]);
        let iAlunos = encontrarIndice(cabs, ["QTD ALUNOS"]);
        let iPerAcad = encontrarIndice(cabs, ["ID PERIODO", "PERIODO"]);
        
        let iDataIni = -1, iDataFim = -1;
        for (let i = 0; i < cabs.length; i++) {
            let cl = cabs[i].replace(/_/g, " ");
            if (cl.includes("ALOCACA") && !cl.includes("FIM")) iDataIni = i;
            if (cl.includes("DT FIM") || cl.includes("FIM ALOCACA")) iDataFim = i;
        }
        if (iDataIni === -1) iDataIni = encontrarIndice(cabs, ["ALOCACA"]);
        if (iDataFim === -1) iDataFim = encontrarIndice(cabs, ["DT FIM"]);

        let novos = [];
        let index = 1;

        function lote() {
            let limit = Math.min(index + 5000, linhas.length);
            for (; index < limit; index++) {
                let v = linhas[index].split('\t');
                if (v.length <= 1 && v[0].trim() === "") continue;
                novos.push({
                    idUnico: gerarIdUnico(),
                    diaSemana: v[iDia !== -1 ? iDia : -1] || "", curso: v[iCurso !== -1 ? iCurso : -1] || "",
                    horaIni: v[iHoraIni !== -1 ? iHoraIni : -1] || "", horaFim: v[iHoraFim !== -1 ? iHoraFim : -1] || "",
                    codDisc: v[iCodDisc !== -1 ? iCodDisc : -1] || "", turma: v[iTurma !== -1 ? iTurma : -1] || "",
                    nomDisc: v[iNomDisc !== -1 ? iNomDisc : -1] || "", prof: v[iProf !== -1 ? iProf : -1] || "",
                    alunos: parseInt(v[iAlunos !== -1 ? iAlunos : -1]) || 0,
                    dataIni: v[iDataIni !== -1 ? iDataIni : -1] || "", dataFimAloc: v[iDataFim !== -1 ? iDataFim : -1] || "",
                    periodoAcad: v[iPerAcad !== -1 ? iPerAcad : -1] || ""
                });
            }
            if (index < linhas.length) requestAnimationFrame(lote);
            else {
                dados = dados.concat(novos);
                salvarDadosDB(dados); 
                
                atualizarFiltrosDinamicos(); 
                executarProcessamento(); 
                
                btn.innerHTML = txtOriginal; btn.disabled = false;
                document.getElementById("textoExcel").value = ""; 

                let mod = document.getElementById('modalImportacao');
                if (mod) mod.classList.remove('active');

                alert(`✅ ${novos.length} linhas adicionadas!\nElas foram salvas na memória fixa.\nTotal: ${dados.length}.`);
            }
        }
        lote(); 
    }, 50);
}

function processarArquivoCascata(event) {
    const file = event.target.files[0];
    if (!file) return;

    mostrarLoading();
    const reader = new FileReader();

    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonBruto = XLSX.utils.sheet_to_json(sheet, { raw: false });

            let index = 0;
            const total = jsonBruto.length;
            const tamanhoLote = 500; 
            let novosDados = [];

            function processarLote() {
                const fim = Math.min(index + tamanhoLote, total);
                
                for (; index < fim; index++) {
                    const v = jsonBruto[index];
                    novosDados.push({
                        idUnico: gerarIdUnico(),
                        diaSemana: v["Dia da Seman"] || "",
                        curso: v["NOM_CURSO"] || v["NOM CURSO"] || "",
                        horaIni: v["MinDeHH_INI_AUL"] || v["MinDeHH INI"] || "",
                        horaFim: v["MaxDeHH_FIM_AUL"] || v["MaxDeHH FIM"] || "",
                        codDisc: v["COD_DISCIPLIN"] || v["COD DISCIPLIN"] || "",
                        turma: v["COD_TURM"] || v["COD TURM"] || "",
                        nomDisc: v["NOM_DISCIPLINA"] || v["NOM DISCIPLINA"] || "",
                        prof: v["NOM_PROFESSOR"] || v["NOM PROFESSOR"] || "",
                        alunos: parseInt(v["QTD_ALUNOS_MATRICULADO"]) || parseInt(v["QTD ALUNOS"]) || 0,
                        dataIni: v["ALOCACA"] || "",
                        dataFimAloc: v["DT_FIM_ALOCACA"] || v["DT FIM"] || "",
                        periodoAcad: v["ID_PERIODO"] || v["ID PERIODO"] || v["PERIODO"] || ""
                    });
                }

                if (index < total) {
                    requestAnimationFrame(processarLote); 
                } else {
                    dados = dados.concat(novosDados);
                    salvarDadosDB(dados);
                    atualizarFiltrosDinamicos();
                    executarProcessamento();
                    ocultarLoading();
                    const modal = document.getElementById('modalImportacao');
                    if (modal) modal.classList.remove('active');
                    alert(`✅ Sucesso! ${total} linhas processadas e salvas.`);
                }
            }
            processarLote();
        } catch (err) {
            ocultarLoading();
            alert("Erro ao ler o arquivo Excel. Verifique o formato.");
        }
    };
    reader.readAsArrayBuffer(file);
}

function converterDataComparacao(valorData) {
    if (!valorData) return "";
    let dataStr = String(valorData).trim();
    if (dataStr.includes("/")) {
        let p = dataStr.split("/");
        return p[0].length === 4 ? `${p[0]}-${p[1]}-${p[2]}` : `${p[2]}-${p[1]}-${p[0]}`;
    }
    return dataStr;
}

function formatarDataVisual(valorData) {
    if (!valorData) return "";
    let dataStr = String(valorData).trim();
    if (dataStr.includes("-")) {
        let p = dataStr.split("-");
        return p[0].length === 4 ? `${p[2]}/${p[1]}/${p[0]}` : dataStr;
    }
    return dataStr;
}

function limparMemoria() {
    if(confirm("Apagar APENAS as aulas importadas da memória? (Configurações, Mapa e Consolidados serão mantidos intactos)")) {
        dados = []; 
        windowDadosResultantes = [];
        turmasOcultas = []; // Limpa apenas as turmas ocultadas, pois as aulas sumiram
        
        limparDB(); // Chama a função que agora apaga só o ID 1
        salvarManutencaoDB(); // Atualiza a memória para limpar turmasOcultas, mantendo salasEmManutencao intactas
        
        document.getElementById("resultado").innerHTML = "";
        let contador = document.getElementById('contadorAulasFound');
        if (contador) contador.innerHTML = '0 AULAS ENCONTRADAS';
        
        if(chartDiasInstancia) { chartDiasInstancia.destroy(); chartDiasInstancia = null; }
        if(chartCursosInstancia) { chartCursosInstancia.destroy(); chartCursosInstancia = null; }
        
        atualizarFiltrosDinamicos();
        renderizarMapaSalas(); // Atualiza o mapa para zerar os percentuais de uso
        alert("Aulas importadas apagadas com sucesso! Suas configurações foram mantidas.");
    }
}

/// =========================================================
// CONTROLE DOS MODAIS FLUTUANTES E ABAS
// =========================================================

let timerOcultarAbas = null;

function reiniciarTimerAbas() {
    clearTimeout(timerOcultarAbas);
    
    // --- NOVA TRAVA: Se a aba Consolidado estiver aberta, aborta o fechamento automático ---
    let abaCons = document.getElementById('abaConsolidado');
    if (abaCons && abaCons.classList.contains('active')) {
        return; 
    }

    timerOcultarAbas = setTimeout(() => {
        document.querySelectorAll(".tab-content.active").forEach(c => {
            // Aplica a transição CSS diretamente via estilo inline para recolher suavemente
            c.style.transition = "all 0.5s ease-in-out";
            c.style.opacity = "0";
            c.style.transform = "translateY(-10px)";
            
            // Aguarda a transição terminar antes de realmente esconder o elemento
            setTimeout(() => {
                c.classList.remove("active");
                // Reseta os estilos inline para a próxima vez que for aberto
                c.style.opacity = "";
                c.style.transform = "";
                c.style.transition = "";
            }, 500); 
        });
        document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
    }, 20000); // 20 segundos
}

function mudarAba(evento, idAba) {
    let abaDestino = document.getElementById(idAba);
    let botaoClicado = evento.currentTarget;
    let jaEstaAtiva = abaDestino.classList.contains("active");
    
    // Remove a classe 'active' de todas as abas e botões
    document.querySelectorAll(".tab-content").forEach(c => c.classList.remove("active"));
    document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));

    // Seleciona os elementos que devem sumir (Filtros e o Card da Tabela Principal)
    let containerFiltros = document.getElementById('containerHubBase');
    let cardTabela = document.getElementById('resultado') ? document.getElementById('resultado').closest('.card') : null;

    // Se não estava ativa, ativa agora. Se já estava, ela apenas foi escondida pelo código acima.
    if (!jaEstaAtiva) {
        if (abaDestino) abaDestino.classList.add("active");
        if (botaoClicado) botaoClicado.classList.add("active");
        
        // --- NOVA LÓGICA DE OCULTAÇÃO ---
        if (idAba === 'abaConsolidado') {
            if (containerFiltros) containerFiltros.style.display = 'none';
            if (cardTabela) cardTabela.style.display = 'none';
        } else {
            if (containerFiltros) containerFiltros.style.display = 'flex';
            if (cardTabela) cardTabela.style.display = 'block';
        }
        
        // Se a aba clicada for a do mapa, força a renderização das salas
        if (idAba === 'abaMapa') {
            if (typeof renderizarMapaSalas === "function") {
                renderizarMapaSalas();
            }
        }
        
        // Reinicia a contagem dos 10 segundos ao abrir a aba
        reiniciarTimerAbas();
    } else {
         // Se fechou a aba manualmente, cancela o timer e restaura a visualização padrão
         clearTimeout(timerOcultarAbas);
         if (containerFiltros) containerFiltros.style.display = 'flex';
         if (cardTabela) cardTabela.style.display = 'block';
    }
}

// Impede que as abas fechem enquanto o usuário estiver com o mouse sobre elas
document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll(".tab-content").forEach(aba => {
        aba.addEventListener('mouseenter', () => clearTimeout(timerOcultarAbas));
        aba.addEventListener('mouseleave', () => {
            if(aba.classList.contains('active')) reiniciarTimerAbas();
        });
    });
});

function abrirModalImportacao(e) {
    if(e) e.preventDefault();
    document.getElementById('modalImportacao').classList.add('active');
    setTimeout(() => document.getElementById('textoExcel').focus(), 100);
}

function fecharModalImportacao() {
    document.getElementById('modalImportacao').classList.remove('active');
}

function abrirModalConfig(e) {
    if(e) e.preventDefault();
    document.getElementById('modalConfig').classList.add('active');
    renderizarConfiguracoes(); 
}

function fecharModalConfig() {
    document.getElementById('modalConfig').classList.remove('active');
    if (typeof cancelarEdicaoSala === "function") {
        cancelarEdicaoSala(); 
    }
}

function renderizarMapaSalas() {
    let cont = document.getElementById('mapaSalasContainer');
    if(!cont) return;

    // Agrupa as salas por bloco
    let blocos = {};
    blocosCadastrados.forEach(b => blocos[b] = []);
    
    salas.forEach(s => {
        let letraBloco = s.name.charAt(0).toUpperCase();
        if(!blocos[letraBloco]) blocos[letraBloco] = [];
        blocos[letraBloco].push(s);
    });

    let html = '';
    
    for(let b of blocosCadastrados) {
        if(!blocos[b] || blocos[b].length === 0) continue;
        
        html += `<div class="bloco-container">
                    <div class="bloco-title">Bloco ${b}</div>
                    <div class="bloco-salas">`;
        
        // Ordena as salas numericamente e cria os cards
        blocos[b].sort((a,b) => a.name.localeCompare(b.name)).forEach(s => {
            let usoMax = 0;
            if(windowDadosResultantes && windowDadosResultantes.length > 0) {
                let aulasSala = windowDadosResultantes.filter(r => r.salaSugerida === s.name);
                aulasSala.forEach(a => { if(a.percUso > usoMax) usoMax = a.percUso; });
            }
            
            let w = Math.min(usoMax, 100);
            let cor = usoMax > 100 ? '#ef4444' : '#10b981'; 
            if (usoMax === 0) cor = '#cbd5e1';

            // VERIFICA O STATUS DA SALA PARA APLICAR A CLASSE E O ÍCONE
            let classeStatus = '';
            let iconeStatus = '';
            
            if (salasEmManutencao.includes(s.name)) {
                classeStatus = 'status-manutencao';
                iconeStatus = '🔴 ';
            } else if (salasProjetorRuim.includes(s.name)) {
                classeStatus = 'status-projetor';
                iconeStatus = '🟡 ';
            }

            html += `
                <div class="sala-card ${classeStatus}" onclick="destacarSalaNaTabela('${s.name}')" oncontextmenu="alternarManutencaoMapa('${s.name}'); return false;" title="Esquerdo: Buscar na Tabela | Direito: Alterar Status">
                    <div class="sala-name">${iconeStatus}${s.name}</div>
                    <div class="sala-cap">Cap: ${s.cap}</div>
                    <div class="sala-progress-bg">
                        <div class="sala-progress-fill" style="width: ${w}%; background: ${cor};"></div>
                    </div>
                    <div class="sala-uso">${usoMax}% de uso max.</div>
                </div>
            `;
        });
        html += `</div></div>`;
    }
    
    cont.innerHTML = html;
}

function destacarSalaNaTabela(nomeSala) {
    let linhas = document.querySelectorAll('#resultado tr');
    let primeiraLinha = null;
    
    // Varre as linhas procurando a sala
    linhas.forEach(tr => {
        tr.classList.remove('linha-piscando'); // Limpa animações anteriores
        
        // Verifica a coluna 8 (onde fica a Sala Sugerida)
        if (tr.children[8] && tr.children[8].innerText.includes(nomeSala)) {
            setTimeout(() => tr.classList.add('linha-piscando'), 50);
            
            // Salva a primeira linha encontrada na variável
            if (!primeiraLinha) {
                primeiraLinha = tr;
            }
        }
    });

    // Rola a tela exatamente para a linha encontrada, centralizando no monitor
    if (primeiraLinha) {
        primeiraLinha.scrollIntoView({ behavior: 'smooth', block: 'center' });
    } else {
        // Se a sala estiver vazia (sem aulas), rola apenas para o topo da tabela
        let tabelaContainer = document.querySelector('.card:last-child');
        if (tabelaContainer) tabelaContainer.scrollIntoView({ behavior: 'smooth' });
    }
}

// =========================================================
// FUNÇÕES DA JANELA SUSPENSA (EDITAR E INSERIR NA TABELA)
// =========================================================

function abrirModalFormularioAula(idUnico = null) {
    // Remove modal antigo para não duplicar
    let modalAntigo = document.getElementById('modalFormAula');
    if (modalAntigo) modalAntigo.remove();

    let aula = null;
    let titulo = "➕ Adicionar Nova Aula";
    let btnTexto = "Adicionar Aula";

    if (idUnico) {
        aula = dados.find(d => d.idUnico === idUnico);
        if (!aula) return alert("Aula não encontrada!");
        titulo = `✏️ Editando Aula (ID: ${idUnico})`;
        btnTexto = "Salvar Alterações";
    } else {
        // Objeto em branco para nova aula
        aula = { diaSemana: 'SEGUNDA', horaIni: '07:50', horaFim: '10:30', curso: '', nomDisc: '', turma: '', prof: '', alunos: 0, salaManual: '' };
    }

    let cursos = [...new Set(dados.map(d => d.curso).filter(Boolean))].sort();
    let profs = [...new Set(dados.map(d => d.prof).filter(Boolean))].sort();
    let disciplinas = [...new Set(dados.map(d => d.nomDisc).filter(Boolean))].sort();
    let periodos = [...new Set(dados.map(d => d.periodoAcad).filter(Boolean))].sort(); // Lista de períodos para o autocompletar

    let htmlModal = `
    <div id="modalFormAula" class="modal-overlay-custom active" style="display:flex; z-index:100000;">
        <div class="modal-content-custom" style="max-width: 600px; text-align: left;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; border-bottom: 2px solid #eee; padding-bottom: 10px;">
                <h3 style="margin:0; color:#1d7bd8;">${titulo}</h3>
                <button onclick="document.getElementById('modalFormAula').remove()" style="background:none; border:none; color:#dc3545; font-size:24px; cursor:pointer; font-weight:bold; padding:0; line-height:1;">✖</button>
            </div>
            
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 20px;">
                <div class="form-group">
                    <label style="font-size:12px; color:#555;">Dia da Semana:</label>
                    <select id="modalEdDia">
                        ${['SEGUNDA', 'TERÇA', 'QUARTA', 'QUINTA', 'SEXTA', 'SÁBADO'].map(d => `<option value="${d}" ${(aula.diaSemana||'').toUpperCase() === d ? 'selected' : ''}>${d}</option>`).join('')}
                    </select>
                </div>
                <div style="display:flex; gap:10px;">
                    <div class="form-group" style="flex:1;">
                        <label style="font-size:12px; color:#555;">Início:</label>
                        <input type="time" id="modalEdIni" value="${aula.horaIni || '07:50'}">
                    </div>
                    <div class="form-group" style="flex:1;">
                        <label style="font-size:12px; color:#555;">Término:</label>
                        <input type="time" id="modalEdFim" value="${aula.horaFim || '10:30'}">
                    </div>
                </div>
                <div class="form-group">
                    <label style="font-size:12px; color:#555;">Curso:</label>
                    <input type="text" id="modalEdCurso" value="${(aula.curso || '').replace(/"/g, '&quot;')}" list="dlModalCursos" placeholder="Digite ou selecione">
                    <datalist id="dlModalCursos">${cursos.map(c => `<option value="${c}"></option>`).join('')}</datalist>
                </div>
                <div class="form-group">
                    <label style="font-size:12px; color:#555;">Disciplina:</label>
                    <input type="text" id="modalEdDisc" value="${(aula.nomDisc || '').replace(/"/g, '&quot;')}" list="dlModalDisc" placeholder="Digite ou selecione">
                    <datalist id="dlModalDisc">${disciplinas.map(d => `<option value="${d}"></option>`).join('')}</datalist>
                </div>
                <div class="form-group">
                    <label style="font-size:12px; color:#555;">Turma(s):</label>
                    <input type="text" id="modalEdTurma" value="${(aula.turma || '').replace(/"/g, '&quot;')}" placeholder="Ex: 1001, 1002">
                </div>
                <div class="form-group">
                    <label style="font-size:12px; color:#555;">Período:</label>
                    <input type="text" id="modalEdPeriodo" value="${(aula.periodoAcad || '').replace(/"/g, '&quot;')}" list="dlModalPeriodos" placeholder="Ex: 1º PERÍODO">
                    <datalist id="dlModalPeriodos">${periodos.map(p => `<option value="${p}"></option>`).join('')}</datalist>
                </div>
                <div class="form-group">
                    <label style="font-size:12px; color:#555;">Professor:</label>
                    <input type="text" id="modalEdProf" value="${(aula.prof || '').replace(/"/g, '&quot;')}" list="dlModalProfs" placeholder="Digite ou selecione">
                    <datalist id="dlModalProfs">${profs.map(p => `<option value="${p}"></option>`).join('')}</datalist>
                </div>
                <div class="form-group">
                    <label style="font-size:12px; color:#555;">Qtd. Alunos:</label>
                    <input type="number" id="modalEdAlunos" value="${aula.alunos || 0}">
                </div>
                <div class="form-group">
                    <label style="font-size:12px; color:#555;">Forçar Sala / Status:</label>
                    <select id="modalEdSala">
                        <option value="">-- MODO AUTOMÁTICO --</option>
                        <option value="FALTOU" ${aula.salaManual === 'FALTOU' ? 'selected' : ''} style="color: #dc3545; font-weight: bold;">❌ FALTOU</option>
                        ${salas.map(s => `<option value="${s.name}" ${aula.salaManual === s.name ? 'selected' : ''}>${s.name} (Cap: ${s.cap})</option>`).join('')}
                    </select>
                </div>
            </div>
            <button type="button" onclick="salvarModalFormAula('${idUnico || ''}')" style="background:#1fbf75; width:100%; padding:12px; border:none; color:white; font-weight:bold; border-radius:6px; cursor:pointer;">💾 ${btnTexto}</button>
        </div>
    </div>`;
    
    document.body.insertAdjacentHTML('beforeend', htmlModal);
}

function salvarModalFormAula(idExistente) {
    let aula;
    if (idExistente) {
        aula = dados.find(d => d.idUnico === idExistente);
        if (!aula) return alert("Erro ao salvar: Aula não encontrada na memória.");
    } else {
        // Criar nova aula
        aula = { idUnico: gerarIdUnico(), codDisc: '', periodoAcad: '' };
        dados.push(aula);
    }

    aula.diaSemana = document.getElementById('modalEdDia').value;
    aula.horaIni = document.getElementById('modalEdIni').value;
    aula.horaFim = document.getElementById('modalEdFim').value;
    aula.curso = document.getElementById('modalEdCurso').value.trim().toUpperCase();
    aula.nomDisc = document.getElementById('modalEdDisc').value.trim().toUpperCase();
    aula.turma = document.getElementById('modalEdTurma').value.trim().toUpperCase();
    aula.periodoAcad = document.getElementById('modalEdPeriodo').value.trim().toUpperCase(); // <-- ADICIONE ESTA LINHA AQUI
    aula.prof = document.getElementById('modalEdProf').value.trim().toUpperCase();
    aula.alunos = parseInt(document.getElementById('modalEdAlunos').value) || 0;
    aula.salaManual = document.getElementById('modalEdSala').value;

    salvarDadosDB(dados);
    document.getElementById('modalFormAula').remove();
    
    alert(idExistente ? "✅ Alterações salvas com sucesso!" : "✅ Nova aula inserida com sucesso!");
    
    atualizarFiltrosDinamicos();
    executarProcessamento();
}

// Funções ponte para os botões do HTML
function abrirEdicaoDireta(id) { abrirModalFormularioAula(id); }
function abrirModalNovaAula() { abrirModalFormularioAula(null); }

function iniciarConsolidacao() {
    if (!windowDadosResultantes || windowDadosResultantes.length === 0) {
        alert("⚠️ Não há dados processados na tabela! Clique em 'Processar Dados' antes de consolidar.");
        return;
    }
    
    // Mostra o balão com a mensagem suavemente
    let balaoMsg = document.getElementById('msgConsolidacao');
    if (balaoMsg) balaoMsg.style.opacity = '1';
    
    // Abre o calendário nativo do HTML para o usuário escolher a data
    let inputData = document.getElementById('inputDataConsolidacao');
    if (inputData) {
        inputData.showPicker();
        
        // Se o usuário clicar fora do calendário (cancelar), o balão some
        inputData.addEventListener('blur', () => {
            setTimeout(() => {
                if (balaoMsg) balaoMsg.style.opacity = '0';
            }, 200);
        }, { once: true });
    }
}

function finalizarConsolidacao(dataEscolhida) {
    // Esconde o balão imediatamente após a escolha
    let balaoMsg = document.getElementById('msgConsolidacao');
    if (balaoMsg) balaoMsg.style.opacity = '0';

    if (!dataEscolhida) return;
    
    // Salva uma cópia profunda da tabela atual na chave da data escolhida
    dadosConsolidados[dataEscolhida] = JSON.parse(JSON.stringify(windowDadosResultantes));
    salvarConsolidadoDB();
    
    alert(`✅ Consolidação salva com sucesso para o dia: ${dataEscolhida.split('-').reverse().join('/')}`);
    
    let inputData = document.getElementById('inputDataConsolidacao');
    if (inputData) inputData.value = ''; // Limpa o input
}

function renderizarTabelaConsolidada() {
    let data = document.getElementById('dataFiltroConsolidado').value;
    let tbody = document.getElementById('resultadoConsolidado');
    tbody.innerHTML = "";
    
    if (!data || !dadosConsolidados[data]) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align:center; color: #dc3545; font-weight: bold;">Nenhuma consolidação encontrada para esta data.</td></tr>';
        return;
    }
    
    let html = "";
    dadosConsolidados[data].forEach(row => {
        let salaStyle = (row.salaSugerida === "SEM SALA" || row.salaSugerida === "FALTOU") ? "color:#dc3545; font-weight:bold;" : "color:#10b981; font-weight:bold;";
        
        let widthBar = Math.min(row.percUso, 100);
        let styleBarraVermelha = row.percUso > 100 ? 'background: linear-gradient(90deg, #ef4444, #b91c1c);' : '';
        
        // Muda o fundo da linha e a borda se faltou
        let trStyle = row.salaSugerida === "FALTOU" ? "background-color: #fee2e2; border-bottom: 1px solid #fca5a5;" : "border-bottom: 1px solid #eee;";

        html += `<tr style="${trStyle}">
            <td style="padding: 8px;">${row.diaSemana}</td>
            <td>${(row.horaIni || "").substring(0, 5)} às ${(row.horaFim || "").substring(0, 5)}</td>
            <td>${row.curso}</td>
            <td style="font-size: 11px; font-weight: 600; color: #333;">${row.codDisc || ''} ${row.turma || ''} ${row.nomDisc || ''}</td>
            <td>${row.prof}</td>
            <td style="${salaStyle}">${row.salaSugerida}</td>
            <td>
                <div class="progress">
                    <div class="progress-bar" style="width:${widthBar}%; ${styleBarraVermelha}"></div>
                </div>
                ${row.percUso}%
            </td>
        </tr>`;
    });
    tbody.innerHTML = html;
}

// Variável para armazenar as listas copiadas em cada coluna
let dadosColunasTemp = { dia: [], curso: [], horario: [], periodo: [], disciplina: [], prof: [], sala: [] };
let colunaAtiva = '';

function abrirModalConsolidadoManual() {
    document.getElementById('modalConsolidadoManual').classList.add('active');
    let dataFiltro = document.getElementById('dataFiltroConsolidado').value;
    if (dataFiltro) document.getElementById('manConsData').value = dataFiltro;
    
    // Zera os dados e esconde a caixa de texto sempre que abrir a janela
    dadosColunasTemp = { dia: [], curso: [], horario: [], periodo: [], disciplina: [], prof: [], sala: [] };
    colunaAtiva = '';
    document.getElementById('areaCaixaTexto').style.display = 'none';
    
    // Reseta as cores dos botões
    document.querySelectorAll('[id^="btnCol-"]').forEach(b => {
        b.style.background = '#f8f9fa'; b.style.color = 'black'; b.style.borderColor = '#ccc';
    });
}

function fecharModalConsolidadoManual() {
    document.getElementById('modalConsolidadoManual').classList.remove('active');
}

// Ao clicar em um ícone de coluna, abre a caixa correspondente
function abrirCaixaColuna(coluna) {
    colunaAtiva = coluna;
    document.getElementById('areaCaixaTexto').style.display = 'block';
    
    let nomes = { dia: 'Dia da Semana', curso: 'Curso', horario: 'Horário', periodo: 'Período', disciplina: 'Disciplina', prof: 'Professor(a)', sala: 'Salas' };
    document.getElementById('labelCaixaTexto').innerText = `Cole os dados copiados para: ${nomes[coluna]}`;
    
    // Mostra o texto que o usuário já tinha colado antes nesta mesma coluna (se houver)
    document.getElementById('textoColunaAtual').value = dadosColunasTemp[coluna].join('\n');
    atualizarContadorLinhas();
    
    // Pinta o botão clicado de azul
    document.querySelectorAll('[id^="btnCol-"]').forEach(b => {
        b.style.background = '#f8f9fa'; b.style.color = 'black'; b.style.borderColor = '#ccc';
    });
    let btnAtivo = document.getElementById(`btnCol-${coluna}`);
    btnAtivo.style.background = '#1d7bd8'; btnAtivo.style.color = 'white'; btnAtivo.style.borderColor = '#1d7bd8';
}

// Salva o texto digitado/colado sempre que o usuário altera a caixa
function salvarDadosColunaAtual() {
    if (!colunaAtiva) return;
    let texto = document.getElementById('textoColunaAtual').value;
    // Separa o texto por quebra de linha (cada linha copiada vira um item)
    dadosColunasTemp[colunaAtiva] = texto.split('\n');
    atualizarContadorLinhas();
}

function atualizarContadorLinhas() {
    // Conta quantas linhas não estão totalmente vazias
    let num = dadosColunasTemp[colunaAtiva] ? dadosColunasTemp[colunaAtiva].filter(t => t.trim() !== "").length : 0;
    document.getElementById('contadorLinhasColuna').innerText = num;
}

// Monta a tabela juntando todas as colunas linha por linha
function salvarColunasNoConsolidado() {
    let dataCons = document.getElementById('manConsData').value;
    if (!dataCons) { alert("⚠️ Selecione a Data da Consolidação no topo do formulário!"); return; }

    // Descobre qual coluna tem mais itens para saber quantas linhas criar
    let maxLinhas = 0;
    for (let key in dadosColunasTemp) {
        if (dadosColunasTemp[key].length > maxLinhas) maxLinhas = dadosColunasTemp[key].length;
    }

    if (maxLinhas === 0) { alert("⚠️ Nenhuma coluna foi preenchida!"); return; }

    if (!dadosConsolidados[dataCons]) dadosConsolidados[dataCons] = [];

    // Junta o Item 1 de cada coluna para fazer a Linha 1, depois o Item 2 para a Linha 2...
    for (let i = 0; i < maxLinhas; i++) {
        // Separa o horário (Ex: "07:50 às 10:30") em Inicio e Fim se necessário
        let horarioBruto = (dadosColunasTemp.horario[i] || "").trim();
        let hIni = horarioBruto, hFim = "";
        let partesHora = horarioBruto.split(/às|-/).filter(p => p.trim() !== "");
        if (partesHora.length >= 2) {
            hIni = partesHora[0].trim();
            hFim = partesHora[1].trim();
        }

        let novaLinha = {
            diaSemana: (dadosColunasTemp.dia[i] || "").trim().toUpperCase(),
            horaIni: hIni,
            horaFim: hFim,
            curso: (dadosColunasTemp.curso[i] || "").trim().toUpperCase(),
            periodoAcad: (dadosColunasTemp.periodo[i] || "").trim().toUpperCase(),
            nomDisc: (dadosColunasTemp.disciplina[i] || "").trim().toUpperCase(),
            prof: (dadosColunasTemp.prof[i] || "").trim().toUpperCase(),
            salaSugerida: (dadosColunasTemp.sala[i] || "").trim().toUpperCase(),
            percUso: 0,
            codDisc: "",
            turma: ""
        };
        
        // Só ignora linhas onde absolutamente tudo está vazio
        if (novaLinha.curso !== "" || novaLinha.nomDisc !== "" || novaLinha.prof !== "") {
            dadosConsolidados[dataCons].push(novaLinha);
        }
    }

    salvarConsolidadoDB();
    alert(`✅ Consolidação gerada com sucesso! Lidas ${maxLinhas} linhas.`);
    fecharModalConsolidadoManual();
    
    // Atualiza a tabela na tela
    if (document.getElementById('dataFiltroConsolidado').value === dataCons) {
        renderizarTabelaConsolidada();
    } else {
        document.getElementById('dataFiltroConsolidado').value = dataCons;
        renderizarTabelaConsolidada();
    }
}

function limparTodaMemoriaConsolidado() {
    if (confirm("⚠️ ATENÇÃO: Tem certeza que deseja APAGAR TODA a memória de tabelas consolidadas? Essa ação não poderá ser desfeita.")) {
        // Zera o objeto na memória do sistema
        dadosConsolidados = {}; 
        
        // Salva o objeto vazio no banco de dados (ID 4)
        initDB(function(db) {
            let tx = db.transaction(storeName, "readwrite");
            tx.objectStore(storeName).put({ id: 4, consolidado: {} });
        });
        
        // Limpa a visualização na Aba Consolidado se ela estiver aberta
        if (document.getElementById('resultadoConsolidado')) {
            document.getElementById('dataFiltroConsolidado').value = '';
            document.getElementById('resultadoConsolidado').innerHTML = '<tr><td colspan="7" style="text-align:center; color: #666;">Selecione uma data no calendário acima para visualizar.</td></tr>';
        }
        
        alert("🗑️ Memória do consolidado excluída com sucesso!");
    }
}

// =========================================================
// CONTROLE DO MODAL DE EXPORTAÇÃO DE PDF
// =========================================================

function abrirModalExportacao() {
    let dataCons = document.getElementById('dataFiltroConsolidado').value;
    
    // Trava de segurança: só deixa abrir se tiver uma data com dados selecionada
    if (!dataCons || !dadosConsolidados[dataCons] || dadosConsolidados[dataCons].length === 0) {
        alert("⚠️ Por favor, selecione primeiro uma data no calendário que possua dados consolidados para poder exportar.");
        return;
    }
    
    document.getElementById('modalExportacao').classList.add('active');
}

function fecharModalExportacao() {
    document.getElementById('modalExportacao').classList.remove('active');
}

function gerarPDFCompleto() {
    fecharModalExportacao(); 
    abrirConfigExportacao(); 
}

function executarExportacaoPDF(event) {
    let containerVisivel = document.getElementById('containerPreviewsPDF');
    if (!containerVisivel || containerVisivel.children.length === 0) {
        alert("⚠️ Gere a pré-visualização na tela primeiro!");
        return;
    }

    let btn = event.currentTarget;
    let textoOriginal = btn.innerHTML;
    btn.innerHTML = "⏳ Processando PDF...";
    btn.disabled = true;

    // 1. Clona o resultado da pré-visualização para imprimir
    let printWrapper = document.createElement('div');
    printWrapper.id = 'print-wrapper-temp';
    printWrapper.innerHTML = containerVisivel.innerHTML;
    document.body.appendChild(printWrapper);

    // 2. CSS Inteligente: Captura os valores AO VIVO da tela e força as regras no PDF
    let pTabTop = document.getElementById('printTabTop') ? document.getElementById('printTabTop').value : 25;
    let pTabLat = document.getElementById('printTabLat') ? document.getElementById('printTabLat').value : 5;
    let pTitTop = document.getElementById('printTitTop') ? document.getElementById('printTitTop').value : 8;
    let pTitFonte = document.getElementById('printTitFonte') ? document.getElementById('printTitFonte').value : 32;
    let pDataTop = document.getElementById('printDataTop') ? document.getElementById('printDataTop').value : 12;
    let pDataFonte = document.getElementById('printDataFonte') ? document.getElementById('printDataFonte').value : 24;
    let pTabFonte = document.getElementById('printTabFonte') ? document.getElementById('printTabFonte').value : 14;

    let w1 = document.getElementById('wCol1') ? document.getElementById('wCol1').value : 15;
    let w2 = document.getElementById('wCol2') ? document.getElementById('wCol2').value : 10;
    let w3 = document.getElementById('wCol3') ? document.getElementById('wCol3').value : 5;
    let w4 = document.getElementById('wCol4') ? document.getElementById('wCol4').value : 35;
    let w5 = document.getElementById('wCol5') ? document.getElementById('wCol5').value : 25;
    let w6 = document.getElementById('wCol6') ? document.getElementById('wCol6').value : 10;

    let stylePrint = document.createElement('style');
    stylePrint.innerHTML = `
        @media screen {
            #print-wrapper-temp { display: none !important; }
        }

        @media print {
            @page { size: 16in 9in; margin: 0; }
            
            body > *:not(#print-wrapper-temp):not(style):not(script) { display: none !important; }
            
            body { margin: 0; padding: 0; background: white; }
            * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
            
            #print-wrapper-temp { display: block !important; position: absolute; top: 0; left: 0; width: 100%; }
            
            #print-wrapper-temp > div {
                width: 100vw !important; height: 100vh !important;
                margin: 0 !important; border: none !important; box-shadow: none !important;
                page-break-after: always !important; page-break-inside: avoid !important;
                
                padding-top: ${pTabTop}% !important;
                padding-left: ${pTabLat}% !important;
                padding-right: ${pTabLat}% !important;
            }
            
            #print-wrapper-temp > div > div:nth-child(1) { top: ${pTitTop}% !important; }
            #print-wrapper-temp > div > div:nth-child(1) > div { font-size: ${pTitFonte}pt !important; }
            
            #print-wrapper-temp > div > div:nth-child(2) { top: ${pDataTop}% !important; }
            #print-wrapper-temp > div > div:nth-child(2) > div { font-size: ${pDataFonte}pt !important; }
            
            #print-wrapper-temp table {
                font-size: ${pTabFonte}pt !important;
                table-layout: fixed !important;
            }
            #print-wrapper-temp th, 
            #print-wrapper-temp td {
                font-size: ${pTabFonte}pt !important;
                white-space: nowrap !important;
                overflow: hidden !important;
                text-overflow: ellipsis !important; /* Fallback seguro */
                text-overflow: "." !important; /* Ponto final se o navegador suportar */
            }
            
            #print-wrapper-temp th:nth-child(1), #print-wrapper-temp td:nth-child(1) { width: ${w1}% !important; }
            #print-wrapper-temp th:nth-child(2), #print-wrapper-temp td:nth-child(2) { width: ${w2}% !important; text-align: center !important; }
            #print-wrapper-temp th:nth-child(3), #print-wrapper-temp td:nth-child(3) { width: ${w3}% !important; text-align: center !important; }
            #print-wrapper-temp th:nth-child(4), #print-wrapper-temp td:nth-child(4) { width: ${w4}% !important; }
            #print-wrapper-temp th:nth-child(5), #print-wrapper-temp td:nth-child(5) { width: ${w5}% !important; }
            #print-wrapper-temp th:nth-child(6), #print-wrapper-temp td:nth-child(6) { width: ${w6}% !important; text-align: center !important; }
        }
    `;
    document.head.appendChild(stylePrint);
   

   // 3. Invoca a impressão e limpa o lixo
    setTimeout(() => {
        window.print();
        
        printWrapper.remove();
        stylePrint.remove();
        
        btn.innerHTML = textoOriginal;
        btn.disabled = false;
    }, 500);
}

function gerarPDFResumido() {
    alert("Função em desenvolvimento.");
}

function gerarPDFImpressao() {
    alert("Função em desenvolvimento.");
}


// =========================================================
// CONFIGURAÇÕES DE EXPORTAÇÃO (LAYOUT PDF) E PRÉVIA
// =========================================================

let configPDF = { 
    bgImage: "", 
    mgTop: 35, mgBottom: 15, mgLeft: 5, mgRight: 5, 
    fontSize: 14,
    linhasPorPagina: 18,
    padCabecalho: 2,
    padLinha: 2,
    corCabecalho: "#ffffff", 
    corCorpo: "#ffffff",

    printTabFonte: 14,
    printTabTop: 25,
    printTabLat: 5,
    printTitFonte: 32,
    printTitTop: 8,
    printDataFonte: 24,
    printDataTop: 12,

    wCol1: 15, wCol2: 10, wCol3: 5, wCol4: 35, wCol5: 25, wCol6: 10,
    
    tituloTexto: "ENCONTRE SUA SALA",
    tituloPosicao: "right", 
    tituloTop: 8, 
    tituloPaddingX: 5, 
    tituloFontFamily: "Montserrat",
    tituloFontSize: 28,
    tituloCor: "#ffffff",

    dataAlign: "right",
    dataTop: 12,
    dataPaddingX: 5,
    dataFontFamily: "Montserrat",
    dataFontSize: 20,
    dataCor: "#ffffff"
};

function salvarConfigPdfDB() {
    initDB(function(db) {
        let tx = db.transaction(storeName, "readwrite");
        tx.objectStore(storeName).put({ id: 5, config: configPDF }); 
    });
}

function carregarConfigPdfDB(callback) {
    initDB(function(db) {
        let tx = db.transaction(storeName, "readonly");
        let req = tx.objectStore(storeName).get(5);
        req.onsuccess = function() {
            if (req.result && req.result.config) {
                configPDF = { ...configPDF, ...req.result.config };
            }
            if (callback) callback();
        };
    });
}

function abrirConfigExportacao() {
    document.getElementById('modalConfigExportacao').classList.add('active');
    
    document.getElementById('mgTop').value = configPDF.mgTop !== undefined ? configPDF.mgTop : 35;
    document.getElementById('mgBottom').value = configPDF.mgBottom !== undefined ? configPDF.mgBottom : 15;
    document.getElementById('mgLeft').value = configPDF.mgLeft !== undefined ? configPDF.mgLeft : 5;
    document.getElementById('mgRight').value = configPDF.mgRight !== undefined ? configPDF.mgRight : 5;
    document.getElementById('fontSizePDF').value = configPDF.fontSize || 14;
    document.getElementById('linhasPorPagina').value = configPDF.linhasPorPagina || 18;
    document.getElementById('padCabecalho').value = configPDF.padCabecalho !== undefined ? configPDF.padCabecalho : 2;
    document.getElementById('padLinha').value = configPDF.padLinha !== undefined ? configPDF.padLinha : 2;

    if(document.getElementById('printTabFonte')) {
        document.getElementById('printTabFonte').value = configPDF.printTabFonte || 14;
        document.getElementById('printTabTop').value = configPDF.printTabTop !== undefined ? configPDF.printTabTop : 25;
        document.getElementById('printTabLat').value = configPDF.printTabLat !== undefined ? configPDF.printTabLat : 5;
        document.getElementById('printTitFonte').value = configPDF.printTitFonte || 32;
        document.getElementById('printTitTop').value = configPDF.printTitTop !== undefined ? configPDF.printTitTop : 8;
        document.getElementById('printDataFonte').value = configPDF.printDataFonte || 24;
        document.getElementById('printDataTop').value = configPDF.printDataTop !== undefined ? configPDF.printDataTop : 12;

        document.getElementById('wCol1').value = configPDF.wCol1 || 15;
        document.getElementById('wCol2').value = configPDF.wCol2 || 10;
        document.getElementById('wCol3').value = configPDF.wCol3 || 5;
        document.getElementById('wCol4').value = configPDF.wCol4 || 35;
        document.getElementById('wCol5').value = configPDF.wCol5 || 25;
        document.getElementById('wCol6').value = configPDF.wCol6 || 10;

    }
    
    document.getElementById('corFonteCabecalho').value = configPDF.corCabecalho || "#ffffff";
    document.getElementById('corFonteCorpo').value = configPDF.corCorpo || "#ffffff";

    document.getElementById('tituloTexto').value = configPDF.tituloTexto || "ENCONTRE SUA SALA";
    document.getElementById('tituloPosicao').value = configPDF.tituloPosicao || "right";
    document.getElementById('tituloTop').value = configPDF.tituloTop !== undefined ? configPDF.tituloTop : 8;
    document.getElementById('tituloPaddingX').value = configPDF.tituloPaddingX !== undefined ? configPDF.tituloPaddingX : 5;
    document.getElementById('tituloFontSize').value = configPDF.tituloFontSize || 28;
    document.getElementById('tituloFontFamily').value = configPDF.tituloFontFamily || "Montserrat";
    document.getElementById('tituloCor').value = configPDF.tituloCor || "#ffffff";

    document.getElementById('dataAlign').value = configPDF.dataAlign || "right";
    document.getElementById('dataTop').value = configPDF.dataTop !== undefined ? configPDF.dataTop : 12;
    document.getElementById('dataPaddingX').value = configPDF.dataPaddingX !== undefined ? configPDF.dataPaddingX : 5;
    document.getElementById('dataFontFamily').value = configPDF.dataFontFamily || "Montserrat";
    document.getElementById('dataFontSize').value = configPDF.dataFontSize || 20;
    document.getElementById('dataCor').value = configPDF.dataCor || "#ffffff";
    
    atualizarPreviewPDF(); 
}

function fecharConfigExportacao() {
    document.getElementById('modalConfigExportacao').classList.remove('active');
}

function atualizarPreviewPDF() {
    let t = document.getElementById('mgTop').value || 0;
    let b = document.getElementById('mgBottom').value || 0;
    let l = document.getElementById('mgLeft').value || 0;
    let r = document.getElementById('mgRight').value || 0;
    let fs = document.getElementById('fontSizePDF').value || 14;
    let linhasPorPagina = parseInt(document.getElementById('linhasPorPagina').value) || 18;
    let padCab = document.getElementById('padCabecalho') ? document.getElementById('padCabecalho').value : 2;
    let padLin = document.getElementById('padLinha') ? document.getElementById('padLinha').value : 2;
    
    let cCab = document.getElementById('corFonteCabecalho').value;
    let cCorpo = document.getElementById('corFonteCorpo').value;

    let titTexto = document.getElementById('tituloTexto').value || "ENCONTRE SUA SALA";
    let titPos = document.getElementById('tituloPosicao').value || "right";
    let titTop = document.getElementById('tituloTop').value || 8;
    let titPadX = document.getElementById('tituloPaddingX').value || 5;
    let titSize = document.getElementById('tituloFontSize').value || 28;
    let titFont = document.getElementById('tituloFontFamily').value || "Montserrat";
    let titCor = document.getElementById('tituloCor').value || "#ffffff";

    let dAlign = document.getElementById('dataAlign').value || "right";
    let dTop = document.getElementById('dataTop').value || 12;
    let dPadX = document.getElementById('dataPaddingX').value || 5;
    let dFont = document.getElementById('dataFontFamily').value || "Montserrat";
    let dSize = document.getElementById('dataFontSize').value || 20;
    let dCor = document.getElementById('dataCor').value || "#ffffff";

    let container = document.getElementById('containerPreviewsPDF');
    if (!container) return;
    

    // Lendo os valores AO VIVO das caixinhas das colunas
    let w1 = document.getElementById('wCol1') ? document.getElementById('wCol1').value : 15;
    let w2 = document.getElementById('wCol2') ? document.getElementById('wCol2').value : 10;
    let w3 = document.getElementById('wCol3') ? document.getElementById('wCol3').value : 5;
    let w4 = document.getElementById('wCol4') ? document.getElementById('wCol4').value : 35;
    let w5 = document.getElementById('wCol5') ? document.getElementById('wCol5').value : 25;
    let w6 = document.getElementById('wCol6') ? document.getElementById('wCol6').value : 10;
    
    // A ARMA SECRETA: Injeta um CSS exclusivo pro simulador que domina as regras globais
    container.innerHTML = `
        <style>
            #containerPreviewsPDF table {
                table-layout: fixed !important;
            }
            #containerPreviewsPDF th, #containerPreviewsPDF td { 
                box-sizing: border-box !important;
                white-space: nowrap !important; 
                overflow: hidden !important;
                text-overflow: ellipsis !important; /* Fallback seguro */
                text-overflow: "." !important; /* Ponto final se o navegador suportar */
            }
            
            /* LARGURAS EXCLUSIVAS DE CADA COLUNA E CENTRALIZAÇÃO */
            #containerPreviewsPDF th:nth-child(1), #containerPreviewsPDF td:nth-child(1) { width: ${w1}% !important; }
            #containerPreviewsPDF th:nth-child(2), #containerPreviewsPDF td:nth-child(2) { width: ${w2}% !important; text-align: center !important; }
            #containerPreviewsPDF th:nth-child(3), #containerPreviewsPDF td:nth-child(3) { width: ${w3}% !important; text-align: center !important; }
            #containerPreviewsPDF th:nth-child(4), #containerPreviewsPDF td:nth-child(4) { width: ${w4}% !important; }
            #containerPreviewsPDF th:nth-child(5), #containerPreviewsPDF td:nth-child(5) { width: ${w5}% !important; }
            #containerPreviewsPDF th:nth-child(6), #containerPreviewsPDF td:nth-child(6) { width: ${w6}% !important; text-align: center !important; }
        </style>
    `;

    // 1. Obter dados do documento atual (Consolidado ou Aba Tabela Principal)
    let dataCons = document.getElementById('dataFiltroConsolidado').value;
    
    // Verifica se a aba ativa é a do Consolidado
    let abaConsolidadoAtiva = document.getElementById('abaConsolidado') && document.getElementById('abaConsolidado').classList.contains('active');
    
    let dadosAtuais = [];
    if (abaConsolidadoAtiva && dataCons && dadosConsolidados[dataCons]) {
        dadosAtuais = dadosConsolidados[dataCons];
    } else if (windowDadosResultantes && windowDadosResultantes.length > 0) {
        dadosAtuais = windowDadosResultantes;
    }
    
    // Se não tiver dados, cria uma linha fake para visualização
    if (!dadosAtuais || dadosAtuais.length === 0) {
        dadosAtuais = [{ curso: "EXEMPLO CURSO", horaIni: "19:00", periodoAcad: "1º", nomDisc: "NENHUM DADO ATIVO", prof: "PROFESSOR(A)", salaSugerida: "C101" }];
    }

    let totalPaginas = Math.ceil(dadosAtuais.length / linhasPorPagina);
    if (totalPaginas === 0) totalPaginas = 1;

    // 2. Data formatada para o cabeçalho
    const dias = ['DOMINGO', 'SEGUNDA-FEIRA', 'TERÇA-FEIRA', 'QUARTA-FEIRA', 'QUINTA-FEIRA', 'SEXTA-FEIRA', 'SÁBADO'];
    let dataCabecalho = "";
    if (abaConsolidadoAtiva && dataCons) {
        let p = dataCons.split('-');
        let dObj = new Date(p[0], p[1] - 1, p[2]);
        dataCabecalho = `${dias[dObj.getDay()]} (${p[2]}/${p[1]}/${p[0]})`;
    } else if (dadosAtuais.length > 0 && dadosAtuais[0].dataIni) {
        // Tenta pegar a data da primeira aula processada
        let dataAula = converterDataComparacao(dadosAtuais[0].dataIni);
        let p = dataAula.split('-');
        if(p.length === 3) {
            let dObj = new Date(p[0], p[1] - 1, p[2]);
            dataCabecalho = `${dias[dObj.getDay()]} (${p[2]}/${p[1]}/${p[0]})`;
        } else {
            dataCabecalho = dadosAtuais[0].dataIni;
        }
    } else {
        const hoje = new Date();
        dataCabecalho = `${dias[hoje.getDay()]} (${String(hoje.getDate()).padStart(2, '0')}/${String(hoje.getMonth() + 1).padStart(2, '0')}/${hoje.getFullYear()})`;
    }

    let bgStyle = configPDF.bgImage ? `background-image: url('${configPDF.bgImage}'); background-color: transparent;` : `background-image: none; background-color: #000;`;
    
    // 3. Loop de criação de todas as páginas
    for (let p = 0; p < totalPaginas; p++) {
        let inicio = p * linhasPorPagina;
        let fim = inicio + linhasPorPagina;
        let linhasPagina = dadosAtuais.slice(inicio, fim);

        let linhasHTML = '';
        linhasPagina.forEach(row => {
            let horario = row.horaIni ? row.horaIni.substring(0,5) : "";
            let isAusente = (row.salaSugerida === "FALTOU" || row.salaSugerida === "AUSENTE");
            let bgLinha = isAusente ? "background: #ef4444; color: #ffffff;" : `background: transparent; color: ${cCorpo};`;
            let corBorda = isAusente ? "#ffffff" : cCorpo;
            let textSala = isAusente ? "AUSENTE" : (row.salaSugerida || "");
            let fwSala = isAusente ? "font-weight: bold;" : "";
            let discCompleta = `${row.codDisc || ''} ${row.turma || ''} ${row.nomDisc || ''}`.trim();
            
            // Remove os emojis visuais dos nomes das salas para o PDF
            textSala = textSala.replace("🔒 Forçada", "").replace("⚠️ Proj. Ruim", "").trim();
            
            linhasHTML += `
            <tr style="${bgLinha}">
                <td style="border: 1px solid ${corBorda};">${row.curso || ''}</td>
                <td style="border: 1px solid ${corBorda};">${horario}</td>
                <td style="border: 1px solid ${corBorda};">${row.periodoAcad || ''}</td>
                <td style="border: 1px solid ${corBorda};">${discCompleta}</td>
                <td style="border: 1px solid ${corBorda};">${row.prof || ''}</td>
                <td style="border: 1px solid ${corBorda}; ${fwSala}">${textSala}</td>
            </tr>`;
        });

        let paginaHTML = `
        <div style="width: 100%; aspect-ratio: 16/9; overflow: hidden; background-size: 100% 100%; background-position: center; background-repeat: no-repeat; border: 1px solid #94a3b8; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); margin-bottom: 20px; position: relative; box-sizing: border-box; display: flex; flex-direction: column; ${bgStyle} padding-top: ${(t / 168.75 * 100)}%; padding-bottom: ${(b / 168.75 * 100)}%; padding-left: ${(l / 300 * 100)}%; padding-right: ${(r / 300 * 100)}%;">
            
            <div style="position: absolute; top: ${titTop}%; left: 0; width: 100%; padding: 0 ${titPadX}%; box-sizing: border-box; text-align: ${titPos}; z-index: 10;">
                <div style="font-weight: bold; line-height: 1.1; font-family: '${titFont}', sans-serif; font-size: ${(parseFloat(titSize) * 0.45)}pt; color: ${titCor}; text-transform: uppercase;">
                    ${titTexto}
                </div>
            </div>
            
            <div style="position: absolute; top: ${dTop}%; left: 0; width: 100%; padding: 0 ${dPadX}%; box-sizing: border-box; text-align: ${dAlign}; z-index: 10;">
                <div style="font-weight: bold; line-height: 1.1; text-transform: uppercase; font-family: '${dFont}', sans-serif; font-size: ${(parseFloat(dSize) * 0.45)}pt; color: ${dCor};">
                    ${dataCabecalho}
                </div>
            </div>
            
            <div style="width: 100%; height: 100%; display: flex; flex-direction: column;">
                <table style="width: 100%; border-collapse: collapse; font-family: 'Aptos Narrow', 'Arial Narrow', sans-serif; background: transparent; font-weight: bold; font-size: ${(parseFloat(fs) * 0.65)}pt;">
                    <thead style="color: ${cCab};">
                        <tr style="text-align: left; border-bottom: 2px solid ${cCab};">
                            <th style="border: 1px solid ${cCab};">CURSO</th>
                            <th style="border: 1px solid ${cCab};">HORÁRIO</th>
                            <th style="border: 1px solid ${cCab};">PERÍODO</th>
                            <th style="border: 1px solid ${cCab};">DISCIPLINA</th>
                            <th style="border: 1px solid ${cCab};">PROFESSOR (A)</th>
                            <th style="border: 1px solid ${cCab};">SALAS</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${linhasHTML}
                    </tbody>
                </table>
            </div>
        </div>`;
        
        container.insertAdjacentHTML('beforeend', paginaHTML);
    }
}
function salvarConfigPDF() {
    configPDF.mgTop = parseInt(document.getElementById('mgTop').value) || 0;
    configPDF.mgBottom = parseInt(document.getElementById('mgBottom').value) || 0;
    configPDF.mgLeft = parseInt(document.getElementById('mgLeft').value) || 0;
    configPDF.mgRight = parseInt(document.getElementById('mgRight').value) || 0;
    configPDF.fontSize = parseInt(document.getElementById('fontSizePDF').value) || 14;
    configPDF.linhasPorPagina = parseInt(document.getElementById('linhasPorPagina').value) || 18;
    configPDF.padCabecalho = parseInt(document.getElementById('padCabecalho').value) || 0;
    configPDF.padLinha = parseInt(document.getElementById('padLinha').value) || 0;

    if(document.getElementById('printTabFonte')) {
        configPDF.printTabFonte = parseInt(document.getElementById('printTabFonte').value) || 14;
        configPDF.printTabTop = parseInt(document.getElementById('printTabTop').value) || 0;
        configPDF.printTabLat = parseInt(document.getElementById('printTabLat').value) || 0;
        configPDF.printTitFonte = parseInt(document.getElementById('printTitFonte').value) || 32;
        configPDF.printTitTop = parseInt(document.getElementById('printTitTop').value) || 0;
        configPDF.printDataFonte = parseInt(document.getElementById('printDataFonte').value) || 24;
        configPDF.printDataTop = parseInt(document.getElementById('printDataTop').value) || 0;

        configPDF.wCol1 = parseInt(document.getElementById('wCol1').value) || 15;
        configPDF.wCol2 = parseInt(document.getElementById('wCol2').value) || 10;
        configPDF.wCol3 = parseInt(document.getElementById('wCol3').value) || 5;
        configPDF.wCol4 = parseInt(document.getElementById('wCol4').value) || 35;
        configPDF.wCol5 = parseInt(document.getElementById('wCol5').value) || 25;
        configPDF.wCol6 = parseInt(document.getElementById('wCol6').value) || 10;

    }
    
    configPDF.corCabecalho = document.getElementById('corFonteCabecalho').value;
    configPDF.corCorpo = document.getElementById('corFonteCorpo').value;

    configPDF.tituloTexto = document.getElementById('tituloTexto').value;
    configPDF.tituloPosicao = document.getElementById('tituloPosicao').value;
    configPDF.tituloTop = parseInt(document.getElementById('tituloTop').value) || 0;
    configPDF.tituloPaddingX = parseInt(document.getElementById('tituloPaddingX').value) || 0;
    configPDF.tituloFontSize = parseInt(document.getElementById('tituloFontSize').value) || 28;
    configPDF.tituloFontFamily = document.getElementById('tituloFontFamily').value;
    configPDF.tituloCor = document.getElementById('tituloCor').value;

    configPDF.dataAlign = document.getElementById('dataAlign').value;
    configPDF.dataTop = parseInt(document.getElementById('dataTop').value) || 0;
    configPDF.dataPaddingX = parseInt(document.getElementById('dataPaddingX').value) || 0;
    configPDF.dataFontFamily = document.getElementById('dataFontFamily').value;
    configPDF.dataFontSize = parseInt(document.getElementById('dataFontSize').value) || 20;
    configPDF.dataCor = document.getElementById('dataCor').value;
    
    salvarConfigPdfDB();
    alert("✅ Layout 16:9, Cores e Textos salvos com sucesso!");
    fecharConfigExportacao();
}

function salvarAjustesFinos() {
    if(document.getElementById('printTabFonte')) {
        configPDF.printTabFonte = parseInt(document.getElementById('printTabFonte').value) || 14;
        configPDF.printTabTop = parseInt(document.getElementById('printTabTop').value) || 0;
        configPDF.printTabLat = parseInt(document.getElementById('printTabLat').value) || 0;
        configPDF.printTitFonte = parseInt(document.getElementById('printTitFonte').value) || 32;
        configPDF.printTitTop = parseInt(document.getElementById('printTitTop').value) || 0;
        configPDF.printDataFonte = parseInt(document.getElementById('printDataFonte').value) || 24;
        configPDF.printDataTop = parseInt(document.getElementById('printDataTop').value) || 0;
    }
    
    // Salva silenciosamente no banco de dados
    salvarConfigPdfDB();
    
    // Exibe o aviso sem chamar a função de fechar a janela
    alert("✅ Configuração atualizada com sucesso!");
}

document.getElementById('inputBgPDF').addEventListener('change', function(e) {
    let file = e.target.files[0];
    if (!file) return;
    
    let reader = new FileReader();
    reader.onload = function(event) {
        configPDF.bgImage = event.target.result;
        atualizarPreviewPDF(); 
    };
    reader.readAsDataURL(file); 
});

// =========================================================
// ATALHOS DE TECLADO (FECHAR JANELAS COM 'ESC')
// =========================================================
document.addEventListener('keydown', function(event) {
    if (event.key === "Escape") {
        // 1. Fecha APENAS os modais flutuantes que estiverem abertos (equivalente a clicar no X)
        document.querySelectorAll('.modal-overlay-custom.active').forEach(modal => {
            modal.classList.remove('active');
        });
        
        // 2. Se estiver editando o nome de uma sala na janela de Configurações, cancela a edição
        if (typeof cancelarEdicaoSala === "function") cancelarEdicaoSala();
    }
});