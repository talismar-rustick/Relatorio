document.getElementById('fileUpload').addEventListener('change', handleFile, false);
document.getElementById('fileUploadNew').addEventListener('change', handleFile, false);

// Listas globais
let listIfoodCanc = [];
let listAnotaCanc = [];

// Variáveis para controlar a ordenação
let currentPlatformView = '';
let currentSort = { column: '', desc: false };

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);
        processReport(json);
    };
    reader.readAsArrayBuffer(file);
}

function parseMoeda(valorRaw) {
    if (typeof valorRaw === 'number') return valorRaw;
    if (typeof valorRaw === 'string') {
        let limpo = valorRaw.replace(/[R$\s]/g, '').replace(/\./g, '').replace(',', '.');
        return parseFloat(limpo) || 0;
    }
    return 0;
}

function processReport(data) {
    listIfoodCanc = [];
    listAnotaCanc = [];
    document.getElementById('cancelled-details').style.display = 'none';

    let ifoodOrders = 0, ifoodFrete = 0, ifoodSubtotal = 0, ifoodTotal = 0;
    let anotaOrders = 0, anotaFrete = 0, anotaSubtotal = 0, anotaTotal = 0;

    let ifoodCancOrders = 0, ifoodCancFrete = 0, ifoodCancSubtotal = 0, ifoodCancTotal = 0;
    let anotaCancOrders = 0, anotaCancFrete = 0, anotaCancSubtotal = 0, anotaCancTotal = 0;

    data.forEach(row => {
        let statusOriginal = row['Status'] || row['status'] || '';
        let status = statusOriginal.toString().trim().toLowerCase();
        
        const statusFinalizados = ['finalizado', 'pronto pra entrega'];
        const statusCancelados = ['cancelado', 'rejeitado', 'pedido online expirado'];

        let isFinalizado = statusFinalizados.includes(status);
        let isCancelado = statusCancelados.includes(status);

        if (!isFinalizado && !isCancelado) {
            return; 
        }

        let origem = row['Origem'] || row['origem'] || '';
        origem = origem.toString().trim().toLowerCase();
        
        let frete = parseMoeda(row['Valor do frete']);
        let subtotal = parseMoeda(row['Subtotal (Total - Valor do frete)']);
        let total = parseMoeda(row['Valor total'] || row['Valor Total']);
        
        let dataPedido = row['Data'] || row['data'] || '-';
        let numPedido = row['Número do Pedido'] || row['numero do pedido'] || '-';

        if (origem === 'ifood') {
            if (isFinalizado) {
                ifoodOrders++;
                ifoodFrete += frete;
                ifoodSubtotal += subtotal;
                ifoodTotal += total;
            } else if (isCancelado) {
                ifoodCancOrders++;
                ifoodCancFrete += frete;
                ifoodCancSubtotal += subtotal;
                ifoodCancTotal += total;
                // Salva o detalhe incluindo o status original
                listIfoodCanc.push({ data: dataPedido, numero: numPedido, valor: total, motivo: statusOriginal });
            }
        } else { 
            if (isFinalizado) {
                anotaOrders++;
                anotaFrete += frete;
                anotaSubtotal += subtotal;
                anotaTotal += total;
            } else if (isCancelado) {
                anotaCancOrders++;
                anotaCancFrete += frete;
                anotaCancSubtotal += subtotal;
                anotaCancTotal += total;
                // Salva o detalhe incluindo o status original
                listAnotaCanc.push({ data: dataPedido, numero: numPedido, valor: total, motivo: statusOriginal });
            }
        }
    });

    renderDashboard(
        ifoodOrders, ifoodFrete, ifoodSubtotal, ifoodTotal,
        anotaOrders, anotaFrete, anotaSubtotal, anotaTotal,
        ifoodCancOrders, ifoodCancFrete, ifoodCancSubtotal, ifoodCancTotal,
        anotaCancOrders, anotaCancFrete, anotaCancSubtotal, anotaCancTotal
    );
}

function renderDashboard(
    ifoodOrders, ifoodFrete, ifoodSubtotal, ifoodTotal,
    anotaOrders, anotaFrete, anotaSubtotal, anotaTotal,
    ifoodCancOrders, ifoodCancFrete, ifoodCancSubtotal, ifoodCancTotal,
    anotaCancOrders, anotaCancFrete, anotaCancSubtotal, anotaCancTotal
) {
    document.getElementById('home-screen').style.display = 'none';
    document.getElementById('dashboard').style.display = 'block';

    const configMoeda = { style: 'currency', currency: 'BRL' };

    document.getElementById('ifood-orders').innerText = ifoodOrders;
    document.getElementById('ifood-frete').innerText = ifoodFrete.toLocaleString('pt-BR', configMoeda);
    document.getElementById('ifood-subtotal').innerText = ifoodSubtotal.toLocaleString('pt-BR', configMoeda);
    document.getElementById('ifood-total').innerText = ifoodTotal.toLocaleString('pt-BR', configMoeda);
    
    document.getElementById('anota-orders').innerText = anotaOrders;
    document.getElementById('anota-frete').innerText = anotaFrete.toLocaleString('pt-BR', configMoeda);
    document.getElementById('anota-subtotal').innerText = anotaSubtotal.toLocaleString('pt-BR', configMoeda);
    document.getElementById('anota-total').innerText = anotaTotal.toLocaleString('pt-BR', configMoeda);

    document.getElementById('ifood-canc-orders').innerText = ifoodCancOrders;
    document.getElementById('ifood-canc-frete').innerText = ifoodCancFrete.toLocaleString('pt-BR', configMoeda);
    document.getElementById('ifood-canc-subtotal').innerText = ifoodCancSubtotal.toLocaleString('pt-BR', configMoeda);
    document.getElementById('ifood-canc-total').innerText = ifoodCancTotal.toLocaleString('pt-BR', configMoeda);
    
    document.getElementById('anota-canc-orders').innerText = anotaCancOrders;
    document.getElementById('anota-canc-frete').innerText = anotaCancFrete.toLocaleString('pt-BR', configMoeda);
    document.getElementById('anota-canc-subtotal').innerText = anotaCancSubtotal.toLocaleString('pt-BR', configMoeda);
    document.getElementById('anota-canc-total').innerText = anotaCancTotal.toLocaleString('pt-BR', configMoeda);
}

// Funções dos botões e Ordenação da Tabela
document.getElementById('btn-ifood-canc').addEventListener('click', () => showCancelledList('ifood'));
document.getElementById('btn-anota-canc').addEventListener('click', () => showCancelledList('anota'));

function showCancelledList(plataforma) {
    const container = document.getElementById('cancelled-details');
    const title = document.getElementById('cancelled-details-title');
    
    // Reseta a ordenação se o usuário mudar a plataforma
    if (currentPlatformView !== plataforma) {
        currentSort = { column: '', desc: false };
        document.querySelectorAll('.sort-arrow').forEach(el => {
            el.style.transform = 'rotate(0deg)';
        });
    }
    
    currentPlatformView = plataforma;

    if (plataforma === 'ifood') {
        title.innerText = 'Lista de Pedidos Cancelados - iFood';
    } else {
        title.innerText = 'Lista de Pedidos Cancelados - Anota AI';
    }

    renderTable();
    container.style.display = 'block';
}

function sortTable(column) {
    // Se clicou na mesma coluna, inverte a ordem; senão, começa do padrão (crescente)
    if (currentSort.column === column) {
        currentSort.desc = !currentSort.desc;
    } else {
        currentSort.column = column;
        currentSort.desc = false;
    }

    // Rotaciona a seta correspondente e zera as outras
    document.querySelectorAll('.sort-arrow').forEach(el => {
        el.style.transform = 'rotate(0deg)';
    });
    const arrow = document.getElementById(`arrow-${column}`);
    if (arrow) {
        arrow.style.transform = currentSort.desc ? 'rotate(0deg)' : 'rotate(180deg)';
    }

    renderTable();
}

function renderTable() {
    const tbody = document.getElementById('cancelled-details-body');
    tbody.innerHTML = '';
    
    // Copia a lista para não bagunçar a ordem original
    let listaAtual = currentPlatformView === 'ifood' ? [...listIfoodCanc] : [...listAnotaCanc];

    // Lógica inteligente de ordenação
    if (currentSort.column) {
        listaAtual.sort((a, b) => {
            let valA = a[currentSort.column];
            let valB = b[currentSort.column];

            // Trata números (Valor e Número do pedido)
            if (currentSort.column === 'valor' || currentSort.column === 'numero') {
                valA = parseFloat(valA) || 0;
                valB = parseFloat(valB) || 0;
            } 
            // Trata a Data (converte de DD/MM/AAAA para um formato que o sistema consegue medir o tempo)
            else if (currentSort.column === 'data') {
                let partsA = valA.toString().split('/');
                let partsB = valB.toString().split('/');
                if (partsA.length === 3 && partsB.length === 3) {
                    valA = new Date(partsA[2], partsA[1] - 1, partsA[0]).getTime();
                    valB = new Date(partsB[2], partsB[1] - 1, partsB[0]).getTime();
                }
            } 
            // Trata texto puro (Status/Motivo)
            else {
                valA = valA.toString().toLowerCase();
                valB = valB.toString().toLowerCase();
            }

            if (valA < valB) return currentSort.desc ? 1 : -1;
            if (valA > valB) return currentSort.desc ? -1 : 1;
            return 0;
        });
    }

    const configMoeda = { style: 'currency', currency: 'BRL' };

    if (listaAtual.length === 0) {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td colspan="4" style="text-align: center; color: #7f8c8d;">Nenhum pedido cancelado para esta plataforma.</td>`;
        tbody.appendChild(tr);
    } else {
        listaAtual.forEach(item => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${item.data}</td>
                <td>${item.numero}</td>
                <td>${item.valor.toLocaleString('pt-BR', configMoeda)}</td>
                <td>${item.motivo}</td> `;
            tbody.appendChild(tr);
        });
    }
}
