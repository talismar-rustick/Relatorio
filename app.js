document.getElementById('fileUpload').addEventListener('change', handleFile, false);
document.getElementById('fileUploadNew').addEventListener('change', handleFile, false);

// Listas globais para guardar os detalhes dos pedidos cancelados em memória
let listIfoodCanc = [];
let listAnotaCanc = [];

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
    // Resetando as listas a cada novo upload de planilha
    listIfoodCanc = [];
    listAnotaCanc = [];
    document.getElementById('cancelled-details').style.display = 'none';

    // Variáveis para pedidos FINALIZADOS
    let ifoodOrders = 0, ifoodFrete = 0, ifoodSubtotal = 0, ifoodTotal = 0;
    let anotaOrders = 0, anotaFrete = 0, anotaSubtotal = 0, anotaTotal = 0;

    // Variáveis para pedidos CANCELADOS
    let ifoodCancOrders = 0, ifoodCancFrete = 0, ifoodCancSubtotal = 0, ifoodCancTotal = 0;
    let anotaCancOrders = 0, anotaCancFrete = 0, anotaCancSubtotal = 0, anotaCancTotal = 0;

    data.forEach(row => {
        let status = row['Status'] || row['status'] || '';
        status = status.toString().trim().toLowerCase();
        
        // --- INÍCIO DAS REGRAS DE NEGÓCIO ---
        const statusFinalizados = ['finalizado', 'pronto pra entrega'];
        const statusCancelados = ['cancelado', 'rejeitado', 'pedido online expirado'];

        let isFinalizado = statusFinalizados.includes(status);
        let isCancelado = statusCancelados.includes(status);

        // Se o status não for nenhum dos informados na regra de negócio, ignora e pula para a próxima linha
        if (!isFinalizado && !isCancelado) {
            return; 
        }
        // --- FIM DAS REGRAS DE NEGÓCIO ---

        let origem = row['Origem'] || row['origem'] || '';
        origem = origem.toString().trim().toLowerCase();
        
        let frete = parseMoeda(row['Valor do frete']);
        let subtotal = parseMoeda(row['Subtotal (Total - Valor do frete)']);
        let total = parseMoeda(row['Valor total'] || row['Valor Total']);
        
        // Dados para a listagem
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
                // Salva o detalhe na memória
                listIfoodCanc.push({ data: dataPedido, numero: numPedido, valor: total });
            }
        } else if (['site', 'whatsapp', 'compartilhamento do menu'].includes(origem)) {
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
                // Salva o detalhe na memória
                listAnotaCanc.push({ data: dataPedido, numero: numPedido, valor: total });
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

// Funções dos botões de "Ver pedidos"
document.getElementById('btn-ifood-canc').addEventListener('click', () => showCancelledList('ifood'));
document.getElementById('btn-anota-canc').addEventListener('click', () => showCancelledList('anota'));

function showCancelledList(plataforma) {
    const container = document.getElementById('cancelled-details');
    const title = document.getElementById('cancelled-details-title');
    const tbody = document.getElementById('cancelled-details-body');
    
    // Limpa a tabela atual para preencher com a nova
    tbody.innerHTML = '';
    
    let listaAtual = [];

    if (plataforma === 'ifood') {
        listaAtual = listIfoodCanc;
        title.innerText = 'Lista de Pedidos Cancelados - iFood';
    } else {
        listaAtual = listAnotaCanc;
        title.innerText = 'Lista de Pedidos Cancelados - Anota AI';
    }

    const configMoeda = { style: 'currency', currency: 'BRL' };

    if (listaAtual.length === 0) {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td colspan="3" style="text-align: center; color: #7f8c8d;">Nenhum pedido cancelado para esta plataforma.</td>`;
        tbody.appendChild(tr);
    } else {
        listaAtual.forEach(item => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${item.data}</td>
                <td>${item.numero}</td>
                <td>${item.valor.toLocaleString('pt-BR', configMoeda)}</td>
            `;
            tbody.appendChild(tr);
        });
    }

    // Exibe a tabela no final da tela
    container.style.display = 'block';
}
