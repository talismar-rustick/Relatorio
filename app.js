document.getElementById('fileUpload').addEventListener('change', handleFile, false);
document.getElementById('fileUploadNew').addEventListener('change', handleFile, false);

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
    // Variáveis para pedidos FINALIZADOS
    let ifoodOrders = 0, ifoodFrete = 0, ifoodSubtotal = 0, ifoodTotal = 0;
    let anotaOrders = 0, anotaFrete = 0, anotaSubtotal = 0, anotaTotal = 0;

    // Variáveis para pedidos CANCELADOS (Não finalizados)
    let ifoodCancOrders = 0, ifoodCancFrete = 0, ifoodCancSubtotal = 0, ifoodCancTotal = 0;
    let anotaCancOrders = 0, anotaCancFrete = 0, anotaCancSubtotal = 0, anotaCancTotal = 0;

    data.forEach(row => {
        let status = row['Status'] || row['status'] || '';
        status = status.toString().trim().toLowerCase();
        
        let origem = row['Origem'] || row['origem'] || '';
        origem = origem.toString().trim().toLowerCase();
        
        let frete = parseMoeda(row['Valor do frete']);
        let subtotal = parseMoeda(row['Subtotal (Total - Valor do frete)']);
        let total = parseMoeda(row['Valor total'] || row['Valor Total']);

        // Verifica se o status é finalizado
        let isFinalizado = (status === 'finalizado');

        if (origem === 'ifood') {
            if (isFinalizado) {
                ifoodOrders++;
                ifoodFrete += frete;
                ifoodSubtotal += subtotal;
                ifoodTotal += total;
            } else {
                ifoodCancOrders++;
                ifoodCancFrete += frete;
                ifoodCancSubtotal += subtotal;
                ifoodCancTotal += total;
            }
        } else if (['site', 'whatsapp', 'compartilhamento do menu'].includes(origem)) {
            if (isFinalizado) {
                anotaOrders++;
                anotaFrete += frete;
                anotaSubtotal += subtotal;
                anotaTotal += total;
            } else {
                anotaCancOrders++;
                anotaCancFrete += frete;
                anotaCancSubtotal += subtotal;
                anotaCancTotal += total;
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

    // Injeta dados dos FINALIZADOS
    document.getElementById('ifood-orders').innerText = ifoodOrders;
    document.getElementById('ifood-frete').innerText = ifoodFrete.toLocaleString('pt-BR', configMoeda);
    document.getElementById('ifood-subtotal').innerText = ifoodSubtotal.toLocaleString('pt-BR', configMoeda);
    document.getElementById('ifood-total').innerText = ifoodTotal.toLocaleString('pt-BR', configMoeda);
    
    document.getElementById('anota-orders').innerText = anotaOrders;
    document.getElementById('anota-frete').innerText = anotaFrete.toLocaleString('pt-BR', configMoeda);
    document.getElementById('anota-subtotal').innerText = anotaSubtotal.toLocaleString('pt-BR', configMoeda);
    document.getElementById('anota-total').innerText = anotaTotal.toLocaleString('pt-BR', configMoeda);

    // Injeta dados dos CANCELADOS (Não finalizados)
    document.getElementById('ifood-canc-orders').innerText = ifoodCancOrders;
    document.getElementById('ifood-canc-frete').innerText = ifoodCancFrete.toLocaleString('pt-BR', configMoeda);
    document.getElementById('ifood-canc-subtotal').innerText = ifoodCancSubtotal.toLocaleString('pt-BR', configMoeda);
    document.getElementById('ifood-canc-total').innerText = ifoodCancTotal.toLocaleString('pt-BR', configMoeda);
    
    document.getElementById('anota-canc-orders').innerText = anotaCancOrders;
    document.getElementById('anota-canc-frete').innerText = anotaCancFrete.toLocaleString('pt-BR', configMoeda);
    document.getElementById('anota-canc-subtotal').innerText = anotaCancSubtotal.toLocaleString('pt-BR', configMoeda);
    document.getElementById('anota-canc-total').innerText = anotaCancTotal.toLocaleString('pt-BR', configMoeda);
}
