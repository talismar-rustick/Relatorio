document.getElementById('fileUpload').addEventListener('change', handleFile, false);

let ordersChartInstance = null;
let revenueChartInstance = null;

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

// Função auxiliar para tratar formatação de moeda da planilha
function parseMoeda(valorRaw) {
    if (typeof valorRaw === 'number') return valorRaw;
    if (typeof valorRaw === 'string') {
        let limpo = valorRaw.replace(/[R$\s]/g, '').replace(/\./g, '').replace(',', '.');
        return parseFloat(limpo) || 0;
    }
    return 0;
}

function processReport(data) {
    // Inicializadores iFood
    let ifoodOrders = 0;
    let ifoodFrete = 0;
    let ifoodSubtotal = 0;
    let ifoodTotal = 0;

    // Inicializadores Anota AI
    let anotaOrders = 0;
    let anotaFrete = 0;
    let anotaSubtotal = 0;
    let anotaTotal = 0;

    data.forEach(row => {
        let origem = row['Origem'] || row['origem'] || '';
        origem = origem.toString().trim().toLowerCase();
        
        // Mapeamento dos novos campos solicitados
        let frete = parseMoeda(row['Valor do frete']);
        let subtotal = parseMoeda(row['Subtotal (Total - Valor do frete)']);
        let total = parseMoeda(row['Valor total'] || row['Valor Total']);

        if (origem === 'ifood') {
            ifoodOrders++;
            ifoodFrete += frete;
            ifoodSubtotal += subtotal;
            ifoodTotal += total;
        } else if (['site', 'whatsapp', 'compartilhamento do menu'].includes(origem)) {
            anotaOrders++;
            anotaFrete += frete;
            anotaSubtotal += subtotal;
            anotaTotal += total;
        }
    });

    renderDashboard(
        ifoodOrders, ifoodFrete, ifoodSubtotal, ifoodTotal,
        anotaOrders, anotaFrete, anotaSubtotal, anotaTotal
    );
}

function renderDashboard(ifoodOrders, ifoodFrete, ifoodSubtotal, ifoodTotal, anotaOrders, anotaFrete, anotaSubtotal, anotaTotal) {
    document.getElementById('dashboard').style.display = 'block';

    // Formatação de moeda PT-BR
    const configMoeda = { style: 'currency', currency: 'BRL' };

    // Injeta dados do iFood na tela
    document.getElementById('ifood-orders').innerText = ifoodOrders;
    document.getElementById('ifood-frete').innerText = ifoodFrete.toLocaleString('pt-BR', configMoeda);
    document.getElementById('ifood-subtotal').innerText = ifoodSubtotal.toLocaleString('pt-BR', configMoeda);
    document.getElementById('ifood-total').innerText = ifoodTotal.toLocaleString('pt-BR', configMoeda);
    
    // Injeta dados da Anota AI na tela
    document.getElementById('anota-orders').innerText = anotaOrders;
    document.getElementById('anota-frete').innerText = anotaFrete.toLocaleString('pt-BR', configMoeda);
    document.getElementById('anota-subtotal').innerText = anotaSubtotal.toLocaleString('pt-BR', configMoeda);
    document.getElementById('anota-total').innerText = anotaTotal.toLocaleString('pt-BR', configMoeda);

    // Renderização dos Gráficos (Baseado no Faturamento Líquido/Total)
    const ctxOrders = document.getElementById('ordersChart').getContext('2d');
    const ctxRevenue = document.getElementById('revenueChart').getContext('2d');

    if (ordersChartInstance) ordersChartInstance.destroy();
    if (revenueChartInstance) revenueChartInstance.destroy();

    ordersChartInstance = new Chart(ctxOrders, {
        type: 'doughnut',
        data: {
            labels: ['iFood', 'Anota AI'],
            datasets: [{
                data: [ifoodOrders, anotaOrders],
                backgroundColor: ['#ea1d2c', '#00c853'],
                borderWidth: 0
            }]
        },
        options: {
            plugins: { title: { display: true, text: 'Divisão por Volume de Pedidos', font: {size: 15} } }
        }
    });

    revenueChartInstance = new Chart(ctxRevenue, {
        type: 'bar',
        data: {
            labels: ['iFood', 'Anota AI'],
            datasets: [{
                label: 'Faturamento Total',
                data: [ifoodTotal, anotaTotal],
                backgroundColor: ['#ea1d2c', '#00c853'],
                borderRadius: 6
            }]
        },
        options: {
            plugins: { 
                title: { display: true, text: 'Faturamento Total por Canal (R$)', font: {size: 15} },
                legend: { display: false } 
            }
        }
    });
}
