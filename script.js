let tableData = [];
let headers = ['Formula', 'description', 'Kg', 'type', 'obs'];

// Função para carregar CSV
function loadCSV() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];

    if (file) {
        Papa.parse(file, {
            complete: function(results) {
                tableData = results.data;
                headers = tableData[0];
                renderTable();
            }
        });
    }
}

// Função para salvar CSV
function saveCSV() {
    const csv = Papa.unparse({
        fields: headers,
        data: tableData.slice(1)
    });
    
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    link.setAttribute('href', url);
    link.setAttribute('download', 'bd.csv');
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Função para pesquisar na tabela
function searchTable() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    const filteredData = tableData.filter((row, index) => {
        if (index === 0) return true; // Mantém o cabeçalho
        return row.some(cell => cell.toString().toLowerCase().includes(searchTerm));
    });
    renderTable(filteredData);
}

// Função para adicionar linha
function addRow() {
    const newRow = Array(headers.length).fill('');
    tableData.push(newRow);
    renderTable();
}

// Função para remover linha
function removeRow() {
    if (tableData.length > 1) {
        tableData.pop();
        renderTable();
    }
}

// Função para adicionar coluna
function addColumn() {
    headers.push('Nova Coluna');
    tableData.forEach(row => row.push(''));
    renderTable();
}

// Função para remover coluna
function removeColumn() {
    if (headers.length > 1) {
        headers.pop();
        tableData.forEach(row => row.pop());
        renderTable();
    }
}

// Função para renderizar a tabela
function renderTable(data = tableData) {
    const table = document.getElementById('csvTable');
    table.innerHTML = '';

    // Criar cabeçalho
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Criar corpo da tabela
    const tbody = document.createElement('tbody');
    data.slice(1).forEach(row => {
        const tr = document.createElement('tr');
        headers.forEach((header, index) => {
            const td = document.createElement('td');
            const input = document.createElement('input');
            input.type = 'text';
            input.value = row[index] || '';
            input.style.width = '100%';
            input.style.border = 'none';
            input.style.background = 'transparent';
            input.addEventListener('change', (e) => {
                row[index] = e.target.value;
            });
            td.appendChild(input);
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
}

// Inicializar a tabela vazia
renderTable(); 