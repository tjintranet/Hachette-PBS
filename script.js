// Global variable to store processed data
let processedData = [];

document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('excelFile');
    fileInput.addEventListener('change', handleFileSelect);
});

// Handle file selection
async function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    try {
        const data = await readExcelFile(file);
        processData(data);
        updatePreviewTable();
        updateButtonStates();
        showStatus('File processed successfully', 'success');
    } catch (error) {
        showStatus(`Error processing file: ${error.message}`, 'danger');
        console.error(error);
    }
}

// Read Excel file
async function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { 
                    type: 'array',
                    cellDates: true,
                    dateNF: 'dd/mm/yyyy'
                });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, {
                    raw: true, // Changed to true to preserve numeric values
                    defval: ''
                });
                resolve(jsonData);
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Process Excel data
function processData(data) {
    if (!data || data.length === 0) return;

    // Generate a tracking reference for this batch
    const trackingRef = generateTrackingReference();
    
    // Find the quantity field name (case-insensitive)
    const getQuantityFieldName = (row) => {
        const possibleNames = ['Quantity', 'quantity', 'QUANTITY', 'Qty', 'qty', 'QTY'];
        return possibleNames.find(name => name in row);
    };

    processedData = data.map((row, index) => {
        // Set default date to 26112024 if not provided
        const defaultDate = '26112024';
        
        // Find the quantity field name
        const quantityField = getQuantityFieldName(row);
        
        // Handle quantity properly
        let quantity = '1'; // Default
        if (quantityField && row[quantityField] !== undefined && row[quantityField] !== null) {
            // Convert to number and then string to handle various formats
            // This ensures that 0 is preserved as "0" and not defaulted to "1"
            quantity = String(Number(row[quantityField]));
            
            // If conversion resulted in NaN, use default
            if (quantity === 'NaN') {
                quantity = '1';
            }
        }
        
        return {
            reference: row.Reference?.toString() || '',
            lineNumber: String(index + 1).padStart(5, '0'),
            isbn: row.ISBN?.toString() || '',
            date: defaultDate,
            courier: 'DPD',
            quantity: quantity,
            status: '1',
            trackingRef: trackingRef
        };
    });
}

// Generate tracking reference in the format %0SL30HE + numbers
function generateTrackingReference() {
    const prefix = '%0SL30HE1550';
    const randomNum = Math.floor(Math.random() * 9000000000000) + 1000000000000;
    return prefix + randomNum.toString();
}

// Update preview table
function updatePreviewTable() {
    const tbody = document.getElementById('previewBody');
    tbody.innerHTML = '';
    
    if (!processedData || processedData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="10" class="text-center">No data loaded</td></tr>';
        return;
    }
    
    processedData.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="checkbox" class="row-checkbox"></td>
            <td>
                <button class="btn btn-sm btn-danger" onclick="deleteRow(this)">
                    <i class="fas fa-trash"></i>
                </button>
            </td>
            <td>${row.reference}</td>
            <td>${row.lineNumber}</td>
            <td>${row.isbn}</td>
            <td>${row.date}</td>
            <td>${row.courier}</td>
            <td>${row.quantity}</td>
            <td>${row.status}</td>
            <td>${row.trackingRef}</td>
        `;
        tbody.appendChild(tr);
    });
}

// Download CSV file
function downloadCsv() {
    if (processedData.length === 0) {
        showStatus('No data to download', 'warning');
        return;
    }
    
    const orderNumber = processedData[0].reference;
    const csvRows = processedData.map(row => 
        [
            row.reference,
            row.lineNumber,
            row.isbn,
            row.date,
            row.courier,
            row.quantity,
            row.status,
            row.trackingRef
        ].join(',')
    );
    
    const csvContent = csvRows.join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    
    a.href = url;
    a.download = `T1.M${orderNumber}.PBS`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    
    showStatus('CSV downloaded successfully', 'success');
}

// Clear all data
function clearAll() {
    processedData = [];
    document.getElementById('excelFile').value = '';
    updatePreviewTable();
    updateButtonStates();
    showStatus('All data cleared', 'info');
}

// Update button states
function updateButtonStates() {
    const hasData = processedData && processedData.length > 0;
    document.getElementById('clearBtn').disabled = !hasData;
    document.getElementById('downloadBtn').disabled = !hasData;
    document.getElementById('deleteSelectedBtn').disabled = !hasData;
}

// Show status message
function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.className = `alert alert-${type}`;
    statusDiv.textContent = message;
    statusDiv.style.display = 'block';
    
    setTimeout(() => {
        statusDiv.style.display = 'none';
    }, 3000);
}

// Toggle all checkboxes
function toggleAllCheckboxes() {
    const selectAll = document.getElementById('selectAll');
    document.querySelectorAll('.row-checkbox').forEach(checkbox => {
        checkbox.checked = selectAll.checked;
    });
}

// Delete selected rows
function deleteSelected() {
    const checkboxes = document.querySelectorAll('.row-checkbox:checked');
    if (checkboxes.length === 0) {
        showStatus('No rows selected', 'warning');
        return;
    }
    
    const rowsToDelete = Array.from(checkboxes).map(checkbox => 
        checkbox.closest('tr'));
    
    rowsToDelete.forEach(row => {
        const index = Array.from(row.parentNode.children).indexOf(row);
        processedData.splice(index, 1);
        row.remove();
    });
    
    // Update line numbers
    processedData = processedData.map((row, index) => ({
        ...row,
        lineNumber: String(index + 1).padStart(5, '0')
    }));
    
    updateButtonStates();
    showStatus(`Deleted ${checkboxes.length} row(s)`, 'success');
}

// Delete single row
function deleteRow(button) {
    const row = button.closest('tr');
    const rowIndex = Array.from(row.parentNode.children).indexOf(row);
    
    processedData.splice(rowIndex, 1);
    
    // Update line numbers
    processedData = processedData.map((row, index) => ({
        ...row,
        lineNumber: String(index + 1).padStart(5, '0')
    }));
    
    updatePreviewTable();
    updateButtonStates();
    showStatus('Row deleted', 'success');
}