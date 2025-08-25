// Global variables to store uploaded and processed data
let uploadedData = null;
let processedData = {
    AMD: null,
    INTEL: null,
    MICROSOFT: null
};

/**
 * Column mappings for each vendor export
 * Maps vendor-specific column names to source data column names
 */
const exportMappings = {
    AMD: {
        'Sold To Customer Name': 'Sold to Party name',
        'Sold to Customer #': 'Sold to Party ID',
        'Sold to Address Ln1': 'Sold to Party address',
        'Sold To City': 'Sold to Party city',
        'Sold to Postal': 'Sold to Party postal code',
        'Sold to Country': 'Sold to Party Country (ISO 2)',
        'Invoice #': 'Invoice number',
        'Invoice Line #': 'Invoice line number',
        'Invoice Date': 'Invoice date',
        'Incoming SKU': 'Manufacturer code long',
        'Target SKU': 'Manufacturer code',
        'Qty': 'Invoiced quantity'
    },
    INTEL: {
        'Sold To Customer Name': 'Sold to Party name',
        'Sold to Customer #': 'Sold to Party ID',
        'Sold to Address Ln1': 'Sold to Party address',
        'Sold To City': 'Sold to Party city',
        'Sold to Postal': 'Sold to Party postal code',
        'Sold to Country': 'Sold to Party Country (ISO 2)',
        'Invoice #': 'Invoice number',
        'Invoice Line #': 'Invoice line number',
        'Invoice Date': 'Invoice date',
        'Incoming SKU': 'Manufacturer code long',
        'Target SKU': 'Manufacturer code',
        'Qty': 'Invoiced quantity'
    },
    MICROSOFT: {
        'Invoice Number': 'Invoice number',
        'Invoice Date': 'Invoice date',
        'Reseller TPID': 'Sold to Party ID',
        'Reseller Name': 'Sold to Party name',
        'Reseller Country': 'Sold to Party Country (ISO 2)',
        'OEM Name': 'Product Hierarchy 1 Code',
        'OEM Device Model Name': 'Vendor SKU description',
        'OEM Device Model SKU': 'Manufacturer code',
        'Quantity Sold': 'Invoiced quantity',
        'Operating System': 'Operating system installed',
        'Price Per Unit': 'Unit price',
        'Currency': 'Currency code'
    }
};

/**
 * Defines the exact column order for each vendor's export file
 * These columns will appear in the exported Excel files in this order
 */
const exportColumns = {
    AMD: [
        'Partner ID (GSID)',
        'Sold To Customer Name',
        'Sold to Customer #',
        'Sold to Address Ln1',
        'Sold to Address Ln2',
        'Sold To City',
        'Sold To State',
        'Sold to Postal',
        'Sold to Country',
        'Sold to CPM ID',
        'Distributor Branch',
        'Invoice #',
        'Invoice Line #',
        'Invoice Date',
        'Incoming SKU',
        'Target SKU',
        'Qty',
        'Product Description'
    ],
    INTEL: [
        'Partner ID (GSID)',
        'Sold To Customer Name',
        'Sold to Customer #',
        'Sold to Address Ln1',
        'Sold to Address Ln2',
        'Sold To City',
        'Sold To State',
        'Sold to Postal',
        'Sold to Country',
        'Sold to CPM ID',
        'Distributor Branch',
        'Invoice #',
        'Invoice Line #',
        'Invoice Date',
        'Incoming SKU',
        'Target SKU',
        'Qty',
        'Product Description'
    ],
    MICROSOFT: [
        'Disti TPID',
        'Invoice Number',
        'Invoice Date',
        'Reseller TPID',
        'Reseller Name',
        'Reseller Country',
        'OEM Name',
        'OEM Device Model Name',
        'OEM Device Model SKU',
        'Quantity Sold',
        'Operating System',
        'Price Per Unit',
        'Currency',
        'End Customer Name',
        'Educational'
    ]
};

// ===== DRAG AND DROP SETUP =====
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
});

/**
 * Handles file selection from the file input element
 */
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
        handleFile(file);
    }
}

/**
 * Processes the uploaded Excel file
 * Validates file type and reads the data
 */
function handleFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/)) {
        showStatus('Please select a valid Excel file (.xlsx or .xls)', 'error');
        return;
    }

    showStatus('Processing file...', 'success');

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Get the first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Convert to JSON with raw values to preserve date strings
            uploadedData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, dateNF: 'dd/mm/yyyy hh:mm:ss' });
            
            if (uploadedData.length === 0) {
                showStatus('The Excel file appears to be empty', 'error');
                return;
            }

            processData();
            showStatus('File processed successfully! You can now download the exports.', 'success');
            document.getElementById('exports').style.display = 'block';
            
        } catch (error) {
            showStatus('Error processing file: ' + error.message, 'error');
        }
    };
    
    reader.readAsArrayBuffer(file);
}

/**
 * Main data processing function
 * Filters unwanted product categories and prepares data for each vendor format
 */
function processData() {
    if (!uploadedData || uploadedData.length === 0) return;
    
    // Get headers from first row
    const headers = uploadedData[0];
    const dataRows = uploadedData.slice(1);
    
    // Create mapping of header names to column indices (case-insensitive)
    const headerMap = {};
    headers.forEach((header, index) => {
        if (header && typeof header === 'string') {
            headerMap[header.toLowerCase().trim()] = index;
        }
    });
    
    // Product categories to exclude from exports (gaming peripherals, PC components, etc.)
    const excludeValues = [
        "Muismatten", "Voedingen", "Game controllers", "Videokaarten", 
        "Behuizingen", "Moederborden", "Desktop monitoren", "Processor koeling", 
        "Handheld Consoles", "Headsets", "Netwerkadapters", "Toetsenbord en muis sets", "Muizen", "Toetsenborden"
    ];
    
    const webHierarchyIndex = headerMap['web hierarchy description'];
    const filteredDataRows = dataRows.filter(row => {
        if (webHierarchyIndex !== undefined && webHierarchyIndex < row.length) {
            const webHierarchyValue = row[webHierarchyIndex];
            return !excludeValues.includes(webHierarchyValue);
        }
        return true; // Include row if Web Hierarchy Description column not found
    });
    
    // Process each export type
    Object.keys(exportColumns).forEach(exportType => {
        processedData[exportType] = processExportData(exportType, headerMap, filteredDataRows);
    });
}

/**
 * Processes data for a specific vendor export format
 * @param {string} exportType - The vendor type (AMD, INTEL, or MICROSOFT)
 * @param {object} headerMap - Mapping of column names to indices
 * @param {array} dataRows - Filtered data rows to process
 * @returns {array} Processed data ready for export
 */
function processExportData(exportType, headerMap, dataRows) {
    const requiredColumns = exportColumns[exportType];
    const mappings = exportMappings[exportType];
    const exportData = [];
    
    // Add headers as first row
    exportData.push(requiredColumns);
    
    // Create column mapping using predefined mappings
    const columnMapping = {};
    requiredColumns.forEach(exportCol => {
        const sourceColumn = mappings[exportCol];
        if (sourceColumn && headerMap.hasOwnProperty(sourceColumn.toLowerCase().trim())) {
            columnMapping[exportCol] = headerMap[sourceColumn.toLowerCase().trim()];
        } else {
            columnMapping[exportCol] = null; // Column not found or not mapped
        }
    });
    
    // For AMD/INTEL: Product Description combines multiple hardware specification columns
    let productDescriptionIndices = null;
    if ((exportType === 'AMD' || exportType === 'INTEL') && requiredColumns.includes('Product Description')) {
        productDescriptionIndices = {
            processorManufacturer: headerMap['processor manufacturer'] || null,
            processor: headerMap['processor model'] || null,
            discreteGraphics: headerMap['discrete graphics card model'] || null,
            onboardGraphics: headerMap['on-board graphics card model'] || null,
            operatingSystem: headerMap['operating system installed'] || null
        };
    }
    
    // Process data rows
    dataRows.forEach(row => {
        const exportRow = [];
        requiredColumns.forEach(colName => {
            // Set vendor-specific fixed values
            if ((exportType === 'AMD' || exportType === 'INTEL') && colName === 'Partner ID (GSID)') {
                exportRow.push('COPACO');
            } else if ((exportType === 'AMD' || exportType === 'INTEL') && colName === 'Distributor Branch') {
                exportRow.push('NL');
            } else if (exportType === 'MICROSOFT' && colName === 'Disti TPID') {
                exportRow.push('201286');
            } else if (colName === 'Product Description' && productDescriptionIndices) {
                // Build Product Description from hardware components
                const components = [];
                
                // Processor: Manufacturer + Model (e.g., "Intel Core i7-10700")
                let processorInfo = '';
                if (productDescriptionIndices.processorManufacturer !== null && row[productDescriptionIndices.processorManufacturer]) {
                    processorInfo += row[productDescriptionIndices.processorManufacturer];
                }
                if (productDescriptionIndices.processor !== null && row[productDescriptionIndices.processor]) {
                    processorInfo += (processorInfo ? ' ' : '') + row[productDescriptionIndices.processor];
                }
                if (processorInfo) {
                    components.push(processorInfo);
                }
                
                // Graphics cards and OS information
                if (productDescriptionIndices.discreteGraphics !== null && row[productDescriptionIndices.discreteGraphics]) {
                    if (row[productDescriptionIndices.discreteGraphics] === "Not available") {
                        components.push("No graphics card information found");
                    } else {
                        components.push(row[productDescriptionIndices.discreteGraphics]);
                    }
                }
                if (productDescriptionIndices.onboardGraphics !== null && row[productDescriptionIndices.onboardGraphics]) {
                    components.push(row[productDescriptionIndices.onboardGraphics]);
                }
                if (productDescriptionIndices.operatingSystem !== null && row[productDescriptionIndices.operatingSystem]) {
                    components.push(row[productDescriptionIndices.operatingSystem]);
                }
                exportRow.push(components.join(' / '));
            } else {
                const sourceIndex = columnMapping[colName];
                if (sourceIndex !== null && sourceIndex !== undefined && sourceIndex < row.length) {
                    let value = row[sourceIndex] || '';
                    
                    // Date formatting: Preserve dd/mm/yyyy hh:mm:ss format
                    if (colName === 'Invoice Date' && value) {
                        if (typeof value === 'string') {
                            exportRow.push(value);
                        } else if (typeof value === 'number') {
                            // Handle Excel serial date numbers
                            const excelEpoch = new Date(1900, 0, 1);
                            const daysSinceEpoch = value - 2; // Excel has a leap year bug for 1900
                            const date = new Date(excelEpoch.getTime() + daysSinceEpoch * 24 * 60 * 60 * 1000);
                            
                            const day = String(date.getDate()).padStart(2, '0');
                            const month = String(date.getMonth() + 1).padStart(2, '0');
                            const year = date.getFullYear();
                            const hours = String(date.getHours()).padStart(2, '0');
                            const minutes = String(date.getMinutes()).padStart(2, '0');
                            const seconds = String(date.getSeconds()).padStart(2, '0');
                            exportRow.push(`${day}/${month}/${year} ${hours}:${minutes}:${seconds}`);
                        } else if (value instanceof Date) {
                            // If it's a Date object, format it as dd/mm/yyyy hh:mm:ss
                            const day = String(value.getDate()).padStart(2, '0');
                            const month = String(value.getMonth() + 1).padStart(2, '0');
                            const year = value.getFullYear();
                            const hours = String(value.getHours()).padStart(2, '0');
                            const minutes = String(value.getMinutes()).padStart(2, '0');
                            const seconds = String(value.getSeconds()).padStart(2, '0');
                            exportRow.push(`${day}/${month}/${year} ${hours}:${minutes}:${seconds}`);
                        } else {
                            exportRow.push(value);
                        }
                    } else {
                        exportRow.push(value);
                    }
                } else {
                    exportRow.push(''); // Empty if column not found
                }
            }
        });
        exportData.push(exportRow);
    });
    
    return exportData;
}

/**
 * Downloads the processed data as an Excel file for the specified vendor
 * @param {string} exportType - The vendor type (AMD, INTEL, or MICROSOFT)
 */
function downloadExport(exportType) {
    if (!processedData[exportType]) {
        showStatus('No data available for ' + exportType + ' export', 'error');
        return;
    }
    
    // Create workbook and worksheet
    const ws = XLSX.utils.aoa_to_sheet(processedData[exportType]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, exportType + '_Export');
    
    // Generate filename
    const filename = exportType + '_export.xlsx';
    
    // Download file
    XLSX.writeFile(wb, filename);
    
    showStatus('Downloaded ' + filename, 'success');
}

/**
 * Displays status messages to the user
 * @param {string} message - The message to display
 * @param {string} type - Message type ('success' or 'error')
 */
function showStatus(message, type) {
    const statusDiv = document.getElementById('status');
    statusDiv.textContent = message;
    statusDiv.className = 'status ' + type;
    statusDiv.style.display = 'block';
    
    // Auto-hide success messages after 3 seconds
    if (type === 'success') {
        setTimeout(() => {
            statusDiv.style.display = 'none';
        }, 3000);
    }
}