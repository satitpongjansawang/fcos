const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const { v4: uuidv4 } = require('uuid');

class ExcelService {
    
    // Parse PO Data file
    async parsePoData(filePath) {
        try {
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(filePath);
            
            const worksheet = workbook.worksheets[0];
            if (!worksheet) {
                return { success: false, error: 'No worksheet found' };
            }

            const data = [];
            const headers = {};
            
            // Get headers from first row
            worksheet.getRow(1).eachCell((cell, colNumber) => {
                headers[colNumber] = cell.value;
            });

            // Validate required columns
            const requiredColumns = ['DO NO', 'CUSTOMER CODE', 'NITERRA PARTS NO', 'CUSTOMER PARTS NO', 'QTY'];
            const headerValues = Object.values(headers).map(h => h ? h.toString().toUpperCase() : '');
            
            for (const col of requiredColumns) {
                if (!headerValues.some(h => h.includes(col.toUpperCase().replace(' ', '')))) {
                    // More flexible matching
                    const found = headerValues.some(h => {
                        const normalized = h.replace(/\s+/g, '');
                        const colNormalized = col.replace(/\s+/g, '');
                        return normalized.includes(colNormalized) || colNormalized.includes(normalized);
                    });
                    if (!found && col !== 'NITERRA PARTS NO') {
                        // NITERRA might be named differently
                    }
                }
            }

            // Parse data rows
            const customers = new Set();
            let deliveryDate = null;

            worksheet.eachRow((row, rowNumber) => {
                if (rowNumber === 1) return; // Skip header
                
                const rowData = {};
                row.eachCell((cell, colNumber) => {
                    const header = headers[colNumber];
                    if (header) {
                        rowData[header.toString().toUpperCase().replace(/\s+/g, '_')] = cell.value;
                    }
                });
                
                if (rowData.DO_NO || rowData.DONO) {
                    data.push(rowData);
                    
                    // Track customers
                    const customerName = rowData.TEXT10 || rowData.CUSTOMER_CODE || '';
                    if (customerName) customers.add(customerName);
                    
                    // Get delivery date
                    if (!deliveryDate && (rowData.DELIVERY_DATE || rowData.DELIVERYDATE)) {
                        deliveryDate = rowData.DELIVERY_DATE || rowData.DELIVERYDATE;
                    }
                }
            });

            return {
                success: true,
                recordCount: data.length,
                deliveryDate: deliveryDate,
                customers: Array.from(customers),
                data: data
            };

        } catch (error) {
            return { success: false, error: error.message };
        }
    }

    // Generate Issue D/O Report
    async generateIssueDO(sourcePath, revisionId) {
        const parseResult = await this.parsePoData(sourcePath);
        if (!parseResult.success) {
            throw new Error(parseResult.error);
        }

        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet('Issue DO');

        // Styles
        const headerStyle = {
            font: { bold: true, size: 10 },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            },
            alignment: { horizontal: 'center', vertical: 'middle', wrapText: true }
        };

        const dataStyle = {
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        };

        const totalStyle = {
            font: { bold: true },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        };

        // Title
        ws.mergeCells('A1:R1');
        ws.getCell('A1').value = 'Issue D/O';
        ws.getCell('A1').font = { bold: true, size: 14 };
        ws.getCell('A1').alignment = { horizontal: 'center' };

        // Date
        let deliveryDate = parseResult.deliveryDate;
        if (deliveryDate instanceof Date) {
            deliveryDate = deliveryDate.toLocaleDateString('en-GB');
        } else if (typeof deliveryDate === 'string' && deliveryDate.includes('-')) {
            const parts = deliveryDate.split('-');
            deliveryDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
        }
        
        ws.mergeCells('A2:R2');
        ws.getCell('A2').value = deliveryDate || '';
        ws.getCell('A2').alignment = { horizontal: 'center' };

        // Headers
        const headers = [
            'Inv. NO', 'DO No.', 'Picking Route', 'CUSTOMER CODE', 'BOX',
            'NGK PARTS NO', 'CUSTOMER PARTS NO', 'QTY', 'DELIVERY DATE',
            'PLAN CODE', 'LOCATION', 'ORIGINAL DELIVERY DATE', 'PERIOD',
            'PO NO', 'PRICE', 'SHIP TO', 'PRIVILEGE Flag', 'CONTACT PRICE NO'
        ];

        const headerRow = ws.getRow(4);
        headers.forEach((header, idx) => {
            const cell = headerRow.getCell(idx + 1);
            cell.value = header;
            Object.assign(cell, headerStyle);
        });
        headerRow.height = 30;

        // Group data by customer and DO
        const groupedData = {};
        parseResult.data.forEach(item => {
            const customer = item.TEXT10 || item.CUSTOMER_CODE || 'Unknown';
            const doNo = item.DO_NO || item.DONO || '';
            const key = `${customer}|${doNo}`;
            
            if (!groupedData[key]) {
                groupedData[key] = { customer, doNo, items: [] };
            }
            groupedData[key].items.push(item);
        });

        // Write data
        let currentRow = 5;
        let grandTotal = 0;

        Object.values(groupedData).forEach(group => {
            let doTotal = 0;
            let firstItemForDo = true;

            group.items.forEach(item => {
                const row = ws.getRow(currentRow);
                
                // Inv. NO & DO No.
                if (firstItemForDo) {
                    row.getCell(1).value = group.doNo;
                    row.getCell(2).value = group.doNo;
                    row.getCell(4).value = group.customer;
                    firstItemForDo = false;
                }

                // Picking Route - empty (not in source)
                row.getCell(3).value = '';
                
                // BOX
                row.getCell(5).value = item.BOX || '';
                
                // NGK Parts No
                row.getCell(6).value = item.NITERRA_PARTS_NO || item.NITERRAPARTSNO || '';
                
                // Customer Parts No
                row.getCell(7).value = item.CUSTOMER_PARTS_NO || item.CUSTOMERPARTSNO || '';
                
                // QTY
                const qty = parseInt(item.QTY) || 0;
                row.getCell(8).value = qty;
                doTotal += qty;
                
                // Delivery Date
                let delDate = item.DELIVERY_DATE || item.DELIVERYDATE || '';
                if (delDate instanceof Date) {
                    delDate = delDate.toLocaleDateString('en-GB');
                } else if (typeof delDate === 'string' && delDate.includes('-')) {
                    const parts = delDate.split('-');
                    delDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
                }
                row.getCell(9).value = delDate;
                
                // Plan Code
                row.getCell(10).value = item.PLAN_CODE || item.PLANCODE || '';
                
                // Location
                row.getCell(11).value = item.LOCATION || '';
                
                // Original Delivery Date
                let origDate = item.ORIGINAL_DELIVERY_DATE || item.ORIGINALDELIVERYDATE || '';
                if (origDate instanceof Date) {
                    origDate = origDate.toLocaleDateString('en-GB');
                } else if (typeof origDate === 'string' && origDate.includes('-')) {
                    const parts = origDate.split('-');
                    origDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
                }
                row.getCell(12).value = origDate;
                
                // Period
                row.getCell(13).value = item.PERIOD || '';
                
                // PO NO
                row.getCell(14).value = item.PONO || item.PO_NO || '';
                
                // Price
                row.getCell(15).value = item.PRICE || '';
                
                // Ship To
                row.getCell(16).value = item.SHIP_TO || item.SHIPTO || '';
                
                // Privilege Flag
                row.getCell(17).value = item.PRIVILEGE_FLAG || item.PRIVILEGEFLAG || '';
                
                // Contact Price No
                row.getCell(18).value = item.CONTACT_PRICE_NO || item.CONTACTPRICENO || '';

                // Apply styles
                for (let i = 1; i <= 18; i++) {
                    Object.assign(row.getCell(i), dataStyle);
                }

                currentRow++;
            });

            // Total row for each DO
            const totalRow = ws.getRow(currentRow);
            totalRow.getCell(7).value = 'TOTAL';
            totalRow.getCell(7).font = { bold: true };
            totalRow.getCell(7).alignment = { horizontal: 'right' };
            totalRow.getCell(8).value = doTotal;
            Object.assign(totalRow.getCell(8), totalStyle);
            
            for (let i = 1; i <= 18; i++) {
                Object.assign(totalRow.getCell(i), dataStyle);
            }
            
            grandTotal += doTotal;
            currentRow++;
        });

        // Grand Total
        const grandTotalRow = ws.getRow(currentRow);
        grandTotalRow.getCell(7).value = 'GRAND TOTAL';
        grandTotalRow.getCell(7).font = { bold: true };
        grandTotalRow.getCell(7).alignment = { horizontal: 'right' };
        grandTotalRow.getCell(8).value = grandTotal;
        Object.assign(grandTotalRow.getCell(8), totalStyle);
        grandTotalRow.getCell(8).font = { bold: true, size: 12 };

        for (let i = 1; i <= 18; i++) {
            Object.assign(grandTotalRow.getCell(i), dataStyle);
        }

        // Column widths
        const colWidths = [12, 12, 12, 24, 6, 22, 20, 10, 12, 10, 10, 15, 8, 15, 10, 8, 12, 15];
        colWidths.forEach((width, idx) => {
            ws.getColumn(idx + 1).width = width;
        });

        // Save file
        const exportDir = path.join(__dirname, '../../exports');
        if (!fs.existsSync(exportDir)) {
            fs.mkdirSync(exportDir, { recursive: true });
        }
        
        const exportPath = path.join(exportDir, `issue_do_${uuidv4()}.xlsx`);
        await workbook.xlsx.writeFile(exportPath);
        
        return exportPath;
    }

    // Generate Delivery Daily Report
    async generateDeliveryDaily(sourcePath, revisionId) {
        const parseResult = await this.parsePoData(sourcePath);
        if (!parseResult.success) {
            throw new Error(parseResult.error);
        }

        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet('Delivery Daily Report');

        // Styles
        const headerStyle = {
            font: { bold: true, size: 9 },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } },
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            },
            alignment: { horizontal: 'center', vertical: 'middle', wrapText: true }
        };

        const dataStyle = {
            border: {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            }
        };

        const yellowFill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFFF00' }
        };

        // Title
        ws.getCell('N1').value = 'Delivery Daily Report';
        ws.getCell('N1').font = { bold: true, size: 14 };
        ws.getCell('N1').alignment = { horizontal: 'center' };

        // Date
        let deliveryDate = parseResult.deliveryDate;
        if (deliveryDate instanceof Date) {
            deliveryDate = deliveryDate.toLocaleDateString('en-GB');
        } else if (typeof deliveryDate === 'string' && deliveryDate.includes('-')) {
            const parts = deliveryDate.split('-');
            deliveryDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
        }
        
        ws.getCell('N3').value = deliveryDate || '';
        ws.getCell('N3').alignment = { horizontal: 'center' };

        // Headers Row 7
        const row7Headers = {
            2: 'Marketing\nSuff',
            4: 'Marketing\nMgt',
            7: 'Driver',
            9: 'Assistant',
            10: 'No. Car',
            30: 'Checker',
            33: 'Logistic\nMgt'
        };

        // Headers Row 8
        const row8Headers = {
            2: 'Customer',
            3: 'INV No.',
            5: 'Date',
            8: 'Customer Parts No.',
            11: 'NGK Parts No.',
            13: 'Pcs.',
            15: 'Lot/Prod\nDate',
            16: 'Plan\nCode',
            17: 'Location',
            19: 'Packing\nStd.',
            20: 'Appearance\nof Package',
            21: 'Tag',
            23: 'Out',
            25: 'Yes',
            26: 'No',
            27: 'In',
            28: 'Deliver\nPlace',
            31: 'Yes',
            32: 'No',
            34: 'Remark'
        };

        // Write headers
        Object.entries(row7Headers).forEach(([col, val]) => {
            const cell = ws.getCell(7, parseInt(col));
            cell.value = val;
            Object.assign(cell, headerStyle);
        });

        Object.entries(row8Headers).forEach(([col, val]) => {
            const cell = ws.getCell(8, parseInt(col));
            cell.value = val;
            Object.assign(cell, headerStyle);
        });

        // Apply header style to all header cells
        for (let row = 7; row <= 8; row++) {
            for (let col = 1; col <= 34; col++) {
                const cell = ws.getCell(row, col);
                cell.fill = headerStyle.fill;
                cell.border = headerStyle.border;
            }
        }

        ws.getRow(7).height = 30;
        ws.getRow(8).height = 30;

        // Group and merge data by Customer Parts + NGK Parts
        const mergedData = {};
        parseResult.data.forEach(item => {
            const customer = item.TEXT10 || item.CUSTOMER_CODE || 'Unknown';
            const doNo = item.DO_NO || item.DONO || '';
            const customerParts = item.CUSTOMER_PARTS_NO || item.CUSTOMERPARTSNO || '';
            const ngkParts = item.NITERRA_PARTS_NO || item.NITERRAPARTSNO || '';
            const key = `${customer}|${doNo}|${customerParts}|${ngkParts}`;
            
            if (!mergedData[key]) {
                mergedData[key] = {
                    customer,
                    doNo,
                    customerParts,
                    ngkParts,
                    qty: 0,
                    deliveryDate: item.DELIVERY_DATE || item.DELIVERYDATE || '',
                    location: item.LOCATION || '',
                    planCode: item.PLAN_CODE || item.PLANCODE || '',
                    shipTo: item.SHIP_TO || item.SHIPTO || '',
                    lotNo: item.CUSTOMER_LOT_NO || item.CUSTOMERLOTNO || ''
                };
            }
            mergedData[key].qty += parseInt(item.QTY) || 0;
        });

        // Sort by customer and DO
        const sortedData = Object.values(mergedData).sort((a, b) => {
            if (a.customer !== b.customer) return a.customer.localeCompare(b.customer);
            return a.doNo.localeCompare(b.doNo);
        });

        // Write data
        let currentRow = 9;
        let lastCustomer = '';
        let lastDoNo = '';

        // Manual fill columns (to be filled by user)
        const manualCols = [2, 4, 7, 9, 10, 15, 19, 20, 21, 23, 25, 26, 27, 28, 30, 31, 32, 33];

        sortedData.forEach(item => {
            const row = ws.getRow(currentRow);

            // Customer (show only for first row of customer)
            if (item.customer !== lastCustomer) {
                row.getCell(2).value = item.customer.replace(/\s{2,}/g, '\n');
                row.getCell(2).alignment = { wrapText: true, vertical: 'top' };
                lastCustomer = item.customer;
                lastDoNo = '';
            }

            // INV No (show only for first row of DO)
            if (item.doNo !== lastDoNo) {
                row.getCell(3).value = item.doNo;
                lastDoNo = item.doNo;
            }

            // Date
            let delDate = item.deliveryDate;
            if (delDate instanceof Date) {
                delDate = delDate.toLocaleDateString('en-GB');
            } else if (typeof delDate === 'string' && delDate.includes('-')) {
                const parts = delDate.split('-');
                delDate = `${parts[2]}/${parts[1]}/${parts[0]}`;
            }
            row.getCell(5).value = delDate;

            // Customer Parts No
            row.getCell(8).value = item.customerParts;

            // NGK Parts No
            row.getCell(11).value = item.ngkParts;

            // Pcs (merged quantity)
            row.getCell(13).value = item.qty;

            // Plan Code
            row.getCell(16).value = item.planCode;

            // Location
            row.getCell(17).value = item.location;

            // Remark (Ship To)
            row.getCell(34).value = item.shipTo;

            // Apply borders and yellow fill for manual columns
            for (let col = 1; col <= 34; col++) {
                const cell = row.getCell(col);
                cell.border = dataStyle.border;
                
                if (manualCols.includes(col) && !cell.value) {
                    cell.fill = yellowFill;
                }
            }

            currentRow++;
        });

        // Column widths
        const colWidths = {
            1: 3, 2: 18, 3: 12, 4: 3, 5: 10, 6: 3, 7: 8, 8: 16,
            9: 3, 10: 3, 11: 28, 12: 3, 13: 8, 14: 3, 15: 10, 16: 8, 17: 8,
            18: 3, 19: 10, 20: 10, 21: 5, 22: 3, 23: 8, 24: 3, 25: 5, 26: 5,
            27: 8, 28: 8, 29: 3, 30: 8, 31: 5, 32: 5, 33: 10, 34: 6
        };
        
        Object.entries(colWidths).forEach(([col, width]) => {
            ws.getColumn(parseInt(col)).width = width;
        });

        // Save file
        const exportDir = path.join(__dirname, '../../exports');
        if (!fs.existsSync(exportDir)) {
            fs.mkdirSync(exportDir, { recursive: true });
        }
        
        const exportPath = path.join(exportDir, `delivery_daily_${uuidv4()}.xlsx`);
        await workbook.xlsx.writeFile(exportPath);
        
        return exportPath;
    }

    // Get preview data
    async getPreviewData(sourcePath) {
        const parseResult = await this.parsePoData(sourcePath);
        if (!parseResult.success) {
            throw new Error(parseResult.error);
        }

        // Return summary and first 10 records
        return {
            recordCount: parseResult.recordCount,
            deliveryDate: parseResult.deliveryDate,
            customers: parseResult.customers,
            sampleData: parseResult.data.slice(0, 10)
        };
    }
}

module.exports = ExcelService;
