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
                return { success: false, error: 'No worksheet found in the file' };
            }

            const data = [];
            const headers = {};
            
            const headerRow = worksheet.getRow(1);
            headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                const value = cell.value;
                if (value) {
                    headers[colNumber] = typeof value === 'string' ? value : String(value);
                }
            });

            if (Object.keys(headers).length === 0) {
                return { success: false, error: 'No headers found' };
            }

            const customers = new Set();
            let deliveryDate = null;

            worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                if (rowNumber === 1) return;
                
                const rowData = {};
                row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                    const header = headers[colNumber];
                    if (header) {
                        let cellValue = cell.value;
                        if (cellValue && typeof cellValue === 'object') {
                            if (cellValue.result !== undefined) cellValue = cellValue.result;
                            else if (cellValue.text !== undefined) cellValue = cellValue.text;
                            else if (cellValue instanceof Date) cellValue = cellValue.toISOString().split('T')[0];
                        }
                        rowData[header.toString().toUpperCase().replace(/\s+/g, '_')] = cellValue;
                    }
                });
                
                if (Object.keys(rowData).length > 0) {
                    data.push(rowData);
                    const customerName = rowData.TEXT10 || rowData.CUSTOMER_CODE || '';
                    if (customerName) customers.add(String(customerName));
                    if (!deliveryDate && rowData.DELIVERY_DATE) {
                        deliveryDate = rowData.DELIVERY_DATE instanceof Date 
                            ? rowData.DELIVERY_DATE.toISOString().split('T')[0] 
                            : String(rowData.DELIVERY_DATE);
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
            return { success: false, error: 'Failed to parse: ' + error.message };
        }
    }

    async generateIssueDO(sourcePath, revisionId) {
        const parseResult = await this.parsePoData(sourcePath);
        if (!parseResult.success) throw new Error(parseResult.error);

        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet('Issue DO');

        const headerStyle = {
            font: { bold: true, size: 10 },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } },
            border: { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} },
            alignment: { horizontal: 'center', vertical: 'middle', wrapText: true }
        };

        ws.mergeCells('A1:R1');
        ws.getCell('A1').value = 'Issue D/O';
        ws.getCell('A1').font = { bold: true, size: 14 };
        ws.getCell('A1').alignment = { horizontal: 'center' };

        const headers = ['Inv. NO', 'DO No.', 'Picking Route', 'CUSTOMER CODE', 'BOX',
            'NGK PARTS NO', 'CUSTOMER PARTS NO', 'QTY', 'DELIVERY DATE',
            'PLAN CODE', 'LOCATION', 'ORIGINAL DELIVERY DATE', 'PERIOD',
            'PO NO', 'PRICE', 'SHIP TO', 'PRIVILEGE Flag', 'CONTACT PRICE NO'];

        const headerRow = ws.getRow(4);
        headers.forEach((h, i) => {
            const cell = headerRow.getCell(i + 1);
            cell.value = h;
            cell.font = headerStyle.font;
            cell.fill = headerStyle.fill;
            cell.border = headerStyle.border;
            cell.alignment = headerStyle.alignment;
        });

        let currentRow = 5;
        let grandTotal = 0;
        const grouped = {};
        
        parseResult.data.forEach(item => {
            const key = (item.TEXT10 || item.CUSTOMER_CODE || '') + '|' + (item.DO_NO || '');
            if (!grouped[key]) grouped[key] = [];
            grouped[key].push(item);
        });

        Object.values(grouped).forEach(items => {
            let doTotal = 0;
            items.forEach((item, idx) => {
                const row = ws.getRow(currentRow);
                if (idx === 0) {
                    row.getCell(1).value = item.DO_NO;
                    row.getCell(2).value = item.DO_NO;
                    row.getCell(4).value = item.TEXT10 || item.CUSTOMER_CODE;
                }
                row.getCell(5).value = item.BOX;
                row.getCell(6).value = item.NITERRA_PARTS_NO;
                row.getCell(7).value = item.CUSTOMER_PARTS_NO;
                const qty = parseInt(item.QTY) || 0;
                row.getCell(8).value = qty;
                doTotal += qty;
                row.getCell(9).value = item.DELIVERY_DATE;
                row.getCell(10).value = item.PLAN_CODE;
                row.getCell(11).value = item.LOCATION;
                row.getCell(12).value = item.ORIGINAL_DELIVERY_DATE;
                row.getCell(14).value = item.PONO;
                row.getCell(15).value = item.PRICE;
                row.getCell(16).value = item.SHIP_TO;
                currentRow++;
            });
            const totalRow = ws.getRow(currentRow);
            totalRow.getCell(7).value = 'TOTAL';
            totalRow.getCell(8).value = doTotal;
            grandTotal += doTotal;
            currentRow++;
        });

        const exportPath = path.join(__dirname, '../../exports', 'issue_do_' + uuidv4() + '.xlsx');
        await workbook.xlsx.writeFile(exportPath);
        return exportPath;
    }

    async generateDeliveryDaily(sourcePath, revisionId) {
        const parseResult = await this.parsePoData(sourcePath);
        if (!parseResult.success) throw new Error(parseResult.error);

        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet('Delivery Daily Report');

        ws.getCell('N1').value = 'Delivery Daily Report';
        ws.getCell('N1').font = { bold: true, size: 14 };

        const row8Headers = { 2:'Customer', 3:'INV No.', 5:'Date', 8:'Customer Parts No.', 
            11:'NGK Parts No.', 13:'Pcs.', 16:'Plan Code', 17:'Location', 34:'Remark' };
        
        Object.entries(row8Headers).forEach(([col, val]) => {
            ws.getCell(8, parseInt(col)).value = val;
        });

        const merged = {};
        parseResult.data.forEach(item => {
            const key = [item.TEXT10||item.CUSTOMER_CODE, item.DO_NO, item.CUSTOMER_PARTS_NO, item.NITERRA_PARTS_NO].join('|');
            if (!merged[key]) {
                merged[key] = { ...item, qty: 0 };
            }
            merged[key].qty += parseInt(item.QTY) || 0;
        });

        let currentRow = 9;
        Object.values(merged).forEach(item => {
            const row = ws.getRow(currentRow);
            row.getCell(2).value = item.TEXT10 || item.CUSTOMER_CODE;
            row.getCell(3).value = item.DO_NO;
            row.getCell(5).value = item.DELIVERY_DATE;
            row.getCell(8).value = item.CUSTOMER_PARTS_NO;
            row.getCell(11).value = item.NITERRA_PARTS_NO;
            row.getCell(13).value = item.qty;
            row.getCell(16).value = item.PLAN_CODE;
            row.getCell(17).value = item.LOCATION;
            row.getCell(34).value = item.SHIP_TO;
            currentRow++;
        });

        const exportPath = path.join(__dirname, '../../exports', 'delivery_daily_' + uuidv4() + '.xlsx');
        await workbook.xlsx.writeFile(exportPath);
        return exportPath;
    }

    async getPreviewData(sourcePath) {
        const parseResult = await this.parsePoData(sourcePath);
        if (!parseResult.success) throw new Error(parseResult.error);
        return {
            recordCount: parseResult.recordCount,
            deliveryDate: parseResult.deliveryDate,
            customers: parseResult.customers,
            sampleData: parseResult.data.slice(0, 10)
        };
    }
}

module.exports = ExcelService;
