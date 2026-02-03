const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const { v4: uuidv4 } = require('uuid');

class ExcelService {
    async parsePoData(filePath) {
        try {
            const workbook = XLSX.readFile(filePath);
            if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                return { success: false, error: 'No worksheet found' };
            }
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            if (rawData.length < 2) {
                return { success: false, error: 'File has no data rows' };
            }
            const headers = rawData[0].map(h => h ? String(h).toUpperCase().replace(/\s+/g, '_') : '');
            const data = [];
            const customers = new Set();
            let deliveryDate = null;
            for (let i = 1; i < rawData.length; i++) {
                const row = rawData[i];
                if (!row || row.length === 0) continue;
                const rowData = {};
                headers.forEach((header, idx) => {
                    if (header && row[idx] !== undefined && row[idx] !== '') {
                        rowData[header] = row[idx];
                    }
                });
                if (Object.keys(rowData).length > 0) {
                    data.push(rowData);
                    const customerName = rowData.TEXT10 || rowData.CUSTOMER_CODE || '';
                    if (customerName) customers.add(String(customerName));
                    if (!deliveryDate && rowData.DELIVERY_DATE) {
                        deliveryDate = String(rowData.DELIVERY_DATE);
                    }
                }
            }
            return { success: true, recordCount: data.length, deliveryDate, customers: Array.from(customers), data };
        } catch (error) {
            return { success: false, error: 'Failed to parse: ' + error.message };
        }
    }

    async generateIssueDO(sourcePath, revisionId) {
        const parseResult = await this.parsePoData(sourcePath);
        if (!parseResult.success) throw new Error(parseResult.error);

        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet('Issue DO');

        // Title
        ws.mergeCells('A1:R1');
        ws.getCell('A1').value = 'Issue D/O';
        ws.getCell('A1').font = { bold: true, size: 14 };
        ws.getCell('A1').alignment = { horizontal: 'center' };

        ws.mergeCells('A2:R2');
        ws.getCell('A2').value = parseResult.deliveryDate || '';
        ws.getCell('A2').alignment = { horizontal: 'center' };

        // Headers
        const headers = ['Inv. NO', 'DO No.', 'Picking Route', 'CUSTOMER CODE', 'BOX', 'NGK PARTS NO', 'CUSTOMER PARTS NO', 'QTY', 'DELIVERY DATE', 'PLAN CODE', 'LOCATION', 'ORIGINAL DELIVERY DATE', 'PERIOD', 'PO NO', 'PRICE', 'SHIP TO', 'PRIVILEGE Flag', 'CONTACT PRICE NO'];
        const headerRow = ws.getRow(4);
        headers.forEach((h, i) => {
            const cell = headerRow.getCell(i + 1);
            cell.value = h;
            cell.font = { bold: true };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
            cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        });
        headerRow.height = 30;

        // Group data by customer and DO
        const grouped = {};
        parseResult.data.forEach(item => {
            const key = (item.TEXT10 || item.CUSTOMER_CODE || '') + '|' + (item.DO_NO || '');
            if (!grouped[key]) grouped[key] = { customer: item.TEXT10 || item.CUSTOMER_CODE, doNo: item.DO_NO, items: [] };
            grouped[key].items.push(item);
        });

        let currentRow = 5;
        let grandTotal = 0;
        const border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };

        Object.values(grouped).forEach(group => {
            let doTotal = 0;
            let firstItem = true;

            group.items.forEach(item => {
                const row = ws.getRow(currentRow);

                if (firstItem) {
                    row.getCell(1).value = group.doNo;
                    row.getCell(2).value = group.doNo;
                    row.getCell(4).value = group.customer;
                    firstItem = false;
                }

                row.getCell(5).value = item.BOX || '';
                row.getCell(6).value = item.NITERRA_PARTS_NO || '';
                row.getCell(7).value = item.CUSTOMER_PARTS_NO || '';
                const qty = parseInt(item.QTY) || 0;
                row.getCell(8).value = qty;
                doTotal += qty;
                row.getCell(9).value = item.DELIVERY_DATE || '';
                row.getCell(10).value = item.PLAN_CODE || '';
                row.getCell(11).value = item.LOCATION || '';
                row.getCell(12).value = item.ORIGINAL_DELIVERY_DATE || '';
                row.getCell(13).value = item.PERIOD || '';
                row.getCell(14).value = item.PONO || '';
                row.getCell(15).value = item.PRICE || '';
                row.getCell(16).value = item.SHIP_TO || '';
                row.getCell(17).value = item.PRIVILEGE_FLAG || '';
                row.getCell(18).value = item.CONTACT_PRICE_NO || '';

                for (let i = 1; i <= 18; i++) row.getCell(i).border = border;
                currentRow++;
            });

            // Total row
            const totalRow = ws.getRow(currentRow);
            totalRow.getCell(7).value = 'TOTAL';
            totalRow.getCell(7).font = { bold: true };
            totalRow.getCell(7).alignment = { horizontal: 'right' };
            totalRow.getCell(8).value = doTotal;
            totalRow.getCell(8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
            for (let i = 1; i <= 18; i++) totalRow.getCell(i).border = border;
            grandTotal += doTotal;
            currentRow++;
        });

        // Grand Total
        const gtRow = ws.getRow(currentRow);
        gtRow.getCell(7).value = 'GRAND TOTAL';
        gtRow.getCell(7).font = { bold: true };
        gtRow.getCell(8).value = grandTotal;
        gtRow.getCell(8).font = { bold: true };
        gtRow.getCell(8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
        for (let i = 1; i <= 18; i++) gtRow.getCell(i).border = border;

        // Column widths
        [12,12,12,24,6,22,20,10,12,10,10,15,8,15,10,8,12,15].forEach((w, i) => ws.getColumn(i+1).width = w);

        const exportDir = path.join(__dirname, '../../exports');
        if (!fs.existsSync(exportDir)) fs.mkdirSync(exportDir, { recursive: true });
        const exportPath = path.join(exportDir, 'issue_do_' + uuidv4() + '.xlsx');
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
        ws.getCell('N3').value = parseResult.deliveryDate || '';

        const row8Headers = { 2:'Customer', 3:'INV No.', 5:'Date', 8:'Customer Parts No.', 11:'NGK Parts No.', 13:'Pcs.', 16:'Plan Code', 17:'Location', 34:'Remark' };
        Object.entries(row8Headers).forEach(([col, val]) => {
            const cell = ws.getCell(8, parseInt(col));
            cell.value = val;
            cell.font = { bold: true };
        });

        const merged = {};
        parseResult.data.forEach(item => {
            const key = [item.TEXT10||item.CUSTOMER_CODE, item.DO_NO, item.CUSTOMER_PARTS_NO, item.NITERRA_PARTS_NO].join('|');
            if (!merged[key]) merged[key] = { ...item, qty: 0 };
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

        const exportDir = path.join(__dirname, '../../exports');
        if (!fs.existsSync(exportDir)) fs.mkdirSync(exportDir, { recursive: true });
        const exportPath = path.join(exportDir, 'delivery_daily_' + uuidv4() + '.xlsx');
        await workbook.xlsx.writeFile(exportPath);
        return exportPath;
    }

    async getPreviewData(sourcePath) {
        const parseResult = await this.parsePoData(sourcePath);
        if (!parseResult.success) throw new Error(parseResult.error);
        return { recordCount: parseResult.recordCount, deliveryDate: parseResult.deliveryDate, customers: parseResult.customers, sampleData: parseResult.data.slice(0, 10) };
    }
}
module.exports = ExcelService;
