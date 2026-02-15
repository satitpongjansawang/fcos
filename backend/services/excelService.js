const ExcelJS = require('exceljs');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

class ExcelService {
    async parsePoData(filePath) {
        const wb = XLSX.readFile(filePath);
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        if (data.length < 2) throw new Error('File has no data rows');
        const rows = data.slice(1).filter(r => r[0]);
        const doNumbers = [...new Set(rows.map(r => r[0]))];
        const customerCodes = [...new Set(rows.map(r => r[1]))];
        const dates = rows.map(r => r[8]).filter(Boolean);
        return { totalRows: rows.length, doNumbers: doNumbers.length, customerCodes: customerCodes.length, dateRange: dates.length ? { from: dates[0], to: dates[dates.length - 1] } : null, rows };
    }

    async generateIssueDO(sourcePath, revisionId) {
        const { rows } = await this.parsePoData(sourcePath);
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Issue D-O');
        const yellow = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
        const headers = ['No.', 'Inv. NO', 'DO No.', 'Date', 'Customer Parts No.', 'NGK Parts No.', 'Pcs.', 'PO NO.', 'Price', 'Ship To', 'Plan Code', 'Location', 'Contact Price No.', 'Privilege Flag', 'Period', 'Original Delivery Date', 'Marketing Suff/Mgt', 'Picking Route'];
        ws.addRow(headers);
        const headerRow = ws.getRow(1);
        headerRow.font = { bold: true, size: 10 };
        headerRow.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        headers.forEach((_, i) => { headerRow.getCell(i + 1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }; });
        [17, 18].forEach(i => { headerRow.getCell(i).fill = yellow; });

        rows.forEach((row, idx) => {
            const r = ws.addRow([
                idx + 1, row[0], row[0], row[8], row[4], row[3], row[5], row[6], row[7], row[9], row[10], row[11], row[15], row[14], row[13], row[12], '', ''
            ]);
            r.eachCell((cell) => { cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }; });
            r.getCell(17).fill = yellow;
            r.getCell(18).fill = yellow;
        });

        ws.columns.forEach(col => { col.width = 14; });
        const exportPath = path.join(__dirname, '../../exports', `issue_do_${revisionId}.xlsx`);
        const exportDir = path.dirname(exportPath);
        if (!fs.existsSync(exportDir)) fs.mkdirSync(exportDir, { recursive: true });
        await wb.xlsx.writeFile(exportPath);
        return exportPath;
    }

    async generateDeliveryDaily(sourcePath, revisionId) {
        const { rows } = await this.parsePoData(sourcePath);
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Delivery Daily');
        const yellow = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
        const headers = ['No.', 'Date', 'Customer Parts No.', 'NGK Parts No.', 'Pcs.', 'No. Lot or Production Date', 'Plan Code', 'Location', 'Packing Std.', 'Appearance of Package', 'Tag', 'Yes', 'No', 'Deliver Place', 'Yes', 'No', 'Yes', 'No', 'Remark'];
        ws.addRow(headers);
        const headerRow = ws.getRow(1);
        headerRow.font = { bold: true, size: 9 };
        headerRow.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        headers.forEach((_, i) => { headerRow.getCell(i + 1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }; });

        // Group by CUSTOMER_PARTS_NO + NITERRA_PARTS_NO + SHIP_TO
        const grouped = new Map();
        rows.forEach(row => {
            const key = `${row[4]}|${row[3]}|${row[9]}`;
            if (!grouped.has(key)) grouped.set(key, { custParts: row[4], niterraParts: row[3], shipTo: row[9], planCode: row[10], location: row[11], date: row[8], qty: 0 });
            grouped.get(key).qty += (parseFloat(row[5]) || 0);
        });

        let idx = 1;
        grouped.forEach((item) => {
            const r = ws.addRow([idx, item.date, item.custParts, item.niterraParts, item.qty, '', item.planCode, item.location, '', '', '', '', '', '', '', '', '', '', item.shipTo]);
            r.eachCell((cell) => { cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }; });
            [6, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18].forEach(i => { r.getCell(i).fill = yellow; });
            idx++;
        });

        // Total row
        const totalQty = [...grouped.values()].reduce((s, i) => s + i.qty, 0);
        const totalRow = ws.addRow(['', '', '', 'TOTAL', totalQty]);
        totalRow.font = { bold: true };
        totalRow.eachCell((cell) => { cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }; });

        ws.columns.forEach(col => { col.width = 13; });
        const exportPath = path.join(__dirname, '../../exports', `delivery_daily_${revisionId}.xlsx`);
        const exportDir = path.dirname(exportPath);
        if (!fs.existsSync(exportDir)) fs.mkdirSync(exportDir, { recursive: true });
        await wb.xlsx.writeFile(exportPath);
        return exportPath;
    }
}

module.exports = ExcelService;
