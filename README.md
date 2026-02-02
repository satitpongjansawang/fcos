# FCOS - File Conversion & Order System

ระบบแปลงไฟล์ Excel สำหรับสร้างรายงาน Issue D/O และ Delivery Daily Report

## 📋 Features

- **Upload Excel File**: อัพโหลดไฟล์ PO Data (.xlsx, .xls)
- **Revision History**: เก็บประวัติไฟล์ที่อัพโหลดทั้งหมด
- **Export Reports**: 
  - Issue Delivery Order (Issue D/O)
  - Delivery Daily Report
- **Data Validation**: ตรวจสอบความถูกต้องของข้อมูล
- **Auto Merge**: รวมจำนวนสินค้าที่มี Parts No. ซ้ำกัน (สำหรับ Delivery Daily Report)

## 🚀 Getting Started

### Prerequisites

- Node.js >= 16.x
- npm >= 8.x

### Installation

```bash
# Clone repository
git clone https://github.com/satitpongjansawang/fcos.git
cd fcos

# Install dependencies
npm install

# Start server
npm start
```

Server will run at: http://localhost:3000

### Development Mode

```bash
npm run dev
```

## 📁 Project Structure

```
fcos/
├── backend/
│   ├── server.js           # Express server
│   ├── routes/
│   │   ├── upload.js       # Upload API
│   │   ├── export.js       # Export API
│   │   └── revision.js     # Revision History API
│   └── services/
│       └── excelService.js # Excel processing logic
├── frontend/
│   ├── index.html          # Main HTML
│   ├── css/
│   │   └── style.css       # Styles
│   └── js/
│       └── app.js          # Frontend JavaScript
├── uploads/                # Uploaded files storage
├── exports/                # Temporary export files
└── package.json
```

## 📊 Data Mapping

### Source File (PO Data)
ไฟล์ต้นทางที่รองรับมีคอลัมน์ดังนี้:

| Column | Field Name |
|--------|------------|
| A | DO NO |
| B | CUSTOMER CODE |
| C | BOX |
| D | NITERRA PARTS NO |
| E | CUSTOMER PARTS NO |
| F | QTY |
| G | PONO |
| H | PRICE |
| I | DELIVERY DATE |
| J | SHIP TO |
| K | PLAN CODE |
| L | LOCATION |
| M | ORIGINAL DELIVERY DATE |
| N | PERIOD |
| O | PRIVILEGE FLAG |
| P | CONTACT PRICE NO |
| AC | TEXT10 (Customer Name) |

### Issue D/O Report

| Field | Source | Status |
|-------|--------|--------|
| Inv. NO | DO NO | ✅ |
| DO No. | DO NO | ✅ |
| Picking Route | - | ❌ ไม่มี |
| CUSTOMER CODE | TEXT10 | ✅ |
| BOX | BOX | ✅ |
| NGK PARTS NO | NITERRA PARTS NO | ✅ |
| CUSTOMER PARTS NO | CUSTOMER PARTS NO | ✅ |
| QTY | QTY | ✅ |
| DELIVERY DATE | DELIVERY DATE | ✅ |
| PLAN CODE | PLAN CODE | ⚠️ อาจว่าง |
| LOCATION | LOCATION | ✅ |
| ORIGINAL DELIVERY DATE | ORIGINAL DELIVERY DATE | ✅ |
| PERIOD | PERIOD | ⚠️ อาจว่าง |
| PO NO | PONO | ✅ |
| PRICE | PRICE | ⚠️ อาจว่าง |
| SHIP TO | SHIP TO | ✅ |
| PRIVILEGE Flag | PRIVILEGE FLAG | ⚠️ อาจว่าง |
| CONTACT PRICE NO | CONTACT PRICE NO | ⚠️ อาจว่าง |

### Delivery Daily Report

| Field | Source | Status |
|-------|--------|--------|
| Customer | TEXT10 | ✅ |
| INV No. | DO NO | ✅ |
| Date | DELIVERY DATE | ✅ |
| Customer Parts No. | CUSTOMER PARTS NO | ✅ |
| NGK Parts No. | NITERRA PARTS NO | ✅ |
| Pcs. | QTY (merged) | ✅ |
| Location | LOCATION | ✅ |
| Remark | SHIP TO | ✅ |
| Marketing Suff/Mgt | - | ❌ ต้องกรอกเอง |
| Driver/Assistant/No. Car | - | ❌ ต้องกรอกเอง |
| Packaging Checking | - | ❌ ต้องกรอกเอง |
| Time Delivery | - | ❌ ต้องกรอกเอง |
| Result Delivery | - | ❌ ต้องกรอกเอง |
| Checker | - | ❌ ต้องกรอกเอง |
| Premium Freight | - | ❌ ต้องกรอกเอง |
| Logistic Mgt | - | ❌ ต้องกรอกเอง |

> ⚠️ ช่องสีเหลืองในไฟล์ที่ export = ฟิลด์ที่ต้องกรอกเพิ่มเติมเอง

## 🔧 API Endpoints

### Upload
```
POST /api/upload
Content-Type: multipart/form-data
Body: file (Excel file)
```

### Get Revisions
```
GET /api/revisions
```

### Get Revision by ID
```
GET /api/revisions/:id
```

### Delete Revision
```
DELETE /api/revisions/:id
```

### Export Issue D/O
```
GET /api/export/issue-do/:revisionId
```

### Export Delivery Daily Report
```
GET /api/export/delivery-daily/:revisionId
```

### Preview Data
```
GET /api/export/preview/:revisionId
```

## 📝 License

MIT

## 👤 Author

satitpongjansawang
