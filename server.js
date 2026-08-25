const express = require('express');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

const app = express();
const PORT = process.env.PORT || 3000;
const EXCEL_FILE = path.join(__dirname, 'Revenue_Cost_Data.xlsx');
const OUTLETS_FILE = path.join(__dirname, 'outlets.json');

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname));

// ฟังก์ชันสำหรับอ่านรายชื่อห้องอาหารจากไฟล์ outlets.json (ถ้ายังไม่มีจะสร้างค่าเริ่มต้นให้)
function getOutlets() {
  const defaultOutlets = [
    'Reveira House',
    'Junsai',
    'Tapanyaki',
    'Terrace',
    'Juniper',
    'Marcele',
    'Canteen'
  ];
  if (!fs.existsSync(OUTLETS_FILE)) {
    fs.writeFileSync(OUTLETS_FILE, JSON.stringify(defaultOutlets, null, 2));
    return defaultOutlets;
  }
  try {
    const data = fs.readFileSync(OUTLETS_FILE, 'utf8');
    const parsed = JSON.parse(data);
    return Array.isArray(parsed) && parsed.length > 0 ? parsed : defaultOutlets;
  } catch (err) {
    return defaultOutlets;
  }
}

// API ดึงรายชื่อห้องอาหาร
app.get('/api/outlets', (req, res) => {
  try {
    const outlets = getOutlets();
    res.json({ success: true, outlets });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// API สำหรับเพิ่ม/แก้ไข/ลบ รายชื่อห้องอาหาร (บันทึกลงไฟล์ outlets.json)
app.post('/api/outlets/update', (req, res) => {
  try {
    const { outlets } = req.body;
    if (!Array.isArray(outlets)) {
      return res.status(400).json({ success: false, message: 'Invalid data format' });
    }
    fs.writeFileSync(OUTLETS_FILE, JSON.stringify(outlets, null, 2));
    res.json({ success: true, message: 'อัปเดตรายชื่อห้องอาหารสำเร็จ' });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.get('/', (req, res) => {
  const possiblePaths = [
    path.join(__dirname, 'public', 'index.html'),
    path.join(__dirname, 'public', 'index.htm'),
    path.join(__dirname, 'public', 'index'),
    path.join(__dirname, 'index.html'),
    path.join(__dirname, 'Revenue_Cost_Data.html')
  ];

  for (const p of possiblePaths) {
    if (fs.existsSync(p) && fs.statSync(p).size > 0) {
      return res.sendFile(p);
    }
  }

  res.send('<h3>⚠️ ไม่พบไฟล์ HTML หน้าแรก</h3>');
});

function getFormattedTimestamp() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
}

function getCellDateStr(cellValue) {
  if (!cellValue) return '';
  if (cellValue instanceof Date) {
    const year = cellValue.getFullYear();
    const month = String(cellValue.getMonth() + 1).padStart(2, '0');
    const day = String(cellValue.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  return String(cellValue).trim().substring(0, 10);
}

// ฟังก์ชันสำรองข้อมูลไฟล์ Excel อัตโนมัติ (เก็บไว้ในโฟลเดอร์ backups)
function backupExcelFile() {
  if (!fs.existsSync(EXCEL_FILE)) return;
  
  const backupDir = path.join(__dirname, 'backups');
  if (!fs.existsSync(backupDir)) {
    fs.mkdirSync(backupDir, { recursive: true });
  }

  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const timestampStr = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}_${pad(now.getHours())}-${pad(now.getMinutes())}-${pad(now.getSeconds())}`;
  
  const backupFileName = `Revenue_Cost_Data_backup_${timestampStr}.xlsx`;
  const backupFilePath = path.join(backupDir, backupFileName);

  try {
    fs.copyFileSync(EXCEL_FILE, backupFilePath);
    console.log(`✅ สำรองข้อมูลสำเร็จ: ${backupFileName}`);
  } catch (err) {
    console.error('❌ สำรองข้อมูลไม่สำเร็จ:', err.message);
  }
}

const DATA_COLUMNS = [
  { header: 'Date', key: 'date', width: 14 },
  { header: 'Outlet', key: 'outlet', width: 18 },
  { header: 'Categories', key: 'category', width: 16 },
  { header: 'Meal', key: 'meal', width: 14 },
  { header: 'Revenue', key: 'revenue', width: 16 },
  { header: 'Cover', key: 'cover', width: 12 },
  { header: 'Cost', key: 'cost', width: 16 },
  { header: 'Timestamp', key: 'timestamp', width: 22 }
];

const PUR_COLUMNS = [
  { header: 'Date', key: 'date', width: 14 },
  { header: 'Outlet', key: 'outlet', width: 18 },
  { header: 'Direct Purchase (THB)', key: 'purchase', width: 22 },
  { header: 'Store Issue (THB)', key: 'issue', width: 20 },
  { header: 'Timestamp', key: 'timestamp', width: 22 }
];

const COST_ADJ_COLUMNS = [
  { header: 'Date', key: 'date', width: 14 },
  { header: 'Outlet', key: 'outlet', width: 18 },
  { header: 'Transfer In (THB)', key: 'transferIn', width: 18 },
  { header: 'Transfer Out (THB)', key: 'transferOut', width: 18 },
  { header: 'Credit Cost (THB)', key: 'crCost', width: 18 },
  { header: 'Timestamp', key: 'timestamp', width: 22 }
];

const HOTEL_STATS_COLUMNS = [
  { header: 'Date', key: 'date', width: 14 },
  { header: 'Room Available', key: 'roomAvailable', width: 16 },
  { header: 'Room Sold', key: 'roomSold', width: 14 },
  { header: 'Com./HU. Room', key: 'comHuRoom', width: 16 },
  { header: 'Total Guests', key: 'totalGuest', width: 14 },
  { header: 'Weather', key: 'weather', width: 14 },
  { header: 'Timestamp', key: 'timestamp', width: 22 }
];

// โครงสร้างคอลัมน์สำหรับ Outlet_Summary Sheet
const SUMMARY_COLUMNS = [
  { header: 'Date', key: 'date', width: 14 },
  { header: 'Outlet', key: 'outlet', width: 18 },
  { header: 'Meal', key: 'meal', width: 14 },
  { header: 'Food Rev', key: 'foodRev', width: 14 },
  { header: 'Bev Rev', key: 'bevRev', width: 14 },
  { header: 'Total Rev', key: 'totalRev', width: 14 },
  { header: 'Cover', key: 'cover', width: 10 },
  { header: 'Food Cost', key: 'foodCost', width: 14 },
  { header: 'Bev Cost', key: 'bevCost', width: 14 },
  { header: 'Total Cost', key: 'totalCost', width: 14 },
  { header: 'Direct Purchase', key: 'purchase', width: 18 },
  { header: 'Store Issue', key: 'issue', width: 16 },
  { header: 'Transfer In', key: 'transferIn', width: 14 },
  { header: 'Transfer Out', key: 'transferOut', width: 14 },
  { header: 'Credit Cost', key: 'crCost', width: 14 },
  { header: 'Timestamp', key: 'timestamp', width: 22 }
];

// โครงสร้างคอลัมน์สำหรับ Pivot_Summary Sheet (Dynamic Summary / Pivot)
const PIVOT_SUMMARY_COLUMNS = [
  { header: 'Outlet', key: 'outlet', width: 20 },
  { header: 'Total Food Rev', key: 'totalFoodRev', width: 16 },
  { header: 'Total Bev Rev', key: 'totalBevRev', width: 16 },
  { header: 'Grand Total Rev', key: 'totalRev', width: 18 },
  { header: 'Total Cover', key: 'totalCover', width: 12 },
  { header: 'Total Food Cost', key: 'totalFoodCost', width: 16 },
  { header: 'Total Bev Cost', key: 'totalBevCost', width: 16 },
  { header: 'Total Cost', key: 'totalCost', width: 16 },
  { header: 'Direct Purchase', key: 'totalPurchase', width: 18 },
  { header: 'Store Issue', key: 'totalIssue', width: 16 },
  { header: 'Transfer Net', key: 'totalTransfer', width: 16 },
  { header: 'Credit Cost', key: 'totalCrCost', width: 16 }
];

function setupSheetColumns(sheet, columns) {
  sheet.columns = columns;
  sheet.getRow(1).font = { bold: true };
}

// ฟังก์ชันสร้าง/อัปเดต Dynamic Summary / Pivot Sheet ด้วยสูตร SUMIFS และเพิ่มแถว Grand Total ด้านล่าง
function updateDynamicPivotSheet(workbook) {
  let pivotSheet = workbook.getWorksheet('Pivot_Summary');
  if (!pivotSheet) {
    pivotSheet = workbook.addWorksheet('Pivot_Summary');
  }
  setupSheetColumns(pivotSheet, PIVOT_SUMMARY_COLUMNS);

  // ล้างข้อมูลเก่า (เว้นหัวตาราง)
  while (pivotSheet.rowCount > 1) {
    pivotSheet.spliceRows(2, 1);
  }

  const outlets = getOutlets();

  // 1. เพิ่มข้อมูลแต่ละ Outlet
  outlets.forEach((outName, index) => {
    const rowNum = index + 2;
    pivotSheet.addRow({
      outlet: outName,
      totalFoodRev: { formula: `SUMIFS(Outlet_Summary!D:D, Outlet_Summary!B:B, A${rowNum})` },
      totalBevRev: { formula: `SUMIFS(Outlet_Summary!E:E, Outlet_Summary!B:B, A${rowNum})` },
      totalRev: { formula: `SUMIFS(Outlet_Summary!F:F, Outlet_Summary!B:B, A${rowNum})` },
      totalCover: { formula: `SUMIFS(Outlet_Summary!G:G, Outlet_Summary!B:B, A${rowNum})` },
      totalFoodCost: { formula: `SUMIFS(Outlet_Summary!H:H, Outlet_Summary!B:B, A${rowNum})` },
      totalBevCost: { formula: `SUMIFS(Outlet_Summary!I:I, Outlet_Summary!B:B, A${rowNum})` },
      totalCost: { formula: `SUMIFS(Outlet_Summary!J:J, Outlet_Summary!B:B, A${rowNum})` },
      totalPurchase: { formula: `SUMIFS(Outlet_Summary!K:K, Outlet_Summary!B:B, A${rowNum})` },
      totalIssue: { formula: `SUMIFS(Outlet_Summary!L:L, Outlet_Summary!B:B, A${rowNum})` },
      totalTransfer: { formula: `SUMIFS(Outlet_Summary!M:M, Outlet_Summary!B:B, A${rowNum})-SUMIFS(Outlet_Summary!N:N, Outlet_Summary!B:B, A${rowNum})` },
      totalCrCost: { formula: `SUMIFS(Outlet_Summary!O:O, Outlet_Summary!B:B, A${rowNum})` }
    });
  });

  // 2. เพิ่มแถว Grand Total รวมทุกคอลัมน์ด้านล่าง
  const startRow = 2;
  const endRow = outlets.length + 1;
  const totalRow = pivotSheet.addRow({
    outlet: 'Grand Total',
    totalFoodRev: { formula: `SUM(B${startRow}:B${endRow})` },
    totalBevRev: { formula: `SUM(C${startRow}:C${endRow})` },
    totalRev: { formula: `SUM(D${startRow}:D${endRow})` },
    totalCover: { formula: `SUM(E${startRow}:E${endRow})` },
    totalFoodCost: { formula: `SUM(F${startRow}:F${endRow})` },
    totalBevCost: { formula: `SUM(G${startRow}:G${endRow})` },
    totalCost: { formula: `SUM(H${startRow}:H${endRow})` },
    totalPurchase: { formula: `SUM(I${startRow}:I${endRow})` },
    totalIssue: { formula: `SUM(J${startRow}:J${endRow})` },
    totalTransfer: { formula: `SUM(K${startRow}:K${endRow})` },
    totalCrCost: { formula: `SUM(L${startRow}:L${endRow})` }
  });

  totalRow.font = { bold: true };
}

async function initWorkbook() {
  const workbook = new ExcelJS.Workbook();
  if (fs.existsSync(EXCEL_FILE)) {
    await workbook.xlsx.readFile(EXCEL_FILE);
  }

  const ALL_OUTLETS = getOutlets();

  let dataSheet = workbook.getWorksheet('Data') || workbook.getWorksheet('Revenue_Cost');
  if (!dataSheet) {
    dataSheet = workbook.addWorksheet('Data');
  }
  setupSheetColumns(dataSheet, DATA_COLUMNS);

  ALL_OUTLETS.forEach(outletName => {
    let outletSheet = workbook.getWorksheet(outletName);
    if (!outletSheet) {
      outletSheet = workbook.addWorksheet(outletName);
    }
    setupSheetColumns(outletSheet, DATA_COLUMNS);
  });

  let purSheet = workbook.getWorksheet('Purchases_Issues');
  if (!purSheet) {
    purSheet = workbook.addWorksheet('Purchases_Issues');
  }
  setupSheetColumns(purSheet, PUR_COLUMNS);

  let costAdjSheet = workbook.getWorksheet('Cost_Adjustments');
  if (!costAdjSheet) {
    costAdjSheet = workbook.addWorksheet('Cost_Adjustments');
  }
  setupSheetColumns(costAdjSheet, COST_ADJ_COLUMNS);

  let statsSheet = workbook.getWorksheet('Hotel_Stats');
  if (!statsSheet) {
    statsSheet = workbook.addWorksheet('Hotel_Stats');
  }
  setupSheetColumns(statsSheet, HOTEL_STATS_COLUMNS);

  // เพิ่มชีท Outlet_Summary
  let summarySheet = workbook.getWorksheet('Outlet_Summary');
  if (!summarySheet) {
    summarySheet = workbook.addWorksheet('Outlet_Summary');
  }
  setupSheetColumns(summarySheet, SUMMARY_COLUMNS);

  // อัปเดต Dynamic Pivot Sheet พร้อมแถว Grand Total
  updateDynamicPivotSheet(workbook);

  return workbook;
}

function sortSheetByDate(sheet) {
  if (!sheet || sheet.rowCount <= 2) return;

  const rowsData = [];
  sheet.eachRow((row, rowNum) => {
    if (rowNum > 1) {
      rowsData.push(row.values);
    }
  });

  rowsData.sort((a, b) => String(a[1] || '').localeCompare(String(b[1] || '')));

  while (sheet.rowCount > 1) {
    sheet.spliceRows(2, 1);
  }

  rowsData.forEach(rowValues => {
    sheet.addRow(rowValues.slice(1));
  });
}

app.post('/api/save-excel', async (req, res) => {
  try {
    const { 
      date, outlet, purchase, issue, transferIn, transferOut, crCost,
      roomAvailable, roomSold, comHuRoom, totalGuest, weather, records 
    } = req.body;

    const workbook = await initWorkbook();
    const dataSheet = workbook.getWorksheet('Data') || workbook.getWorksheet('Revenue_Cost');
    const outletSheet = workbook.getWorksheet(outlet);
    const purSheet = workbook.getWorksheet('Purchases_Issues');
    const costAdjSheet = workbook.getWorksheet('Cost_Adjustments');
    const statsSheet = workbook.getWorksheet('Hotel_Stats');
    const summarySheet = workbook.getWorksheet('Outlet_Summary');
    const currentTimestamp = getFormattedTimestamp();

    if (dataSheet) {
      for (let i = dataSheet.rowCount; i >= 2; i--) {
        const row = dataSheet.getRow(i);
        const rDate = getCellDateStr(row.getCell(1).value);
        const rOutlet = String(row.getCell(2).value || '').trim();
        if (rDate === date && rOutlet === outlet) {
          dataSheet.spliceRows(i, 1);
        }
      }
    }

    if (outletSheet) {
      for (let i = outletSheet.rowCount; i >= 2; i--) {
        const row = outletSheet.getRow(i);
        const rDate = getCellDateStr(row.getCell(1).value);
        if (rDate === date) {
          outletSheet.spliceRows(i, 1);
        }
      }
    }

    if (records && records.length > 0) {
      records.forEach(r => {
        const rowObj = {
          date: r.date,
          outlet: r.outlet,
          category: r.category,
          meal: r.meal,
          revenue: Number(r.revenue) || 0,
          cover: Number(r.cover) || 0,
          cost: Number(r.cost) || 0,
          timestamp: currentTimestamp
        };

        if (dataSheet) dataSheet.addRow(rowObj);
        if (outletSheet) outletSheet.addRow(rowObj);
      });

      if (dataSheet) sortSheetByDate(dataSheet);
      if (outletSheet) sortSheetByDate(outletSheet);
    }

    const p = Number(purchase) || 0;
    const i = Number(issue) || 0;
    const tIn = Number(transferIn) || 0;
    const tOut = Number(transferOut) || 0;
    const cr = Number(crCost) || 0;

    if (purSheet) {
      for (let idx = purSheet.rowCount; idx >= 2; idx--) {
        const row = purSheet.getRow(idx);
        if (getCellDateStr(row.getCell(1).value) === date && String(row.getCell(2).value || '').trim() === outlet) {
          purSheet.spliceRows(idx, 1);
        }
      }
    }
    if ((p > 0 || i > 0) && purSheet) {
      purSheet.addRow({
        date: date,
        outlet: outlet,
        purchase: p,
        issue: i,
        timestamp: currentTimestamp
      });
      sortSheetByDate(purSheet);
    }

    if (costAdjSheet) {
      for (let idx = costAdjSheet.rowCount; idx >= 2; idx--) {
        const row = costAdjSheet.getRow(idx);
        if (getCellDateStr(row.getCell(1).value) === date && String(row.getCell(2).value || '').trim() === outlet) {
          costAdjSheet.spliceRows(idx, 1);
        }
      }
    }
    if ((tIn > 0 || tOut > 0 || cr > 0) && costAdjSheet) {
      costAdjSheet.addRow({
        date: date,
        outlet: outlet,
        transferIn: tIn,
        transferOut: tOut,
        crCost: cr,
        timestamp: currentTimestamp
      });
      sortSheetByDate(costAdjSheet);
    }

    // --- อัปเดตข้อมูลลงชีท Outlet_Summary ---
    if (summarySheet) {
      for (let idx = summarySheet.rowCount; idx >= 2; idx--) {
        const row = summarySheet.getRow(idx);
        if (getCellDateStr(row.getCell(1).value) === date && String(row.getCell(2).value || '').trim() === outlet) {
          summarySheet.spliceRows(idx, 1);
        }
      }

      const mealMap = {};
      if (records && records.length > 0) {
        records.forEach(r => {
          const m = r.meal || 'General';
          if (!mealMap[m]) {
            mealMap[m] = { foodRev: 0, bevRev: 0, cover: 0, foodCost: 0, bevCost: 0 };
          }
          const rev = Number(r.revenue) || 0;
          const cost = Number(r.cost) || 0;
          const cov = Number(r.cover) || 0;

          if (r.category === 'RV-Food') {
            mealMap[m].foodRev += rev;
            mealMap[m].foodCost += cost;
          } else {
            mealMap[m].bevRev += rev;
            mealMap[m].bevCost += cost;
          }
          mealMap[m].cover += cov;
        });
      }

      const mealsList = Object.keys(mealMap);
      if (mealsList.length > 0) {
        mealsList.forEach(m => {
          const item = mealMap[m];
          summarySheet.addRow({
            date: date,
            outlet: outlet,
            meal: m,
            foodRev: item.foodRev,
            bevRev: item.bevRev,
            totalRev: item.foodRev + item.bevRev,
            cover: item.cover,
            foodCost: item.foodCost,
            bevCost: item.bevCost,
            totalCost: item.foodCost + item.bevCost,
            purchase: p,
            issue: i,
            transferIn: tIn,
            transferOut: tOut,
            crCost: cr,
            timestamp: currentTimestamp
          });
        });
      } else if (p > 0 || i > 0 || tIn > 0 || tOut > 0 || cr > 0) {
        summarySheet.addRow({
          date: date,
          outlet: outlet,
          meal: '-',
          foodRev: 0,
          bevRev: 0,
          totalRev: 0,
          cover: 0,
          foodCost: 0,
          bevCost: 0,
          totalCost: 0,
          purchase: p,
          issue: i,
          transferIn: tIn,
          transferOut: tOut,
          crCost: cr,
          timestamp: currentTimestamp
        });
      }
      sortSheetByDate(summarySheet);
    }

    // อัปเดตสูตรและแถว Grand Total ใน Pivot_Summary Sheet
    updateDynamicPivotSheet(workbook);

    if (statsSheet && (Number(roomAvailable) > 0 || Number(roomSold) > 0 || Number(comHuRoom) > 0 || Number(totalGuest) > 0 || weather)) {
      let existingRow = null;
      statsSheet.eachRow((row, rowNum) => {
        if (rowNum > 1 && getCellDateStr(row.getCell(1).value) === date) {
          existingRow = row;
        }
      });

      if (existingRow) {
        if (Number(roomAvailable) > 0) existingRow.getCell(2).value = Number(roomAvailable);
        if (Number(roomSold) > 0) existingRow.getCell(3).value = Number(roomSold);
        if (Number(comHuRoom) > 0) existingRow.getCell(4).value = Number(comHuRoom);
        if (Number(totalGuest) > 0) existingRow.getCell(5).value = Number(totalGuest);
        if (weather) existingRow.getCell(6).value = weather;
        existingRow.getCell(7).value = currentTimestamp;
      } else {
        statsSheet.addRow({
          date: date,
          roomAvailable: Number(roomAvailable) || 0,
          roomSold: Number(roomSold) || 0,
          comHuRoom: Number(comHuRoom) || 0,
          totalGuest: Number(totalGuest) || 0,
          weather: weather || '',
          timestamp: currentTimestamp
        });
      }
      sortSheetByDate(statsSheet);
    }

    // ทำการสำรองข้อมูลอัตโนมัติก่อนเขียนไฟล์จริง
    backupExcelFile();

    await workbook.xlsx.writeFile(EXCEL_FILE);
    res.json({ success: true, message: 'บันทึก สำรองข้อมูล และซิงค์ทุกชีทสำเร็จเรียบร้อยแล้ว' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, message: 'เกิดข้อผิดพลาด: ' + err.message });
  }
});

app.post('/api/check-duplicate', async (req, res) => {
  try {
    const { date, outlet } = req.body;
    if (!fs.existsSync(EXCEL_FILE)) return res.json({ isDuplicate: false });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);
    const dataSheet = workbook.getWorksheet('Data') || workbook.getWorksheet('Revenue_Cost');

    let exists = false;
    if (dataSheet) {
      dataSheet.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
          const rowDate = getCellDateStr(row.getCell(1).value);
          const rowOutlet = String(row.getCell(2).value || '').trim();
          if (rowDate === date && rowOutlet === outlet) exists = true;
        }
      });
    }

    res.json({
      isDuplicate: exists,
      message: exists ? `ตรวจพบว่าเคยบันทึกข้อมูลวันที่ ${date} แผนก ${outlet} แล้ว` : ''
    });
  } catch (err) {
    res.json({ isDuplicate: false });
  }
});

app.get('/api/dashboard-summary', async (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_FILE)) {
      return res.json({
        success: true,
        data: {
          kpi: { 
            totalRevenue: 0, totalCost: 0, totalNetCost: 0, totalPurchase: 0, totalIssue: 0, 
            totalTransferIn: 0, totalTransferOut: 0, totalCrCost: 0, costPercentage: 0, 
            grossProfit: 0, totalCover: 0, totalRoomAvail: 0, totalRoomSold: 0, totalComHu: 0, 
            occupancyRate: 0, roomSoldToday: 0, roomSoldMtd: 0, guestToday: 0, guestMtd: 0, weather: '-',
            foodRevenue: 0, foodCostPct: 0, bevRevenue: 0, bevCostPct: 0
          },
          byOutlet: { labels: [], revenues: [], costs: [], netCosts: [], covers: [], purchases: [], issues: [], transferIns: [], transferOuts: [], crCosts: [] },
          byOutletFood: { labels: [], revenues: [], costs: [], covers: [] },
          byOutletBev: { labels: [], revenues: [], costs: [], covers: [] },
          canteen: { revenue: 0, netCost: 0, cover: 0, purchase: 0, issue: 0, transferIn: 0, transferOut: 0, crCost: 0, costPerDay: 0, costPerMeal: 0, costPct: '0.00%' },
          byCategory: { labels: ['RV-Food', 'RV-Beverage', 'RV-Wine', 'RV-Other'], revenues: [0, 0, 0, 0] },
          dailyTrend: { dates: [], revenues: [], costs: [] },
          dateRange: { min: '', max: '' }
        }
      });
    }

    const { startDate, endDate, outlet } = req.query;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);

    const dataSheet = workbook.getWorksheet('Data') || workbook.getWorksheet('Revenue_Cost');
    const purSheet = workbook.getWorksheet('Purchases_Issues');
    const costAdjSheet = workbook.getWorksheet('Cost_Adjustments');
    const statsSheet = workbook.getWorksheet('Hotel_Stats');

    let totalRevenue = 0, totalCost = 0, totalCover = 0;
    let foodRevenue = 0, foodCostRaw = 0, bevRevenue = 0, bevCostRaw = 0;
    let totalPurchase = 0, totalIssue = 0, totalTransferIn = 0, totalTransferOut = 0, totalCrCost = 0;
    let totalRoomAvail = 0, totalRoomSold = 0, totalComHu = 0;
    
    let roomSoldToday = 0, roomSoldMtd = 0;
    let guestToday = 0, guestMtd = 0;
    let todayWeather = '-';
    
    let canteenRev = 0, canteenCost = 0, canteenCover = 0, canteenPur = 0, canteenIss = 0, canteenTIn = 0, canteenTOut = 0, canteenCr = 0;

    const outletMap = {}, outletFoodMap = {}, outletBevMap = {}, catMap = { 'RV-Food': 0, 'RV-Beverage': 0, 'RV-Wine': 0, 'RV-Other': 0 }, dailyMap = {};
    let dates = [];

    const ALL_OUTLETS = getOutlets();
    const ALL_ACTIVE_OUTLETS = outlet && outlet !== 'ALL' ? [outlet] : ALL_OUTLETS;

    ALL_ACTIVE_OUTLETS.forEach(outName => {
      outletMap[outName] = { rev: 0, cost: 0, cov: 0, pur: 0, iss: 0, tIn: 0, tOut: 0, cr: 0 };
      outletFoodMap[outName] = { rev: 0, cost: 0, cov: 0 };
      outletBevMap[outName] = { rev: 0, cost: 0, cov: 0 };
    });

    if (dataSheet) {
      let colDate = 1, colOutlet = 2, colCat = 3, colRev = 5, colCover = 6, colCost = 7;

      const headerRow = dataSheet.getRow(1);
      headerRow.eachCell((cell, colNumber) => {
        const text = String(cell.value || '').trim().toLowerCase();
        if (text.startsWith('date')) colDate = colNumber;
        else if (text.startsWith('outlet')) colOutlet = colNumber;
        else if (text.includes('categor')) colCat = colNumber;
        else if (text.startsWith('revenue')) colRev = colNumber;
        else if (text.startsWith('cover')) colCover = colNumber;
        else if (text.startsWith('cost')) colCost = colNumber;
      });

      dataSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;

        const d = getCellDateStr(row.getCell(colDate).value);
        const out = String(row.getCell(colOutlet).value || '').trim();
        const cat = String(row.getCell(colCat).value || '').trim();
        const rev = Number(row.getCell(colRev).value) || 0;
        const cov = Number(row.getCell(colCover).value) || 0;
        const cost = Number(row.getCell(colCost).value) || 0;

        if (d) dates.push(d);

        if (startDate && d < startDate) return;
        if (endDate && d > endDate) return;
        if (outlet && outlet !== 'ALL' && out !== outlet) return;

        if (out === 'Canteen') {
          canteenRev += rev;
          canteenCost += cost;
          canteenCover += cov;
          return;
        }

        totalRevenue += rev;
        totalCost += cost;
        totalCover += cov;

        if (!outletMap[out]) {
          outletMap[out] = { rev: 0, cost: 0, cov: 0, pur: 0, iss: 0, tIn: 0, tOut: 0, cr: 0 };
        }
        if (!outletFoodMap[out]) outletFoodMap[out] = { rev: 0, cost: 0, cov: 0 };
        if (!outletBevMap[out]) outletBevMap[out] = { rev: 0, cost: 0, cov: 0 };

        outletMap[out].rev += rev;
        outletMap[out].cost += cost;
        outletMap[out].cov += cov;

        if (cat === 'RV-Food') {
          foodRevenue += rev;
          foodCostRaw += cost;
          outletFoodMap[out].rev += rev;
          outletFoodMap[out].cost += cost;
          outletFoodMap[out].cov += cov;
        } else {
          bevRevenue += rev;
          bevCostRaw += cost;
          outletBevMap[out].rev += rev;
          outletBevMap[out].cost += cost;
          outletBevMap[out].cov += cov;
        }

        if (catMap[cat] !== undefined) catMap[cat] += rev;

        if (!dailyMap[d]) dailyMap[d] = { rev: 0, cost: 0 };
        dailyMap[d].rev += rev;
        dailyMap[d].cost += cost;
      });
    }

    const purMap = {};
    if (purSheet) {
      purSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const d = getCellDateStr(row.getCell(1).value);
        const out = String(row.getCell(2).value || '').trim();
        const pur = Number(row.getCell(3).value) || 0;
        const iss = Number(row.getCell(4).value) || 0;

        if (startDate && d < startDate) return;
        if (endDate && d > endDate) return;
        if (outlet && outlet !== 'ALL' && out !== outlet) return;

        const key = `${d}_${out}`;
        purMap[key] = { out, pur, iss };
      });
    }

    Object.values(purMap).forEach(item => {
      if (item.out === 'Canteen') {
        canteenPur += item.pur;
        canteenIss += item.iss;
        return;
      }
      totalPurchase += item.pur;
      totalIssue += item.iss;
      if (!outletMap[item.out]) {
        outletMap[item.out] = { rev: 0, cost: 0, cov: 0, pur: 0, iss: 0, tIn: 0, tOut: 0, cr: 0 };
      }
      outletMap[item.out].pur += item.pur;
      outletMap[item.out].iss += item.iss;
    });

    const adjMap = {};
    if (costAdjSheet) {
      costAdjSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const d = getCellDateStr(row.getCell(1).value);
        const out = String(row.getCell(2).value || '').trim();
        const tIn = Number(row.getCell(3).value) || 0;
        const tOut = Number(row.getCell(4).value) || 0;
        const cr = Number(row.getCell(5).value) || 0;

        if (startDate && d < startDate) return;
        if (endDate && d > endDate) return;
        if (outlet && outlet !== 'ALL' && out !== outlet) return;

        const key = `${d}_${out}`;
        adjMap[key] = { out, tIn, tOut, cr };
      });
    }

    Object.values(adjMap).forEach(item => {
      if (item.out === 'Canteen') {
        canteenTIn += item.tIn;
        canteenTOut += item.tOut;
        canteenCr += item.cr;
        return;
      }
      totalTransferIn += item.tIn;
      totalTransferOut += item.tOut;
      totalCrCost += item.cr;
      if (!outletMap[item.out]) {
        outletMap[item.out] = { rev: 0, cost: 0, cov: 0, pur: 0, iss: 0, tIn: 0, tOut: 0, cr: 0 };
      }
      outletMap[item.out].tIn += item.tIn;
      outletMap[item.out].tOut += item.tOut;
      outletMap[item.out].cr += item.cr;
    });

    dates.sort();
    const latestDate = dates.length ? dates[dates.length - 1] : new Date().toISOString().split('T')[0];
    const targetTodayStr = endDate || latestDate;
    const currentMonthPrefix = targetTodayStr.substring(0, 7);

    if (statsSheet) {
      statsSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const d = getCellDateStr(row.getCell(1).value);
        const rAvail = Number(row.getCell(2).value) || 0;
        const rSold = Number(row.getCell(3).value) || 0;
        const cHu = Number(row.getCell(4).value) || 0;
        const guests = Number(row.getCell(5).value) || 0;
        const weatherVal = String(row.getCell(6).value || '').trim();

        if (startDate && d < startDate) return;
        if (endDate && d > endDate) return;

        totalRoomAvail += rAvail;
        totalRoomSold += rSold;
        totalComHu += cHu;

        if (d === targetTodayStr && weatherVal) todayWeather = weatherVal;
        if (d === targetTodayStr) { roomSoldToday += rSold; guestToday += guests; }
        if (d.startsWith(currentMonthPrefix) && d <= targetTodayStr) { roomSoldMtd += rSold; guestMtd += guests; }
      });
    }

    const occupancyRate = totalRoomAvail > 0 ? (totalRoomSold / totalRoomAvail) * 100 : 0;
    const totalNetCost = totalCost + totalPurchase + totalIssue + totalTransferIn - totalTransferOut - totalCrCost;
    const totalAdjustmentsNet = totalPurchase + totalIssue + totalTransferIn - totalTransferOut - totalCrCost;
    
    const foodNetCost = foodCostRaw + (totalAdjustmentsNet * (totalCost > 0 ? foodCostRaw / totalCost : 0.7));
    const bevNetCost = bevCostRaw + (totalAdjustmentsNet * (totalCost > 0 ? bevCostRaw / totalCost : 0.3));

    const foodCostPct = foodRevenue > 0 ? (foodNetCost / foodRevenue) * 100 : 0;
    const bevCostPct = bevRevenue > 0 ? (bevNetCost / bevRevenue) * 100 : 0;

    const canteenNetCost = canteenCost + canteenPur + canteenIss + canteenTIn - canteenTOut - canteenCr;
    const totalCoverCount = canteenCover > 0 ? canteenCover : 1;
    const canteenCostPerDay = canteenNetCost / totalCoverCount;
    const canteenCostPerMeal = canteenCostPerDay / 2;

    const sortedDaily = Object.keys(dailyMap).sort();
    const outletKeys = Object.keys(outletMap).filter(k => (outletMap[k].rev > 0 || outletMap[k].cost > 0 || outletMap[k].pur > 0 || outletMap[k].iss > 0) && k !== 'Canteen');

    const foodNetCostsArr = [];
    const bevNetCostsArr = [];

    outletKeys.forEach(k => {
      const o = outletMap[k];
      const fObj = outletFoodMap[k];
      const bObj = outletBevMap[k];
      const outletTotalRawCost = o.cost > 0 ? o.cost : 1;
      const outletAdjustmentsNet = o.pur + o.iss + o.tIn - o.tOut - o.cr;

      const fRaw = fObj.cost;
      const bRaw = bObj.cost;

      const fNet = fRaw + (outletAdjustmentsNet * (fRaw / outletTotalRawCost));
      const bNet = bRaw + (outletAdjustmentsNet * (bRaw / outletTotalRawCost));

      foodNetCostsArr.push(fNet);
      bevNetCostsArr.push(bNet);
    });

    res.json({
      success: true,
      data: {
        kpi: {
          totalRevenue,
          totalCost,
          totalNetCost,
          totalPurchase,
          totalIssue,
          totalTransferIn,
          totalTransferOut,
          totalCrCost,
          costPercentage: totalRevenue > 0 ? (totalNetCost / totalRevenue) * 100 : 0,
          grossProfit: totalRevenue - totalNetCost,
          totalCover,
          totalRoomAvail,
          totalRoomSold,
          totalComHu,
          occupancyRate,
          roomSoldToday,
          roomSoldMtd,
          guestToday,
          guestMtd,
          weather: todayWeather,
          foodRevenue,
          foodCostPct,
          bevRevenue,
          bevCostPct
        },
        byOutlet: {
          labels: outletKeys,
          revenues: outletKeys.map(k => outletMap[k].rev),
          costs: outletKeys.map(k => outletMap[k].cost),
          netCosts: outletKeys.map(k => {
            const o = outletMap[k];
            return o.cost + o.pur + o.iss + o.tIn - o.tOut - o.cr;
          }),
          covers: outletKeys.map(k => outletMap[k].cov),
          purchases: outletKeys.map(k => outletMap[k].pur),
          issues: outletKeys.map(k => outletMap[k].iss),
          transferIns: outletKeys.map(k => outletMap[k].tIn),
          transferOuts: outletKeys.map(k => outletMap[k].tOut),
          crCosts: outletKeys.map(k => outletMap[k].cr)
        },
        byOutletFood: {
          labels: outletKeys,
          revenues: outletKeys.map(k => outletFoodMap[k].rev),
          costs: foodNetCostsArr,
          covers: outletKeys.map(k => outletFoodMap[k].cov)
        },
        byOutletBev: {
          labels: outletKeys,
          revenues: outletKeys.map(k => outletBevMap[k].rev),
          costs: bevNetCostsArr,
          covers: outletKeys.map(k => outletBevMap[k].cov)
        },
        canteen: {
          revenue: canteenRev,
          netCost: canteenNetCost,
          cover: canteenCover,
          purchase: canteenPur,
          issue: canteenIss,
          transferIn: canteenTIn,
          transferOut: canteenTOut,
          crCost: canteenCr,
          costPerDay: canteenCostPerDay,
          costPerMeal: canteenCostPerMeal,
          costPct: canteenRev > 0 ? ((canteenNetCost / canteenRev) * 100).toFixed(2) + '%' : '0.00%'
        },
        byCategory: {
          labels: Object.keys(catMap),
          revenues: Object.values(catMap)
        },
        dailyTrend: {
          dates: sortedDaily,
          revenues: sortedDaily.map(d => dailyMap[d].rev),
          costs: sortedDaily.map(d => dailyMap[d].cost)
        },
        dateRange: {
          min: dates.length ? dates[0] : '',
          max: dates.length ? dates[dates.length - 1] : ''
        }
      }
    });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/delete-records', async (req, res) => {
  try {
    const { date, outlet } = req.body;
    if (!fs.existsSync(EXCEL_FILE)) return res.json({ success: false, message: 'ไม่พบไฟล์ฐานข้อมูล Excel' });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);

    const ALL_OUTLETS = getOutlets();
    const sheetsToCheck = ['Data', 'Revenue_Cost', 'Purchases_Issues', 'Cost_Adjustments', 'Hotel_Stats', 'Outlet_Summary', ...ALL_OUTLETS];

    sheetsToCheck.forEach(sheetName => {
      const sheet = workbook.getWorksheet(sheetName);
      if (sheet) {
        for (let i = sheet.rowCount; i >= 2; i--) {
          const row = sheet.getRow(i);
          const rDate = getCellDateStr(row.getCell(1).value);
          const rOutlet = String(row.getCell(2).value || '').trim();

          if (sheetName === 'Data' || sheetName === 'Revenue_Cost' || sheetName === 'Outlet_Summary') {
            if (rDate === date && (!outlet || outlet === 'ALL' || rOutlet === outlet)) {
              sheet.spliceRows(i, 1);
            }
          } 
          else if (ALL_OUTLETS.includes(sheetName)) {
            if (rDate === date && (!outlet || outlet === 'ALL' || outlet === sheetName)) {
              sheet.spliceRows(i, 1);
            }
          } 
          else {
            if (rDate === date && (!outlet || outlet === 'ALL' || !rOutlet || rOutlet === outlet)) {
              sheet.spliceRows(i, 1);
            }
          }
        }
      }
    });

    // อัปเดตสูตรใน Pivot_Summary หลังลบข้อมูล
    updateDynamicPivotSheet(workbook);

    // ทำการสำรองข้อมูลอัตโนมัติ
    backupExcelFile();

    await workbook.xlsx.writeFile(EXCEL_FILE);
    res.json({ success: true, message: `ลบข้อมูลวันที่ ${date} ของ ${outlet || 'All Outlets'} สำเร็จทุกชีท` });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.get('/api/open-excel', (req, res) => {
  if (!fs.existsSync(EXCEL_FILE)) return res.json({ success: false, message: 'ยังไม่มีไฟล์ Excel ถูกสร้างขึ้น' });

  const command = process.platform === 'win32' ? `start "" "${EXCEL_FILE}"` :
                  process.platform === 'darwin' ? `open "${EXCEL_FILE}"` : `xdg-open "${EXCEL_FILE}"`;

  exec(command, (err) => {
    if (err) return res.json({ success: false, message: 'ไม่สามารถเปิดไฟล์ได้: ' + err.message });
    res.json({ success: true, message: 'เปิดไฟล์สำเร็จ' });
  });
});

app.post('/api/refresh-excel', async (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_FILE)) {
      return res.json({ success: false, message: 'ไม่พบไฟล์ Excel ในระบบ' });
    }
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);
    res.json({ success: true, message: 'รีเฟรชและโหลดข้อมูลล่าสุดจาก Excel สำเร็จแล้ว' });
  } catch (err) {
    res.status(500).json({ success: false, message: 'เกิดข้อผิดพลาด: ' + err.message });
  }
});

app.post('/api/rebuild-excel', async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    const ALL_OUTLETS = getOutlets();

    const dataSheet = workbook.addWorksheet('Data');
    setupSheetColumns(dataSheet, DATA_COLUMNS);

    ALL_OUTLETS.forEach(outletName => {
      const s = workbook.addWorksheet(outletName);
      setupSheetColumns(s, DATA_COLUMNS);
    });

    const purSheet = workbook.addWorksheet('Purchases_Issues');
    setupSheetColumns(purSheet, PUR_COLUMNS);

    const costAdjSheet = workbook.addWorksheet('Cost_Adjustments');
    setupSheetColumns(costAdjSheet, COST_ADJ_COLUMNS);

    const statsSheet = workbook.addWorksheet('Hotel_Stats');
    setupSheetColumns(statsSheet, HOTEL_STATS_COLUMNS);

    const summarySheet = workbook.addWorksheet('Outlet_Summary');
    setupSheetColumns(summarySheet, SUMMARY_COLUMNS);

    updateDynamicPivotSheet(workbook);

    backupExcelFile();
    await workbook.xlsx.writeFile(EXCEL_FILE);
    res.json({ success: true, message: 'สร้างโครงสร้างไฟล์ Excel ใหม่สำเร็จ' });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/seed-random', async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    const ALL_OUTLETS = getOutlets();

    const dataSheet = workbook.addWorksheet('Data');
    setupSheetColumns(dataSheet, DATA_COLUMNS);

    const outletSheetsMap = {};
    ALL_OUTLETS.forEach(outName => {
      const s = workbook.addWorksheet(outName);
      setupSheetColumns(s, DATA_COLUMNS);
      outletSheetsMap[outName] = s;
    });

    const purSheet = workbook.addWorksheet('Purchases_Issues');
    setupSheetColumns(purSheet, PUR_COLUMNS);

    const costAdjSheet = workbook.addWorksheet('Cost_Adjustments');
    setupSheetColumns(costAdjSheet, COST_ADJ_COLUMNS);

    const statsSheet = workbook.addWorksheet('Hotel_Stats');
    setupSheetColumns(statsSheet, HOTEL_STATS_COLUMNS);

    const summarySheet = workbook.addWorksheet('Outlet_Summary');
    setupSheetColumns(summarySheet, SUMMARY_COLUMNS);

    const meals = ['ABF', 'Breakfast', 'Lunch', 'Dinner', 'Supper'];
    const types = [
      { type: 'Food', cat: 'RV-Food' },
      { type: 'Beverage', cat: 'RV-Beverage' },
      { type: 'Wine', cat: 'RV-Wine' },
      { type: 'Other', cat: 'RV-Other' }
    ];
    const weathers = ['Sunny', 'Cloudy', 'Rainy', 'Stormy'];

    const today = new Date();
    for (let i = 30; i >= 0; i--) {
      const d = new Date(today);
      d.setDate(d.getDate() - i);
      const dateStr = d.toISOString().split('T')[0];
      const randomTime = `${String(Math.floor(Math.random() * 12) + 10).padStart(2, '0')}:${String(Math.floor(Math.random() * 60)).padStart(2, '0')}:${String(Math.floor(Math.random() * 60)).padStart(2, '0')}`;
      const rowTimestamp = `${dateStr} ${randomTime}`;

      const roomAvail = 120;
      const roomSold = Math.floor(Math.random() * 40) + 70;
      const comHu = Math.floor(Math.random() * 10) + 2;
      const guests = Math.floor(roomSold * (1.6 + Math.random() * 0.4));
      const weather = weathers[Math.floor(Math.random() * weathers.length)];

      statsSheet.addRow({
        date: dateStr,
        roomAvailable: roomAvail,
        roomSold: roomSold,
        comHuRoom: comHu,
        totalGuest: guests,
        weather: weather,
        timestamp: `${dateStr} ${randomTime}`
      });

      ALL_OUTLETS.forEach(out => {
        const purVal = Math.floor(Math.random() * 8000) + 1000;
        const issVal = Math.floor(Math.random() * 5000) + 500;
        const tInVal = Math.floor(Math.random() * 800) + 50;
        const tOutVal = Math.floor(Math.random() * 500) + 20;
        const crVal = Math.floor(Math.random() * 1500) + 200;

        if (out === 'Canteen') {
          const canteenCover = Math.floor(Math.random() * 50) + 80;
          const canteenCost = canteenCover * (55 + Math.random() * 10);

          const rowData = {
            date: dateStr,
            outlet: 'Canteen',
            category: 'RV-Food',
            meal: 'Lunch',
            revenue: 0,
            cover: canteenCover,
            cost: Number(canteenCost.toFixed ? canteenCost.toFixed(2) : canteenCost),
            timestamp: rowTimestamp
          };

          dataSheet.addRow(rowData);
          if (outletSheetsMap['Canteen']) outletSheetsMap['Canteen'].addRow(rowData);

          summarySheet.addRow({
            date: dateStr,
            outlet: 'Canteen',
            meal: 'Lunch',
            foodRev: 0,
            bevRev: 0,
            totalRev: 0,
            cover: canteenCover,
            foodCost: rowData.cost,
            bevCost: 0,
            totalCost: rowData.cost,
            purchase: purVal,
            issue: issVal,
            transferIn: tInVal,
            transferOut: tOutVal,
            crCost: crVal,
            timestamp: rowTimestamp
          });

        } else {
          const mealMap = {};
          meals.forEach(m => {
            mealMap[m] = { foodRev: 0, bevRev: 0, cover: 0, foodCost: 0, bevCost: 0 };
            types.forEach(t => {
              const rev = Math.floor(Math.random() * 5000) + 500;
              const cov = Math.floor(Math.random() * 20) + 1;
              const cost = Math.floor(rev * (0.28 + Math.random() * 0.12));

              const rowData = {
                date: dateStr,
                outlet: out,
                category: t.cat,
                meal: m,
                revenue: rev,
                cover: cov,
                cost: cost,
                timestamp: rowTimestamp
              };

              dataSheet.addRow(rowData);
              if (outletSheetsMap[out]) outletSheetsMap[out].addRow(rowData);

              if (t.cat === 'RV-Food') {
                mealMap[m].foodRev += rev;
                mealMap[m].foodCost += cost;
              } else {
                mealMap[m].bevRev += rev;
                mealMap[m].bevCost += cost;
              }
              mealMap[m].cover += cov;
            });
          });

          meals.forEach(m => {
            const item = mealMap[m];
            summarySheet.addRow({
              date: dateStr,
              outlet: out,
              meal: m,
              foodRev: item.foodRev,
              bevRev: item.bevRev,
              totalRev: item.foodRev + item.bevRev,
              cover: item.cover,
              foodCost: item.foodCost,
              bevCost: item.bevCost,
              totalCost: item.foodCost + item.bevCost,
              purchase: purVal,
              issue: issVal,
              transferIn: tInVal,
              transferOut: tOutVal,
              crCost: crVal,
              timestamp: rowTimestamp
            });
          });
        }

        const randomTimePur = `${String(Math.floor(Math.random() * 12) + 10).padStart(2, '0')}:${String(Math.floor(Math.random() * 60)).padStart(2, '0')}:${String(Math.floor(Math.random() * 60)).padStart(2, '0')}`;
        purSheet.addRow({
          date: dateStr,
          outlet: out,
          purchase: purVal,
          issue: issVal,
          timestamp: `${dateStr} ${randomTimePur}`
        });

        costAdjSheet.addRow({
          date: dateStr,
          outlet: out,
          transferIn: tInVal,
          transferOut: tOutVal,
          crCost: crVal,
          timestamp: `${dateStr} ${randomTimePur}`
        });
      });
    }

    sortSheetByDate(statsSheet);
    sortSheetByDate(summarySheet);
    updateDynamicPivotSheet(workbook);

    backupExcelFile();
    await workbook.xlsx.writeFile(EXCEL_FILE);
    res.json({ success: true, message: 'สุ่มสร้างข้อมูล 30 วันสำเร็จ' });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/get-record-by-date', async (req, res) => {
  try {
    const { date, outlet } = req.body;
    if (!fs.existsSync(EXCEL_FILE)) return res.json({ success: false, data: null });

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);

    const dataSheet = workbook.getWorksheet('Data') || workbook.getWorksheet('Revenue_Cost');
    const purSheet = workbook.getWorksheet('Purchases_Issues');
    const costAdjSheet = workbook.getWorksheet('Cost_Adjustments');
    const statsSheet = workbook.getWorksheet('Hotel_Stats');

    const records = [];
    let purchase = 0, issue = 0, transferIn = 0, transferOut = 0, crCost = 0;
    let roomAvailable = 0, roomSold = 0, comHuRoom = 0, totalGuest = 0, weather = 'Sunny';

    if (dataSheet) {
      dataSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const rDate = getCellDateStr(row.getCell(1).value);
        const rOutlet = String(row.getCell(2).value || '').trim();
        if (rDate === date && rOutlet === outlet) {
          records.push({
            category: String(row.getCell(3).value || '').trim(),
            meal: String(row.getCell(4).value || '').trim(),
            revenue: Number(row.getCell(5).value) || 0,
            cover: Number(row.getCell(6).value) || 0,
            cost: Number(row.getCell(7).value) || 0
          });
        }
      });
    }

    if (purSheet) {
      purSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const rDate = getCellDateStr(row.getCell(1).value);
        const rOutlet = String(row.getCell(2).value || '').trim();
        if (rDate === date && rOutlet === outlet) {
          purchase = Number(row.getCell(3).value) || 0;
          issue = Number(row.getCell(4).value) || 0;
        }
      });
    }

    if (costAdjSheet) {
      costAdjSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const rDate = getCellDateStr(row.getCell(1).value);
        const rOutlet = String(row.getCell(2).value || '').trim();
        if (rDate === date && rOutlet === outlet) {
          transferIn = Number(row.getCell(3).value) || 0;
          transferOut = Number(row.getCell(4).value) || 0;
          crCost = Number(row.getCell(5).value) || 0;
        }
      });
    }

    if (statsSheet) {
      statsSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const rDate = getCellDateStr(row.getCell(1).value);
        if (rDate === date) {
          roomAvailable = Number(row.getCell(2).value) || 0;
          roomSold = Number(row.getCell(3).value) || 0;
          comHuRoom = Number(row.getCell(4).value) || 0;
          totalGuest = Number(row.getCell(5).value) || 0;
          weather = String(row.getCell(6).value || 'Sunny').trim();
        }
      });
    }

    res.json({
      success: true,
      data: {
        purchase, issue, transferIn, transferOut, crCost,
        roomAvailable, roomSold, comHuRoom, totalGuest, weather, records
      }
    });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.get('/api/shift-summary', async (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_FILE)) {
      return res.json({ success: true, data: { rows: [], chartData: { labels: [], datasets: [] }, dateRange: { min: '', max: '' } } });
    }

    const { startDate, endDate, mode } = req.query;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE);

    const dataSheet = workbook.getWorksheet('Data') || workbook.getWorksheet('Revenue_Cost');
    const purSheet = workbook.getWorksheet('Purchases_Issues');
    const costAdjSheet = workbook.getWorksheet('Cost_Adjustments');
    let dates = [];
    const aggregator = {};
    const outletTotals = {}; 
    const mealsList = ['ABF', 'Breakfast', 'Lunch', 'Dinner', 'Supper'];
    const mealColors = {
      'ABF': '#c5a059',
      'Breakfast': '#38bdf8',
      'Lunch': '#c084fc',
      'Dinner': '#e07a5f',
      'Supper': '#0d9488'
    };

    if (dataSheet) {
      dataSheet.eachRow((row, rowNum) => {
        if (rowNum === 1) return;
        const d = getCellDateStr(row.getCell(1).value);
        const out = String(row.getCell(2).value || '').trim();
        const cat = String(row.getCell(3).value || '').trim();
        const meal = String(row.getCell(4).value || '').trim();
        const rev = Number(row.getCell(5).value) || 0;
        const cov = Number(row.getCell(6).value) || 0;
        const cost = Number(row.getCell(7).value) || 0;

        if (d) dates.push(d);
        if (startDate && d < startDate) return;
        if (endDate && d > endDate) return;
        if (!out || out === 'Canteen') return;

        const key = `${out}_${meal}`;
        if (!aggregator[key]) {
          aggregator[key] = { 
            outlet: out, 
            meal: meal, 
            foodRev: 0, 
            bevRev: 0, 
            totalRev: 0,
            cover: 0, 
            rawFoodCost: 0, 
            rawBevCost: 0, 
            rawTotalCost: 0 
          };
        }

        aggregator[key].cover += cov;
        aggregator[key].rawTotalCost += cost;

        if (cat === 'RV-Food') {
          aggregator[key].foodRev += rev;
          aggregator[key].rawFoodCost += cost;
        } else {
          aggregator[key].bevRev += rev;
          aggregator[key].rawBevCost += cost;
        }
        aggregator[key].totalRev += rev;

        if (!outletTotals[out]) {
          outletTotals[out] = { totalRev: 0, totalRawCost: 0, adjustments: 0 };
        }
        outletTotals[out].totalRev += rev;
        outletTotals[out].totalRawCost += cost;
      });
    }

    if (mode === 'allocated') {
      const purMap = {};
      if (purSheet) {
        purSheet.eachRow((row, rowNum) => {
          if (rowNum === 1) return;
          const d = getCellDateStr(row.getCell(1).value);
          const out = String(row.getCell(2).value || '').trim();
          const pur = Number(row.getCell(3).value) || 0;
          const iss = Number(row.getCell(4).value) || 0;
          if (startDate && d < startDate) return;
          if (endDate && d > endDate) return;
          if (!out || out === 'Canteen') return;

          const key = `${d}_${out}`;
          purMap[key] = { out, pur, iss };
        });
      }
      Object.values(purMap).forEach(item => {
        if (!outletTotals[item.out]) outletTotals[item.out] = { totalRev: 0, totalRawCost: 0, adjustments: 0 };
        outletTotals[item.out].adjustments += (item.pur + item.iss);
      });

      const adjMap = {};
      if (costAdjSheet) {
        costAdjSheet.eachRow((row, rowNum) => {
          if (rowNum === 1) return;
          const d = getCellDateStr(row.getCell(1).value);
          const out = String(row.getCell(2).value || '').trim();
          const tIn = Number(row.getCell(3).value) || 0;
          const tOut = Number(row.getCell(4).value) || 0;
          const cr = Number(row.getCell(5).value) || 0;
          if (startDate && d < startDate) return;
          if (endDate && d > endDate) return;
          if (!out || out === 'Canteen') return;

          const key = `${d}_${out}`;
          adjMap[key] = { out, tIn, tOut, cr };
        });
      }
      Object.values(adjMap).forEach(item => {
        if (!outletTotals[item.out]) outletTotals[item.out] = { totalRev: 0, totalRawCost: 0, adjustments: 0 };
        outletTotals[item.out].adjustments += (item.tIn - item.tOut - item.cr);
      });
    }

    const rows = [];
    Object.keys(aggregator).forEach(key => {
      const item = aggregator[key];
      const outSummary = outletTotals[item.outlet] || { totalRev: 0, totalRawCost: 0, adjustments: 0 };
      
      let adjShare = 0;
      if (mode === 'allocated' && outSummary.totalRev > 0) {
        const revShare = item.totalRev / outSummary.totalRev;
        adjShare = outSummary.adjustments * revShare;
      }

      let foodCostFinal = item.rawFoodCost;
      let bevCostFinal = item.rawBevCost;

      if (mode === 'allocated' && item.rawTotalCost > 0) {
        const foodShare = item.rawFoodCost / item.rawTotalCost;
        const bevShare = item.rawBevCost / item.rawTotalCost;
        foodCostFinal += (adjShare * foodShare);
        bevCostFinal += (adjShare * bevShare);
      } else if (mode === 'allocated' && outSummary.rawTotalCost > 0) {
        const share = item.rawTotalCost / outSummary.rawTotalCost;
        foodCostFinal += (outSummary.adjustments * share * 0.7);
        bevCostFinal += (outSummary.adjustments * share * 0.3);
      }

      const totalCostFinal = foodCostFinal + bevCostFinal;

      rows.push({
        outlet: item.outlet,
        meal: item.meal,
        foodRev: item.foodRev,
        bevRev: item.bevRev,
        totalRev: item.totalRev,
        cover: item.cover,
        foodCost: foodCostFinal,
        bevCost: bevCostFinal,
        totalCost: totalCostFinal
      });
    });

    rows.sort((a, b) => a.outlet.localeCompare(b.outlet) || mealsList.indexOf(a.meal) - mealsList.indexOf(b.meal));

    dates.sort();
    const activeOutlets = [...new Set(rows.map(r => r.outlet))];
    const datasets = mealsList.map(mealName => {
      const dataPoints = activeOutlets.map(outletName => {
        const found = rows.find(r => r.outlet === outletName && r.meal === mealName);
        return found ? found.totalRev : 0;
      });
      return {
        label: mealName,
        data: dataPoints,
        backgroundColor: mealColors[mealName] || '#b8860b',
        borderRadius: 3
      };
    });

    res.json({
      success: true,
      data: {
        rows,
        chartData: {
          labels: activeOutlets,
          datasets: datasets
        },
        dateRange: {
          min: dates.length ? dates[0] : '',
          max: dates.length ? dates[dates.length - 1] : ''
        }
      }
    });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`✅ Server is running at http://localhost:${PORT}`);
});