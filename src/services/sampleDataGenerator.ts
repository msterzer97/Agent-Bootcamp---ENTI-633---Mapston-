/**
 * Realistic Sample Workbook Generator for Acceptance Tests 1-5
 * Creates binary XLSX workbooks directly in-memory using ExcelJS with authentic styling,
 * preserved formulas, leading-zero codes, and exact test conditions.
 */

import ExcelJS from 'exceljs';

export interface SampleWorkbookPair {
  id: string;
  name: string;
  testCaseLabel: string;
  description: string;
  eccFileName: string;
  sourceFileName: string;
  createECC: () => Promise<ArrayBuffer>;
  createSource: () => Promise<ArrayBuffer>;
}

// Styling helper for realistic Excel appearance
function applyHeaderStyle(row: ExcelJS.Row, bgHex = '1E293B', fgHex = 'FFFFFF') {
  row.eachCell((cell) => {
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: fgHex } };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: bgHex },
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = {
      top: { style: 'thin', color: { argb: '334155' } },
      bottom: { style: 'medium', color: { argb: '0F172A' } },
      left: { style: 'thin', color: { argb: '334155' } },
      right: { style: 'thin', color: { argb: '334155' } },
    };
  });
  row.height = 24;
}

function applyDataRowStyle(row: ExcelJS.Row, isEven: boolean) {
  row.eachCell((cell, colNumber) => {
    cell.font = { name: 'Calibri', size: 10 };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: isEven ? 'F8FAFC' : 'FFFFFF' },
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'E2E8F0' } },
      bottom: { style: 'thin', color: { argb: 'E2E8F0' } },
      left: { style: 'thin', color: { argb: 'E2E8F0' } },
      right: { style: 'thin', color: { argb: 'E2E8F0' } },
    };
  });
  row.height = 20;
}

// Generate Baseline ECC Workbook with "CV Budget Items" and "Project Summary"
export async function createECCWorkbook(options: {
  scenario: 'clean' | 'duplicates' | 'strict_threshold' | 'formula_protection';
}): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'LandDev Finance Team';
  wb.lastModifiedBy = 'LandDev ECC Controller';
  wb.created = new Date();
  wb.modified = new Date();

  // 1. Project Summary Sheet
  const sumSheet = wb.addWorksheet('Project Summary', {
    views: [{ showGridLines: true }],
  });

  // Summary headers: Row 1-2
  sumSheet.mergeCells('A1:H1');
  const titleCell = sumSheet.getCell('A1');
  titleCell.value = 'LAND DEVELOPMENT PROJECT SUMMARY - ECC VARIANCE TRACKER';
  titleCell.font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FFFFFF' } };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '0F172A' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
  sumSheet.getRow(1).height = 32;

  // Header row 2
  const sumHeaders = [
    'Phase Code',          // A
    'Phase Description',   // B
    'Original Budget',     // C
    'Approved Changes',    // D
    'Previous Month',      // E (Reserved manual-input column!)
    'Current Month',       // F
    'Forecast at Complete',// G
    '∆ to Previous Budget' // H (Conditional formula =IF(E[row]<>'',F[row]-E[row],''))
  ];
  sumSheet.getRow(2).values = sumHeaders;
  applyHeaderStyle(sumSheet.getRow(2), '334155');

  // Summary data rows:
  // Row 3: Phase 1
  sumSheet.getRow(3).values = [
    'PH-01',
    'Phase 1 - Single Family Lots',
    2450000,
    0,
    2450000, // E3: Previous Month baseline
    2450000, // F3: Current Month
    2450000, // G3
    { formula: 'IF(E3<>"",F3-E3,"")', result: 0 } // H3
  ];
  // Row 4: Phase 2
  sumSheet.getRow(4).values = [
    'PH-02',
    'Phase 2 - Commercial Town Center',
    1820000,
    0,
    1820000, // E4: Previous Month baseline
    1820000, // F4
    1820000, // G4
    { formula: 'IF(E4<>"",F4-E4,"")', result: 0 } // H4
  ];
  // Row 5: Common Elements (Column E intentionally left blank to test blank Column E -> blank Column H)
  sumSheet.getRow(5).values = [
    'COMMON',
    'Common Infrastructure & Offsites',
    950000,
    0,
    null,    // E5: BLANK manual input to verify Column H stays blank!
    950000,  // F5
    950000,  // G5
    { formula: 'IF(E5<>"",F5-E5,"")', result: '' } // H5
  ];
  // Row 6: Total
  sumSheet.getRow(6).values = [
    'TOTAL',
    'Project Grand Total',
    { formula: 'SUM(C3:C5)', result: 5220000 },
    0,
    { formula: 'SUM(E3:E4)', result: 4270000 }, // E6 Total
    { formula: 'SUM(F3:F5)', result: 5220000 }, // F6 Total
    { formula: 'SUM(G3:G5)', result: 5220000 },
    { formula: 'IF(ISNUMBER(E6),F6-E6,"")', result: 950000 } // H6
  ];

  for (let r = 3; r <= 6; r++) {
    const row = sumSheet.getRow(r);
    applyDataRowStyle(row, r % 2 === 0);
    // Format money
    row.getCell(3).numFmt = '"$"#,##0.00';
    row.getCell(4).numFmt = '"$"#,##0.00';
    row.getCell(5).numFmt = '"$"#,##0.00;("$"#,##0.00);"-"';
    row.getCell(6).numFmt = '"$"#,##0.00;("$"#,##0.00);"-"';
    row.getCell(7).numFmt = '"$"#,##0.00;("$"#,##0.00);"-"';
    row.getCell(8).numFmt = '"+"$"#,##0.00;"-"$"#,##0.00;""';
  }
  sumSheet.getRow(6).font = { bold: true };

  sumSheet.columns = [
    { width: 14 },
    { width: 34 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 18 },
    { width: 22 },
    { width: 22 },
  ];

  // 2. CV Budget Items Sheet (The baseline sheet)
  const cvSheet = wb.addWorksheet('CV Budget Items', {
    views: [{ showGridLines: true }],
  });

  const cvHeaders = [
    'Phase',                  // 1
    'Cost Category Code',     // 2 (Preserves leading zeros!)
    'Cost Category Name',     // 3
    'Budget Item Code',       // 4 (Preserves leading zeros!)
    'Budget Item Description',// 5
    'Current Budget',         // 6
    'Previous Budget',        // 7
    'Variance'                // 8
  ];
  cvSheet.getRow(1).values = cvHeaders;
  applyHeaderStyle(cvSheet.getRow(1), '1E3A8A');

  interface RowSpec {
    phase: string;
    catCode: string;
    catName: string;
    itemCode: string;
    itemDesc: string;
    amount: number | { formula: string; result?: number };
    prevAmount?: number;
  }

  let items: RowSpec[] = [];

  if (options.scenario === 'clean') {
    items = [
      // Phase 1
      { phase: 'Phase 1', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0010', itemDesc: 'Mass Rough Grading Cut/Fill', amount: 480000, prevAmount: 480000 },
      { phase: 'Phase 1', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0020', itemDesc: 'Site Stripping & Stockpiling', amount: 120000, prevAmount: 120000 },
      { phase: 'Phase 1', catCode: '0200', catName: 'Wet Utilities', itemCode: '0010', itemDesc: 'Sanitary Sewer Main & Laterals', amount: 350000, prevAmount: 350000 },
      { phase: 'Phase 1', catCode: '0200', catName: 'Wet Utilities', itemCode: '0020', itemDesc: 'Domestic Water Mains & Hydrants', amount: 280000, prevAmount: 280000 },
      { phase: 'Phase 1', catCode: '0300', catName: 'Dry Utilities', itemCode: '0010', itemDesc: 'Joint Trench Electric & Telecom', amount: 210000, prevAmount: 210000 },
      { phase: 'Phase 1', catCode: '0400', catName: 'Paving & Concrete', itemCode: '0010', itemDesc: 'Asphalt Concrete Base Course', amount: 320000, prevAmount: 320000 },
      // Phase 2
      { phase: 'Phase 2', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0010', itemDesc: 'Town Center Structural Pad Grading', amount: 410000, prevAmount: 410000 },
      { phase: 'Phase 2', catCode: '0200', catName: 'Wet Utilities', itemCode: '0010', itemDesc: 'Storm Drainage Retention Vault', amount: 540000, prevAmount: 540000 },
      { phase: 'Phase 2', catCode: '0400', catName: 'Paving & Concrete', itemCode: '0010', itemDesc: 'Pedestrian Hardscape & Curbing', amount: 290000, prevAmount: 290000 },
      // Common
      { phase: 'Common', catCode: '0500', catName: 'Engineering & Survey', itemCode: '0010', itemDesc: 'Civil Engineering Design & Staking', amount: 180000, prevAmount: 180000 },
      { phase: 'Common', catCode: '0600', catName: 'Environmental & Permits', itemCode: '0010', itemDesc: 'Wetlands Mitigation & SWPPP Monitoring', amount: 150000, prevAmount: 150000 },
    ];
  } else if (options.scenario === 'duplicates') {
    items = [
      // Normal items
      { phase: 'Phase 1', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0010', itemDesc: 'Mass Rough Grading', amount: 480000, prevAmount: 480000 },
      // DUPLICATE KEY in ECC: row with same composite key (Phase 1, 0100, 0010)
      { phase: 'Phase 1', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0010', itemDesc: 'Mass Rough Grading - Duplicate Entry', amount: 50000, prevAmount: 50000 },
      // Item that will be MISSING from monthly source (ECC-only)
      { phase: 'Phase 1', catCode: '0200', catName: 'Wet Utilities', itemCode: '0088', itemDesc: 'Special Groundwater Extraction Well (Held)', amount: 95000, prevAmount: 95000 },
      // Normal item
      { phase: 'Phase 1', catCode: '0300', catName: 'Dry Utilities', itemCode: '0010', itemDesc: 'Joint Trench Mainline', amount: 210000, prevAmount: 210000 },
      { phase: 'Phase 2', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0010', itemDesc: 'Town Center Pad Fill', amount: 390000, prevAmount: 390000 },
      { phase: 'Common', catCode: '0500', catName: 'Engineering & Survey', itemCode: '0010', itemDesc: 'Survey Staking & Topo', amount: 120000, prevAmount: 120000 },
    ];
  } else if (options.scenario === 'strict_threshold') {
    // Specifically crafted to test exact threshold boundary:
    // Combo 1: +$25,000.00 EXACT (Must be EXCLUDED)
    // Combo 2: -$25,000.00 EXACT (Must be EXCLUDED)
    // Combo 3: +$25,000.01 STRICT (Must be INCLUDED)
    // Combo 4: -$25,000.01 STRICT (Must be INCLUDED)
    // Combo 5: +$12,000.00 (Must be EXCLUDED)
    items = [
      // Phase 1 - 0100 Earthwork (Will increase by EXACTLY $25,000.00 from 200,000 to 225,000)
      { phase: 'Phase 1', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0010', itemDesc: 'Rough Grading Baseline', amount: 200000, prevAmount: 200000 },
      // Phase 1 - 0200 Wet Utilities (Will decrease by EXACTLY $25,000.00 from 300,000 to 275,000)
      { phase: 'Phase 1', catCode: '0200', catName: 'Wet Utilities', itemCode: '0010', itemDesc: 'Sewer Trunk Line', amount: 300000, prevAmount: 300000 },
      // Phase 2 - 0300 Paving (Will increase by EXACTLY $25,000.01 from 400,000 to 425,000.01)
      { phase: 'Phase 2', catCode: '0300', catName: 'Paving & Concrete', itemCode: '0010', itemDesc: 'Arterial Road Paving', amount: 400000, prevAmount: 400000 },
      // Phase 2 - 0400 Landscape (Will decrease by EXACTLY $25,000.01 from 150,000 to 124,999.99)
      { phase: 'Phase 2', catCode: '0400', catName: 'Landscape & Irrigation', itemCode: '0010', itemDesc: 'Entry Monument Landscape', amount: 150000, prevAmount: 150000 },
      // Common - 0500 Engineering (Will increase by $12,500.00, below threshold)
      { phase: 'Common', catCode: '0500', catName: 'Engineering & Survey', itemCode: '0010', itemDesc: 'Civil Geotech Inspection', amount: 80000, prevAmount: 80000 },
    ];
  } else if (options.scenario === 'formula_protection') {
    items = [
      { phase: 'Phase 1', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0010', itemDesc: 'Rough Grading Mass Haul', amount: 350000, prevAmount: 350000 },
      // ROW 3: CURRENT BUDGET HAS A FORMULA INSTEAD OF TYPED VALUE!
      // Test 5 requirement: App detects formula cell, does NOT overwrite, stops and flags exception!
      { phase: 'Phase 1', catCode: '0200', catName: 'Wet Utilities', itemCode: '0020', itemDesc: 'Calculated Sanitary Sewer (Contingency Formula)', amount: { formula: 'F2*0.65', result: 227500 }, prevAmount: 227500 },
      { phase: 'Phase 2', catCode: '0400', catName: 'Paving & Concrete', itemCode: '0010', itemDesc: 'Curb and Gutter System', amount: 180000, prevAmount: 180000 },
      { phase: 'Common', catCode: '0500', catName: 'Engineering & Survey', itemCode: '0010', itemDesc: 'Topographic Boundary Mapping', amount: 95000, prevAmount: 95000 },
    ];
  }

  // Populate CV Budget Items
  items.forEach((item, idx) => {
    const rowNum = idx + 2;
    const row = cvSheet.getRow(rowNum);

    // Format codes as text to strictly preserve leading zeros
    const catCodeCell = row.getCell(2);
    catCodeCell.value = item.catCode;
    catCodeCell.numFmt = '@';

    const itemCodeCell = row.getCell(4);
    itemCodeCell.value = item.itemCode;
    itemCodeCell.numFmt = '@';

    row.getCell(1).value = item.phase;
    row.getCell(3).value = item.catName;
    row.getCell(5).value = item.itemDesc;

    const currCell = row.getCell(6);
    if (typeof item.amount === 'object' && 'formula' in item.amount) {
      currCell.value = { formula: item.amount.formula, result: item.amount.result };
    } else {
      currCell.value = item.amount;
    }
    currCell.numFmt = '"$"#,##0.00;("$"#,##0.00);"-"';

    const prevCell = row.getCell(7);
    prevCell.value = item.prevAmount ?? item.amount;
    prevCell.numFmt = '"$"#,##0.00;("$"#,##0.00);"-"';

    const varCell = row.getCell(8);
    varCell.value = { formula: `F${rowNum}-G${rowNum}`, result: 0 };
    varCell.numFmt = '"+"$"#,##0.00;"-"$"#,##0.00;""';

    applyDataRowStyle(row, idx % 2 === 0);
  });

  // Set column widths
  cvSheet.columns = [
    { width: 14 },
    { width: 20 },
    { width: 28 },
    { width: 18 },
    { width: 42 },
    { width: 20 },
    { width: 20 },
    { width: 20 },
  ];

  const buffer = await wb.xlsx.writeBuffer();
  return buffer;
}

// Generate Updated Monthly LandDev Source Workbook
export async function createSourceWorkbook(options: {
  scenario: 'clean' | 'duplicates' | 'strict_threshold' | 'formula_protection';
}): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'LandDev Construction Field Office';
  wb.lastModifiedBy = 'Civil Field Engineer';
  wb.created = new Date();
  wb.modified = new Date();

  const sheet = wb.addWorksheet('LandDev Monthly Budget', {
    views: [{ showGridLines: true }],
  });

  const headers = [
    'Phase',
    'Cost Category Code',
    'Cost Category Name',
    'Budget Item Code',
    'Budget Item Description',
    'Updated Monthly Budget'
  ];
  sheet.getRow(1).values = headers;
  applyHeaderStyle(sheet.getRow(1), '047857'); // Emerald green for updated source

  interface SourceItemSpec {
    phase: string;
    catCode: string;
    catName: string;
    itemCode: string;
    itemDesc: string;
    amount: number;
  }

  let items: SourceItemSpec[] = [];

  if (options.scenario === 'clean') {
    // Phase 1: 0100 Grading changes by +$45,000 (Material! > $25k)
    // Phase 1: 0200 Utilities changes by -$15,000 (Immaterial < $25k)
    // Phase 1: 0300 Dry Utilities changes by +$3,000 (Immaterial)
    // Phase 1: 0400 Paving changes by +$35,000 (Material! > $25k)
    // Phase 2: 0100 Grading changes by -$32,000 (Material! < -$25k)
    // Phase 2: 0200 Utilities changes by +$50,000 (Material! > $25k)
    // Phase 2: 0400 Paving unchanged ($0)
    // Common: 0500 Engineering changes by +$8,000
    // Common: 0600 Environmental changes by -$4,000
    items = [
      { phase: 'Phase 1', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0010', itemDesc: 'Mass Rough Grading Cut/Fill', amount: 525000 }, // +45k
      { phase: 'Phase 1', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0020', itemDesc: 'Site Stripping & Stockpiling', amount: 120000 }, // 0
      { phase: 'Phase 1', catCode: '0200', catName: 'Wet Utilities', itemCode: '0010', itemDesc: 'Sanitary Sewer Main & Laterals', amount: 335000 }, // -15k
      { phase: 'Phase 1', catCode: '0200', catName: 'Wet Utilities', itemCode: '0020', itemDesc: 'Domestic Water Mains & Hydrants', amount: 280000 }, // 0
      { phase: 'Phase 1', catCode: '0300', catName: 'Dry Utilities', itemCode: '0010', itemDesc: 'Joint Trench Electric & Telecom', amount: 213000 }, // +3k
      { phase: 'Phase 1', catCode: '0400', catName: 'Paving & Concrete', itemCode: '0010', itemDesc: 'Asphalt Concrete Base Course', amount: 355000 }, // +35k
      { phase: 'Phase 2', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0010', itemDesc: 'Town Center Structural Pad Grading', amount: 378000 }, // -32k
      { phase: 'Phase 2', catCode: '0200', catName: 'Wet Utilities', itemCode: '0010', itemDesc: 'Storm Drainage Retention Vault', amount: 590000 }, // +50k
      { phase: 'Phase 2', catCode: '0400', catName: 'Paving & Concrete', itemCode: '0010', itemDesc: 'Pedestrian Hardscape & Curbing', amount: 290000 }, // 0
      { phase: 'Common', catCode: '0500', catName: 'Engineering & Survey', itemCode: '0010', itemDesc: 'Civil Engineering Design & Staking', amount: 188000 }, // +8k
      { phase: 'Common', catCode: '0600', catName: 'Environmental & Permits', itemCode: '0010', itemDesc: 'Wetlands Mitigation & SWPPP Monitoring', amount: 146000 }, // -4k
    ];
  } else if (options.scenario === 'duplicates') {
    items = [
      // Matched item
      { phase: 'Phase 1', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0010', itemDesc: 'Mass Rough Grading', amount: 480000 },
      // Note: Phase 1 0200 0088 is OMITTED from source to create an ECC-ONLY missing record!
      { phase: 'Phase 1', catCode: '0300', catName: 'Dry Utilities', itemCode: '0010', itemDesc: 'Joint Trench Mainline', amount: 210000 },
      // DUPLICATE in SOURCE:
      { phase: 'Phase 2', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0010', itemDesc: 'Town Center Pad Fill (Contractor Submittal)', amount: 400000 },
      { phase: 'Phase 2', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0010', itemDesc: 'Town Center Pad Fill (Engineer Revision)', amount: 405000 },
      // NEW SOURCE RECORD (Source-only item not in ECC):
      { phase: 'Phase 2', catCode: '0700', catName: 'Retaining Walls', itemCode: '0055', itemDesc: 'Modular Block Gravity Wall 4B', amount: 62000 },
      // Normal item
      { phase: 'Common', catCode: '0500', catName: 'Engineering & Survey', itemCode: '0010', itemDesc: 'Survey Staking & Topo', amount: 125000 },
    ];
  } else if (options.scenario === 'strict_threshold') {
    // Exactly matches the threshold test numbers:
    // Phase 1 - 0100: 200,000 -> 225,000 (Change = +$25,000.00 EXACT => MUST BE EXCLUDED)
    // Phase 1 - 0200: 300,000 -> 275,000 (Change = -$25,000.00 EXACT => MUST BE EXCLUDED)
    // Phase 2 - 0300: 400,000 -> 425,000.01 (Change = +$25,000.01 STRICT => MUST BE INCLUDED)
    // Phase 2 - 0400: 150,000 -> 124,999.99 (Change = -$25,000.01 STRICT => MUST BE INCLUDED)
    // Common - 0500: 80,000 -> 92,500 (Change = +$12,500.00 => MUST BE EXCLUDED)
    items = [
      { phase: 'Phase 1', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0010', itemDesc: 'Rough Grading Baseline', amount: 225000.00 },
      { phase: 'Phase 1', catCode: '0200', catName: 'Wet Utilities', itemCode: '0010', itemDesc: 'Sewer Trunk Line', amount: 275000.00 },
      { phase: 'Phase 2', catCode: '0300', catName: 'Paving & Concrete', itemCode: '0010', itemDesc: 'Arterial Road Paving', amount: 425000.01 },
      { phase: 'Phase 2', catCode: '0400', catName: 'Landscape & Irrigation', itemCode: '0010', itemDesc: 'Entry Monument Landscape', amount: 124999.99 },
      { phase: 'Common', catCode: '0500', catName: 'Engineering & Survey', itemCode: '0010', itemDesc: 'Civil Geotech Inspection', amount: 92500.00 },
    ];
  } else if (options.scenario === 'formula_protection') {
    items = [
      { phase: 'Phase 1', catCode: '0100', catName: 'Earthwork & Grading', itemCode: '0010', itemDesc: 'Rough Grading Mass Haul', amount: 375000 },
      // Attempting to update row 3 which has a formula in ECC:
      { phase: 'Phase 1', catCode: '0200', catName: 'Wet Utilities', itemCode: '0020', itemDesc: 'Calculated Sanitary Sewer (Contingency Formula)', amount: 240000 },
      { phase: 'Phase 2', catCode: '0400', catName: 'Paving & Concrete', itemCode: '0010', itemDesc: 'Curb and Gutter System', amount: 180000 },
      { phase: 'Common', catCode: '0500', catName: 'Engineering & Survey', itemCode: '0010', itemDesc: 'Topographic Boundary Mapping', amount: 105000 },
    ];
  }

  items.forEach((item, idx) => {
    const rowNum = idx + 2;
    const row = sheet.getRow(rowNum);

    row.getCell(1).value = item.phase;

    const catCell = row.getCell(2);
    catCell.value = item.catCode;
    catCell.numFmt = '@';

    row.getCell(3).value = item.catName;

    const itemCell = row.getCell(4);
    itemCell.value = item.itemCode;
    itemCell.numFmt = '@';

    row.getCell(5).value = item.itemDesc;

    const amtCell = row.getCell(6);
    amtCell.value = item.amount;
    amtCell.numFmt = '"$"#,##0.00;("$"#,##0.00);"-"';

    applyDataRowStyle(row, idx % 2 === 0);
  });

  sheet.columns = [
    { width: 14 },
    { width: 20 },
    { width: 28 },
    { width: 18 },
    { width: 44 },
    { width: 24 },
  ];

  const buffer = await wb.xlsx.writeBuffer();
  return buffer;
}

// Built-in Sample Pairs matching Acceptance Tests 1-5
export const ACCEPTANCE_TEST_PRESETS: SampleWorkbookPair[] = [
  {
    id: 'test-1',
    name: 'Acceptance Test 1: Complete Successful Update',
    testCaseLabel: 'Test 1: 1-to-1 Match & Clean Reconciliation',
    description: 'All records have unique 1-to-1 matches across phases. Reconciles perfectly, updates Current Budget, preserves Column E, and identifies material changes.',
    eccFileName: 'ECC_Variance_Tracker_October2026.xlsx',
    sourceFileName: 'LandDev_Monthly_Budget_Source_Nov2026.xlsx',
    createECC: () => createECCWorkbook({ scenario: 'clean' }),
    createSource: () => createSourceWorkbook({ scenario: 'clean' }),
  },
  {
    id: 'test-2',
    name: 'Acceptance Test 2: Missing Required Input',
    testCaseLabel: 'Test 2: Missing Workbook Detection',
    description: 'Supplies only one workbook. App halts, identifies the missing input, and produces no updated workbook, reconciliation, or draft.',
    eccFileName: 'ECC_Variance_Tracker_Only.xlsx',
    sourceFileName: '', // Intentionally omitted
    createECC: () => createECCWorkbook({ scenario: 'clean' }),
    createSource: async () => new ArrayBuffer(0),
  },
  {
    id: 'test-3',
    name: 'Acceptance Test 3: Duplicate, Missing & New Records',
    testCaseLabel: 'Test 3: Reconciliation Exceptions & Provisional Status',
    description: 'Contains 1 duplicate key in ECC, 1 duplicate in Source, 1 item missing from source (ECC-only), and 1 new source item. Labels affected totals as provisional.',
    eccFileName: 'ECC_Tracker_With_Exceptions.xlsx',
    sourceFileName: 'LandDev_Budget_With_Exceptions.xlsx',
    createECC: () => createECCWorkbook({ scenario: 'duplicates' }),
    createSource: () => createSourceWorkbook({ scenario: 'duplicates' }),
  },
  {
    id: 'test-4',
    name: 'Acceptance Test 4: Strict $25,000 Threshold Filter',
    testCaseLabel: 'Test 4: Strict Threshold (+/- $25,000.00 Excluded, +/- $25,000.01 Included)',
    description: 'Evaluates boundary values: +$25,000.00 and -$25,000.00 are strictly excluded. +$25,000.01 and -$25,000.01 are strictly included.',
    eccFileName: 'ECC_Tracker_Strict_Threshold.xlsx',
    sourceFileName: 'LandDev_Budget_Strict_Threshold.xlsx',
    createECC: () => createECCWorkbook({ scenario: 'strict_threshold' }),
    createSource: () => createSourceWorkbook({ scenario: 'strict_threshold' }),
  },
  {
    id: 'test-5',
    name: 'Acceptance Test 5: Formula and Workbook Protection',
    testCaseLabel: 'Test 5: Formula Cell Protection',
    description: 'ECC Current Budget contains a formula cell. App detects formula, blocks overwrite, requests clarification, and preserves original workbook.',
    eccFileName: 'ECC_Tracker_Formula_Protected.xlsx',
    sourceFileName: 'LandDev_Budget_Source.xlsx',
    createECC: () => createECCWorkbook({ scenario: 'formula_protection' }),
    createSource: () => createSourceWorkbook({ scenario: 'formula_protection' }),
  },
];
