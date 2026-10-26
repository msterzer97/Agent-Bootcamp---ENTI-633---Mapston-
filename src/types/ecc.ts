/**
 * Types for LandDev ECC Budget Variance Tracker and Reconciliation
 */

export interface BudgetDetailItem {
  rowNumber: number;
  phase: string;
  costCategoryCode: string; // Preserves leading zeros (e.g. '0100')
  costCategoryName: string; // Displayed cost type/category name
  budgetItemCode: string;   // Preserves leading zeros (e.g. '0010')
  budgetItemDescription: string;
  originalBudgetAmount: number;
  originalBudgetFormula?: string;
  updatedBudgetAmount?: number;
  previousBudgetStored?: number;
  variance?: number;
  status: 'matched' | 'unmatched_ecc' | 'unmatched_source' | 'duplicate_key' | 'formula_protected';
  exceptionMessage?: string;
}

export interface SourceBudgetItem {
  rowNumber: number;
  phase: string;
  costCategoryCode: string;
  costCategoryName: string;
  budgetItemCode: string;
  budgetItemDescription: string;
  budgetAmount: number;
}

export interface CompositeKey {
  phase: string;
  costCategoryCode: string;
  budgetItemCode: string;
}

export interface MatchingException {
  id: string;
  type: 'duplicate_ecc' | 'duplicate_source' | 'ecc_only' | 'source_only' | 'formula_cell' | 'ambiguous_column';
  sheetName: string;
  rowNumber: number;
  compositeKey: string;
  description: string;
  suggestedAction: string;
  affectsAmounts: boolean;
}

export interface PhaseCostTypeVariance {
  phase: string;
  costType: string;
  previousBudget: number;
  updatedBudget: number;
  change: number; // updatedBudget - previousBudget
  exceedsStrictThreshold: boolean; // Math.abs(change) > 25000.00
}

export interface ReconciliationReport {
  isSuccess: boolean;
  isProvisional: boolean;
  totalRecordsMatched: number;
  totalECCRecords: number;
  totalSourceRecords: number;
  
  // Financial totals for CV Budget Items
  totalPreviousBudget: number;
  totalUpdatedBudget: number;
  netVariance: number;
  
  // Checks
  columnEIntegrityVerified: boolean;
  columnHFormulasVerified: boolean;
  sourceTotalsEqualUpdatedTotals: boolean;
  noDoubleCountingVerified: boolean;
  formulaProtectionVerified: boolean;
  
  // Aggregated variances
  phaseCostTypeVariances: PhaseCostTypeVariance[];
  thresholdVariances: PhaseCostTypeVariance[]; // strictly > 25000 or < -25000
  
  // Exceptions
  exceptions: MatchingException[];
  
  // Saved file info
  updatedFileName?: string;
  updatedFileBlob?: Blob;
  
  // Leadership email
  emailDraft: string;
  aiAssisted: boolean;
  auditNotes: string[];

  // ECC Project Summary Sheet data (Columns E, F, H)
  summaryRows: ECCSummaryRow[];
}

export interface ECCSummaryRow {
  rowNumber: number;
  phaseCode: string;
  phaseDesc: string;
  originalBudget: number;
  approvedChanges: number;
  previousMonthColE: number | null;
  currentMonthColF: number;
  forecastColG: number;
  varianceColH: number | null;
  colHFormula?: string;
  isTotal: boolean;
}

export interface InspectionResult {
  hasCVBudgetItemsSheet: boolean;
  hasSummarySheet: boolean;
  sheetNames: string[];
  identifiedColumns: {
    phaseCol?: number;
    costCategoryCodeCol?: number;
    costCategoryNameCol?: number;
    budgetItemCodeCol?: number;
    budgetItemDescCol?: number;
    currentBudgetCol?: number;
    previousBudgetCol?: number;
    changeCol?: number;
  };
  totalDataRows: number;
  formulaCellsCount: number;
  warnings: string[];
}
