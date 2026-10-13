export type UserRole = 'reception' | 'doctor' | 'admin';

export interface User {
  id: string;
  username: string;
  name: string;
  role: UserRole;
  title: string;
}

export interface PatientRecord {
  id: string; // Stable internal identifier, e.g., MK-PAT-0001
  date: string;
  recordedBy: string;
  phoneNo: string;
  examinedBy: string;
  name: string;
  dateOfBirth: string;
  testType: string;
  testFee: number | string;
  dropsNo: number | string;
  medication: number | string;
  total: number | string;
  glassTotal: number | string;
  paid: number | string;
  glassesBal: number | string;
  reviewDate: string;
  reviewTest: string;
  customFields?: Record<string, string | number>;
  createdAt?: string;
  updatedAt?: string;
  lastUpdatedBy?: string;
}

export interface DVLARecord {
  id: string; // e.g. MK-DVLA-0001
  date: string;
  recordedBy: string; // The 2nd column in Excel (labeled with space)
  contact: string;
  dic: string; // Doctor in Charge / Examined By
  patientName: string;
  newFee: number | string;
  renewalFee: number | string;
  notes?: string; // e.g. "RW" found in Excel column
  customFields?: Record<string, string | number>;
  createdAt?: string;
  updatedAt?: string;
  lastUpdatedBy?: string;
}

export interface MedsRecord {
  id: string;
  medicineName: string;
  val1: number | string;
  val2: number | string;
  val3: number | string;
  val4: number | string;
}

export interface Sheet2Record {
  id: string;
  col1: string;
  col2: string;
  col3: string;
  contact: string;
  col5: string;
  date: string;
}

export interface AuditLogEntry {
  id: string;
  timestamp: string;
  actor: string;
  actorRole: UserRole;
  action: 'CREATE' | 'UPDATE' | 'DELETE' | 'SCHEMA_ADD_COLUMN' | 'SHEETS_SYNC';
  worksheet: 'PATIENT\'S DATA' | 'DVLA' | 'MEDS' | 'Sheet2' | 'SETTINGS';
  recordId?: string;
  details: string;
}

export interface ClinicSettings {
  clinicName: string;
  subtitle: string;
  currency: string;
  googleSheetsUrl: string; // Google Apps Script web app URL or Sheets URL
  googleSpreadsheetId: string;
  autoSync: boolean;
  syncIntervalMinutes: number;
  lastSyncedAt?: string;
  syncStatus: 'unconfigured' | 'synced' | 'pending' | 'error';
  syncErrorMessage?: string;
}

export interface ColumnDefinition {
  key: string;
  label: string;
  type: 'text' | 'number' | 'date' | 'select';
  required?: boolean;
  isCustom?: boolean;
}
