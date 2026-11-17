export type ClinicalColumnCategory =
  | 'patient_id'
  | 'prescription'
  | 'clinical_finding'
  | 'medication_treatment'
  | 'surgery'
  | 'normal';

export type CellAlignment = 'left' | 'center' | 'right';

export interface ColumnVisualMeta {
  category: ClinicalColumnCategory;
  align: CellAlignment;
  defaultWidth: number;
}

const MEDICATION_DRUG_NAMES = new Set([
  'isotear',
  'oculose',
  'croire',
  'croire-dx',
  'excipro',
  'ciprocir',
  'timolol',
  'lat+tim',
  'lat',
  'exclofen'
]);

const COMPACT_OPTICAL_HEADERS = new Set([
  'eye',
  'va-rx',
  'va+orx',
  'sph',
  'cyl',
  'axis',
  'va+nrx',
  'add',
  'pd',
  'unaided od',
  'unaided os',
  'unaided ou',
  'bcv od',
  'bcv os',
  'bcv ou',
  'hx bcv od',
  'hx bcv os',
  'sup',
  'nas',
  'inf',
  'temp'
]);

const COMPACT_CLINICAL_METRIC_HEADERS = new Set([
  'cdr',
  'iop',
  'md',
  'f. loss',
  'f. losses',
  'f. positive',
  'f. negative',
  'f. negatives',
  'pachym',
  'bp check',
  'signal strength',
  'questions',
  'p. seen',
  'p. missed'
]);

export function getColumnVisualMeta(
  workspace: 'reception' | 'doctor',
  sheetName: string,
  rawHeader: string,
  colIndex: number
): ColumnVisualMeta {
  const h = (rawHeader ?? '').trim().toLowerCase();
  const sheetLower = (sheetName ?? '').trim().toLowerCase();

  // Reception workspace: standardized widths & alignments with neutral clinical styling
  if (workspace === 'reception') {
    if (h === 'name' || h === "patient's name") {
      return { category: 'normal', align: 'left', defaultWidth: 195 };
    }
    if (h === 'recorded_by' || h === 'examined_by' || h === 'dic') {
      return { category: 'normal', align: 'left', defaultWidth: 135 };
    }
    if (h === 'phone_no' || h === 'contact') {
      return { category: 'normal', align: 'center', defaultWidth: 125 };
    }
    if (h === 'date' || h === 'date_of_birth' || h === 'review date') {
      return { category: 'normal', align: 'center', defaultWidth: 110 };
    }
    if (h === 'test_type' || h === 'medication' || h === 'review test' || h === 'medicine') {
      return { category: 'normal', align: 'left', defaultWidth: 155 };
    }
    if (
      h === 'test_fee' ||
      h === 'total' ||
      h === 'glass total' ||
      h === 'paid' ||
      h === 'glasses bal' ||
      h === 'new' ||
      h === 'renewal'
    ) {
      return { category: 'normal', align: 'right', defaultWidth: 100 };
    }
    if (h === 'drops_no') {
      return { category: 'normal', align: 'center', defaultWidth: 90 };
    }
    return { category: 'normal', align: 'left', defaultWidth: 110 };
  }

  // Doctor workspace special sheet structures with blank headers
  if (sheetLower === 'dm, hpt') {
    if (colIndex === 1) {
      return { category: 'medication_treatment', align: 'left', defaultWidth: 340 };
    }
    return { category: 'normal', align: 'left', defaultWidth: 110 };
  }

  if (sheetLower === 'sheet1') {
    if (colIndex === 0 || h === 'services') {
      return { category: 'normal', align: 'left', defaultWidth: 280 };
    }
    if (colIndex === 1 || h.includes('price')) {
      return { category: 'normal', align: 'right', defaultWidth: 120 };
    }
  }

  // Lens Rx col 0 has whitespace header "  " containing consultation Dates
  if (sheetLower === 'lens rx' && colIndex === 0 && h === '') {
    return { category: 'normal', align: 'center', defaultWidth: 105 };
  }

  // 1. Patient Identification fields
  if (
    h === 'name' ||
    h === "patient's name" ||
    h === 'surname' ||
    h === 'other name' ||
    h === 'other name(s)'
  ) {
    return { category: 'patient_id', align: 'left', defaultWidth: 185 };
  }
  if (
    h === 'patient no.' ||
    h === 'patient no' ||
    h === 'phone no.' ||
    h === 'phone no' ||
    h === 'contact' ||
    h === 'tin' ||
    h === 'tin/gh card no.' ||
    h === 'dvla reference number' ||
    h === 'dvla reference #' ||
    h === 'dvla invoice number' ||
    h === 'dvla invoice #' ||
    h === 'licence number'
  ) {
    return { category: 'patient_id', align: 'center', defaultWidth: 140 };
  }
  if (h === 'date of birth') {
    return { category: 'patient_id', align: 'center', defaultWidth: 115 };
  }
  if (
    h === 'residence' ||
    h === 'address' ||
    h === 'postal address' ||
    h === 'email' ||
    h === 'email address' ||
    h === 'know your client' ||
    h === 'driving school'
  ) {
    return { category: 'patient_id', align: 'left', defaultWidth: 165 };
  }

  // 2. Eye Prescription & Lens Rx fields
  if (COMPACT_OPTICAL_HEADERS.has(h)) {
    return {
      category: 'prescription',
      align: 'center',
      defaultWidth: h === 'va-rx' || h.startsWith('unaided') || h.startsWith('bcv') || h.startsWith('hx bcv') ? 105 : 88
    };
  }
  if (h === 'lensmetry/auto ref' || h === 'refraction result') {
    return { category: 'prescription', align: 'left', defaultWidth: 175 };
  }
  if (h === 'lens type' || h === 'lens treatment') {
    return { category: 'prescription', align: 'left', defaultWidth: 135 };
  }

  // 3. Surgery-related fields
  if (h.includes('post-op')) {
    return { category: 'surgery', align: 'left', defaultWidth: 225 };
  }
  if (h === 'surgery type') {
    return { category: 'surgery', align: 'left', defaultWidth: 165 };
  }
  if (h === 'booking date' || h === 'surgery date') {
    return { category: 'surgery', align: 'center', defaultWidth: 115 };
  }
  if (h === 'payment mode' || h === 'venue') {
    return { category: 'surgery', align: 'left', defaultWidth: 130 };
  }
  if (sheetLower === 'surgery' && h === 'comments') {
    return { category: 'surgery', align: 'left', defaultWidth: 200 };
  }

  // 4. Medication & Treatment fields
  if (h === 'eye drop' || h === 'remarks + medications' || h === 'medication') {
    return { category: 'medication_treatment', align: 'left', defaultWidth: 225 };
  }
  if (h === 'function' || h === 'dosage') {
    return { category: 'medication_treatment', align: 'left', defaultWidth: 205 };
  }
  if (h === 'trade name') {
    return { category: 'medication_treatment', align: 'left', defaultWidth: 165 };
  }
  if (h === 'type' || h === 'availability') {
    return { category: 'medication_treatment', align: 'center', defaultWidth: 98 };
  }
  if (MEDICATION_DRUG_NAMES.has(h)) {
    return { category: 'medication_treatment', align: 'center', defaultWidth: 92 };
  }

  // 5. Clinical Findings & Important Exam Results
  if (
    h === 'chief complaint' ||
    h === 'diagnosis' ||
    h === 'glaucoma hemifield test (ght)' ||
    h === 'pathological remarks'
  ) {
    return { category: 'clinical_finding', align: 'left', defaultWidth: 225 };
  }
  if (
    h === 'medical hx' ||
    h === 'ocular hx' ||
    h === 'drug hx' ||
    h === 'results conclusion' ||
    h === 'result conclusion' ||
    h === 'final results'
  ) {
    return { category: 'clinical_finding', align: 'left', defaultWidth: 165 };
  }
  if (
    h === 'eyelid' ||
    h === 'conj/sclera' ||
    h === 'cornea' ||
    h === 'pupil' ||
    h === 'lens' ||
    h === 'vitreous' ||
    h === 'fundus'
  ) {
    return { category: 'clinical_finding', align: 'left', defaultWidth: 145 };
  }
  if (COMPACT_CLINICAL_METRIC_HEADERS.has(h)) {
    return { category: 'clinical_finding', align: 'center', defaultWidth: 92 };
  }

  // 6. Standard dates, targets & auxiliary fields
  if (h === 'date' || h === 'oct-date' || h === 'time' || h === 'mar') {
    return { category: 'normal', align: 'center', defaultWidth: 105 };
  }
  if (h.includes('target')) {
    return { category: 'normal', align: 'center', defaultWidth: 115 };
  }
  if (h === 'doctor' || h === 'review' || h === 'remarks' || h === 'service type' || h === 'type/place') {
    return { category: 'normal', align: 'left', defaultWidth: 135 };
  }
  if (h === 'sms') {
    return { category: 'normal', align: 'center', defaultWidth: 85 };
  }

  return { category: 'normal', align: 'left', defaultWidth: 105 };
}
