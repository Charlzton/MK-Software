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

/**
 * Deliberate content-proportional column sizing & clinical categorization:
 * - Reception maintains its existing orderly proportions.
 * - Doctor uses 5 balanced width tiers (76 / 96 / 116 / 156 / 216 px) so Lens Rx
 *   and sibling clinical sheets eliminate erratic undersized/oversized jumps.
 */
export function getColumnVisualMeta(
  workspace: 'reception' | 'doctor',
  sheetName: string,
  rawHeader: string,
  colIndex: number
): ColumnVisualMeta {
  const h = (rawHeader ?? '').trim().toLowerCase();
  const sheetLower = (sheetName ?? '').trim().toLowerCase();

  // Reception workspace: preserve orderly spreadsheet proportions
  if (workspace === 'reception') {
    if (h === 'name' || h === "patient's name") {
      return { category: 'normal', align: 'left', defaultWidth: 196 };
    }
    if (h === 'recorded_by' || h === 'examined_by' || h === 'dic') {
      return { category: 'normal', align: 'left', defaultWidth: 136 };
    }
    if (h === 'phone_no' || h === 'contact') {
      return { category: 'normal', align: 'center', defaultWidth: 124 };
    }
    if (h === 'date' || h === 'date_of_birth' || h === 'review date') {
      return { category: 'normal', align: 'center', defaultWidth: 112 };
    }
    if (h === 'test_type' || h === 'medication' || h === 'review test' || h === 'medicine') {
      return { category: 'normal', align: 'left', defaultWidth: 156 };
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
      return { category: 'normal', align: 'right', defaultWidth: 102 };
    }
    if (h === 'drops_no') {
      return { category: 'normal', align: 'center', defaultWidth: 92 };
    }
    return { category: 'normal', align: 'left', defaultWidth: 112 };
  }

  // Doctor workspace special sheet layouts
  if (sheetLower === 'dm, hpt') {
    if (colIndex === 1) {
      return { category: 'medication_treatment', align: 'left', defaultWidth: 248 };
    }
    return { category: 'normal', align: 'left', defaultWidth: 96 };
  }

  if (sheetLower === 'sheet1') {
    if (colIndex === 0 || h === 'services') {
      return { category: 'normal', align: 'left', defaultWidth: 236 };
    }
    if (colIndex === 1 || h.includes('price')) {
      return { category: 'normal', align: 'right', defaultWidth: 116 };
    }
  }

  // Lens Rx col 0 has whitespace header "  " containing consultation Dates
  if (sheetLower === 'lens rx' && colIndex === 0 && h === '') {
    return { category: 'normal', align: 'center', defaultWidth: 108 };
  }

  // Blank spacer / unnamed trailing columns should not inflate to full data width
  if (h === '') {
    return { category: 'normal', align: 'left', defaultWidth: 80 };
  }

  // 1. Patient Identification fields
  if (
    h === 'name' ||
    h === "patient's name" ||
    h === 'surname' ||
    h === 'other name' ||
    h === 'other name(s)'
  ) {
    return { category: 'patient_id', align: 'left', defaultWidth: 180 };
  }
  if (
    h === 'patient no.' ||
    h === 'patient no' ||
    h === 'phone no.' ||
    h === 'phone no' ||
    h === 'contact'
  ) {
    return { category: 'patient_id', align: 'center', defaultWidth: 122 };
  }
  if (
    h === 'tin' ||
    h === 'tin/gh card no.' ||
    h === 'dvla reference number' ||
    h === 'dvla reference #' ||
    h === 'dvla invoice number' ||
    h === 'dvla invoice #' ||
    h === 'licence number'
  ) {
    return { category: 'patient_id', align: 'center', defaultWidth: 144 };
  }
  if (h === 'date of birth') {
    return { category: 'patient_id', align: 'center', defaultWidth: 114 };
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
    return { category: 'patient_id', align: 'left', defaultWidth: 160 };
  }

  // 2. Eye Prescription & Lens Rx fields (consistent compact optical tier)
  if (COMPACT_OPTICAL_HEADERS.has(h)) {
    return {
      category: 'prescription',
      align: 'center',
      defaultWidth: 96
    };
  }
  if (h === 'lensmetry/auto ref' || h === 'refraction result') {
    return { category: 'prescription', align: 'left', defaultWidth: 160 };
  }
  if (h === 'lens type' || h === 'lens treatment') {
    return { category: 'prescription', align: 'left', defaultWidth: 126 };
  }

  // 3. Surgery-related fields
  if (h.includes('post-op')) {
    return { category: 'surgery', align: 'left', defaultWidth: 216 };
  }
  if (h === 'surgery type') {
    return { category: 'surgery', align: 'left', defaultWidth: 160 };
  }
  if (h === 'booking date' || h === 'surgery date') {
    return { category: 'surgery', align: 'center', defaultWidth: 114 };
  }
  if (h === 'payment mode' || h === 'venue') {
    return { category: 'surgery', align: 'left', defaultWidth: 126 };
  }
  if (sheetLower === 'surgery' && h === 'comments') {
    return { category: 'surgery', align: 'left', defaultWidth: 200 };
  }

  // 4. Medication & Treatment fields
  if (h === 'eye drop' || h === 'remarks + medications' || h === 'medication') {
    return { category: 'medication_treatment', align: 'left', defaultWidth: 216 };
  }
  if (h === 'function' || h === 'dosage') {
    return { category: 'medication_treatment', align: 'left', defaultWidth: 196 };
  }
  if (h === 'trade name') {
    return { category: 'medication_treatment', align: 'left', defaultWidth: 156 };
  }
  if (h === 'type' || h === 'availability') {
    return { category: 'medication_treatment', align: 'center', defaultWidth: 96 };
  }
  if (MEDICATION_DRUG_NAMES.has(h)) {
    return { category: 'medication_treatment', align: 'center', defaultWidth: 96 };
  }

  // 5. Clinical Findings & Important Exam Results
  if (
    h === 'chief complaint' ||
    h === 'diagnosis' ||
    h === 'glaucoma hemifield test (ght)' ||
    h === 'pathological remarks'
  ) {
    return { category: 'clinical_finding', align: 'left', defaultWidth: 216 };
  }
  if (
    h === 'medical hx' ||
    h === 'ocular hx' ||
    h === 'drug hx' ||
    h === 'results conclusion' ||
    h === 'result conclusion' ||
    h === 'final results'
  ) {
    return { category: 'clinical_finding', align: 'left', defaultWidth: 148 };
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
    return { category: 'clinical_finding', align: 'left', defaultWidth: 116 };
  }
  if (COMPACT_CLINICAL_METRIC_HEADERS.has(h)) {
    return { category: 'clinical_finding', align: 'center', defaultWidth: 96 };
  }

  // 6. Standard dates, targets & auxiliary fields
  if (h === 'date' || h === 'oct-date' || h === 'time' || h === 'mar') {
    return { category: 'normal', align: 'center', defaultWidth: 108 };
  }
  if (h.includes('target')) {
    return { category: 'normal', align: 'center', defaultWidth: 112 };
  }
  if (h === 'doctor' || h === 'review' || h === 'remarks' || h === 'service type' || h === 'type/place') {
    return { category: 'normal', align: 'left', defaultWidth: 132 };
  }
  if (h === 'sms') {
    return { category: 'normal', align: 'center', defaultWidth: 80 };
  }

  return { category: 'normal', align: 'left', defaultWidth: 112 };
}
