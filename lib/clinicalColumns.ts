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
  'va+orx',
  'va+nrx'
]);

const COMPACT_SPHERICAL_METRIC_HEADERS = new Set([
  'sph',
  'cyl',
  'axis',
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
  'temp',
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
 * Deliberate content-proportional column sizing & alignment without coloured category fills:
 * - Reception maintains its existing orderly proportions.
 * - Doctor sizes each column based on actual clinical data shape (ensuring Eye & VA-Rx
 *   keep values like "LE 6/9", "RE 6/20", "RE CF1M" crisp and unclipped, while narrative
 *   and exam findings wrap comfortably without collapsing rows or bloating metrics).
 */
export function getColumnVisualMeta(
  workspace: 'reception' | 'doctor',
  sheetName: string,
  rawHeader: string,
  colIndex: number
): ColumnVisualMeta {
  const h = (rawHeader ?? '').trim().toLowerCase();
  const sheetLower = (sheetName ?? '').trim().toLowerCase();

  // Reception workspace: preserve existing orderly proportions
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
      return { category: 'normal', align: 'left', defaultWidth: 256 };
    }
    return { category: 'normal', align: 'left', defaultWidth: 96 };
  }

  if (sheetLower === 'sheet1') {
    if (colIndex === 0 || h === 'services') {
      return { category: 'normal', align: 'left', defaultWidth: 240 };
    }
    if (colIndex === 1 || h.includes('price')) {
      return { category: 'normal', align: 'right', defaultWidth: 120 };
    }
  }

  if (sheetLower === 'medications' && colIndex === 0 && h === '') {
    return { category: 'normal', align: 'left', defaultWidth: 208 };
  }

  // Lens Rx col 0 has whitespace header "  " containing consultation Dates
  if (sheetLower === 'lens rx' && colIndex === 0 && h === '') {
    return { category: 'normal', align: 'center', defaultWidth: 112 };
  }

  // Truly blank spacer / unnamed trailing columns
  if (h === '') {
    return { category: 'normal', align: 'left', defaultWidth: 84 };
  }

  // 1. Patient Identification fields
  if (
    h === 'name' ||
    h === "patient's name" ||
    h === 'surname' ||
    h === 'other name' ||
    h === 'other name(s)'
  ) {
    return { category: 'normal', align: 'left', defaultWidth: 196 };
  }
  if (
    h === 'patient no.' ||
    h === 'patient no' ||
    h === 'phone no.' ||
    h === 'phone no' ||
    h === 'contact'
  ) {
    return { category: 'normal', align: 'center', defaultWidth: 128 };
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
    return { category: 'normal', align: 'center', defaultWidth: 148 };
  }
  if (h === 'date of birth') {
    return { category: 'normal', align: 'center', defaultWidth: 118 };
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
    return { category: 'normal', align: 'left', defaultWidth: 168 };
  }

  // 2. Eye & Visual Acuity / Prescription fields (wide enough for RE 6/20, LE 6/9, RE CF1M)
  if (h === 'eye' || h === 'va-rx') {
    return { category: 'normal', align: 'center', defaultWidth: 116 };
  }
  if (COMPACT_OPTICAL_HEADERS.has(h)) {
    return { category: 'normal', align: 'center', defaultWidth: 108 };
  }
  if (COMPACT_SPHERICAL_METRIC_HEADERS.has(h)) {
    return { category: 'normal', align: 'center', defaultWidth: 96 };
  }
  if (h === 'lensmetry/auto ref' || h === 'refraction result') {
    return { category: 'normal', align: 'left', defaultWidth: 172 };
  }
  if (h === 'lens type' || h === 'lens treatment') {
    return { category: 'normal', align: 'left', defaultWidth: 136 };
  }

  // 3. Surgery-related fields
  if (h.includes('post-op')) {
    return { category: 'normal', align: 'left', defaultWidth: 224 };
  }
  if (h === 'surgery type') {
    return { category: 'normal', align: 'left', defaultWidth: 168 };
  }
  if (h === 'booking date' || h === 'surgery date') {
    return { category: 'normal', align: 'center', defaultWidth: 116 };
  }
  if (h === 'payment mode' || h === 'venue') {
    return { category: 'normal', align: 'left', defaultWidth: 132 };
  }
  if (sheetLower === 'surgery' && h === 'comments') {
    return { category: 'normal', align: 'left', defaultWidth: 204 };
  }

  // 4. Medication & Treatment fields
  if (h === 'eye drop' || h === 'remarks + medications' || h === 'medication') {
    return { category: 'normal', align: 'left', defaultWidth: 228 };
  }
  if (h === 'function') {
    return { category: 'normal', align: 'left', defaultWidth: 208 };
  }
  if (h === 'dosage') {
    return { category: 'normal', align: 'left', defaultWidth: 164 };
  }
  if (h === 'trade name') {
    return { category: 'normal', align: 'left', defaultWidth: 168 };
  }
  if (h === 'type' || h === 'availability' || MEDICATION_DRUG_NAMES.has(h)) {
    return { category: 'normal', align: 'center', defaultWidth: 96 };
  }

  // 5. Clinical Narratives & Exam Findings
  if (
    h === 'chief complaint' ||
    h === 'diagnosis' ||
    h === 'glaucoma hemifield test (ght)' ||
    h === 'pathological remarks'
  ) {
    return { category: 'normal', align: 'left', defaultWidth: 228 };
  }
  if (
    h === 'medical hx' ||
    h === 'ocular hx' ||
    h === 'drug hx' ||
    h === 'results conclusion' ||
    h === 'result conclusion' ||
    h === 'final results'
  ) {
    return { category: 'normal', align: 'left', defaultWidth: 168 };
  }
  if (h === 'conj/sclera' || h === 'fundus') {
    return { category: 'normal', align: 'left', defaultWidth: 184 };
  }
  if (h === 'eyelid' || h === 'lens') {
    return { category: 'normal', align: 'left', defaultWidth: 136 };
  }
  if (h === 'cornea' || h === 'pupil' || h === 'vitreous') {
    return { category: 'normal', align: 'left', defaultWidth: 112 };
  }

  // 6. Standard dates, targets & auxiliary fields
  if (h === 'date' || h === 'oct-date' || h === 'time' || h === 'mar') {
    return { category: 'normal', align: 'center', defaultWidth: 112 };
  }
  if (h.includes('target')) {
    return { category: 'normal', align: 'center', defaultWidth: 116 };
  }
  if (h === 'doctor' || h === 'review' || h === 'remarks' || h === 'service type' || h === 'type/place') {
    return { category: 'normal', align: 'left', defaultWidth: 136 };
  }
  if (h === 'sms') {
    return { category: 'normal', align: 'center', defaultWidth: 84 };
  }

  return { category: 'normal', align: 'left', defaultWidth: 116 };
}
