import fs from 'fs';
import path from 'path';

/**
 * Returns the production-ready Google Apps Script Code.gs content.
 */
export function getGoogleAppsScriptCode(): string {
  try {
    const codePath = path.join(process.cwd(), 'google_apps_script', 'Code.gs');
    if (fs.existsSync(codePath)) {
      return fs.readFileSync(codePath, 'utf-8');
    }
  } catch {
    // fallback
  }
  return '';
}

export const GOOGLE_APPS_SCRIPT_TEMPLATE = getGoogleAppsScriptCode();
