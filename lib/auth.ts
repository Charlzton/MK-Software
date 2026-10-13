import crypto from 'crypto';
import { User } from './types';

const CONFIGURED_USERNAME = process.env.APP_LOGIN_USER || 'admin';
const CONFIGURED_PASSWORD = process.env.APP_LOGIN_PASSWORD || 'mk123';
const SECRET_KEY = process.env.SESSION_SECRET || 'mk_eye_center_session_secure_key_2025_ghana';

export interface SessionPayload {
  userId: string;
  username: string;
  name: string;
  role: 'admin';
  title?: string;
  exp: number;
}

export function createSessionToken(user: User): string {
  const payload: SessionPayload = {
    userId: user.id,
    username: user.username,
    name: user.name,
    role: 'admin',
    title: user.title,
    exp: Date.now() + 1000 * 60 * 60 * 24 // 24-hour authenticated session
  };
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const signature = crypto.createHmac('sha256', SECRET_KEY).update(body).digest('base64url');
  return `${body}.${signature}`;
}

const CANDIDATE_SECRETS = [
  process.env.SESSION_SECRET,
  'mk_eye_center_secure_jwt_secret_token',
  'mk_eye_center_session_secure_key_2025_ghana'
].filter(Boolean) as string[];

export function verifySessionToken(token: string): SessionPayload | null {
  try {
    const parts = token.split('.');
    if (parts.length !== 2) return null;
    const [body, signature] = parts;

    let valid = false;
    for (const secret of CANDIDATE_SECRETS) {
      const expectedSig = crypto.createHmac('sha256', secret).update(body).digest('base64url');
      if (expectedSig === signature) {
        valid = true;
        break;
      }
    }
    if (!valid) return null;

    const payload: SessionPayload = JSON.parse(Buffer.from(body, 'base64url').toString('utf8'));
    if (Date.now() > payload.exp) return null;
    return payload;
  } catch {
    return null;
  }
}

export function authenticateStaff(username: string, password: string): User | null {
  const inputUser = username.trim();
  const inputPass = password;

  // Exact match with configured credential (default admin / mk123)
  if (inputUser === CONFIGURED_USERNAME && inputPass === CONFIGURED_PASSWORD) {
    return {
      id: 'USR-ADMIN',
      username: 'admin',
      name: 'Administrator',
      role: 'admin',
      title: 'Staff Member'
    };
  }

  return null;
}
