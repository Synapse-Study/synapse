"""Synapse auth, fail-closed membership, and local owner-only file persistence.
No service-role key, database password, or personal credential is bundled.
Credentials and settings are stored in config.json with 0o600 permissions (no macOS Keychain prompts).
"""
from __future__ import annotations
import base64
import hashlib
import json
import os
import sys
import threading
import time
from pathlib import Path
import requests

SUPABASE_URL = 'https://ifeitebciscjrhxuyvio.supabase.co'
SUPABASE_PUBLISHABLE_KEY = 'sb_publishable_e40T-57cPm2WMn54tMpVxA_9n1Dlfd-'
QUESTIONS_APP_URL = 'https://glistening-cajeta-0b724a.netlify.app/?synapse=1'
SECRET_KEYS = {
    'GEMINI_API_KEY',
    'GEMINI_NOTES_API_KEY',
    'GEMINI_CHAT_API_KEY',
    'NOTION_API_KEY',
    'SUPABASE_ACCESS_TOKEN',
    'SUPABASE_REFRESH_TOKEN'
}
_LOCK = threading.RLock()


def _secure_chmod(path: Path) -> None:
    """Enforce owner-only (0o600) file permissions on macOS / Linux."""
    if os.name != 'nt':
        try:
            os.chmod(path, 0o600)
        except Exception:
            pass


def get_config_path(config_path=None) -> Path:
    """Resolve per-user Synapse configuration path."""
    if config_path is not None:
        return Path(config_path)
    env_dir = os.environ.get('SYNAPSE_DATA_DIR', '').strip()
    if env_dir:
        return Path(env_dir) / 'config.json'
    if os.name == 'nt' or sys.platform == 'win32':
        app_data = os.environ.get('APPDATA') or (Path.home() / 'AppData' / 'Roaming')
        return Path(app_data) / 'Synapse' / 'config.json'
    return Path.home() / 'Library' / 'Application Support' / 'Synapse' / 'config.json'


def load_secure_config(config_path=None) -> dict:
    """Load Synapse configuration from local file with owner-only permissions."""
    path = get_config_path(config_path)
    with _LOCK:
        if not path.exists():
            return {}
        try:
            raw = path.read_text(encoding='utf-8')
            data = json.loads(raw) if raw.strip() else {}
        except Exception:
            return {}

        if not isinstance(data, dict):
            return {}

        _secure_chmod(path)

        # Backward compatibility and migration for split Gemini API keys
        legacy_gemini = (data.get('GEMINI_API_KEY') or '').strip()
        if legacy_gemini:
            if not data.get('GEMINI_NOTES_API_KEY'):
                data['GEMINI_NOTES_API_KEY'] = legacy_gemini
            if not data.get('GEMINI_CHAT_API_KEY'):
                data['GEMINI_CHAT_API_KEY'] = legacy_gemini

        legacy_sess = data.get('USER_SESSION', {})
        if isinstance(legacy_sess, dict):
            for short, k in [('access_token', 'SUPABASE_ACCESS_TOKEN'), ('refresh_token', 'SUPABASE_REFRESH_TOKEN')]:
                if legacy_sess.get(short) and not data.get(k):
                    data[k] = legacy_sess[short]

        return data


def save_secure_config(data: dict, config_path=None) -> None:
    """Save Synapse configuration atomically to local file with owner-only (0o600) permissions."""
    path = get_config_path(config_path)
    with _LOCK:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix('.tmp')
        tmp.write_text(json.dumps(data, indent=2), encoding='utf-8')
        _secure_chmod(tmp)
        tmp.replace(path)
        _secure_chmod(path)


def _request(method: str, path: str, token: str = None, payload: dict = None) -> dict:
    headers = {'apikey': SUPABASE_PUBLISHABLE_KEY, 'Content-Type': 'application/json'}
    if token:
        headers['Authorization'] = 'Bearer ' + token
    response = requests.request(method, SUPABASE_URL + path, headers=headers, json=payload, timeout=15)
    if response.status_code == 204 or (path.startswith('/auth/v1/logout') and response.status_code in (401, 403)):
        return {}
    try:
        data = response.json()
    except ValueError:
        raise RuntimeError('The sign-in server returned an invalid response.')
    if not response.ok:
        message = data.get('msg') or data.get('message') or data.get('error_description') or 'Account request failed.'
        if path.startswith('/rest/'):
            message = 'Approved-user setup is missing or unavailable. Ask the app administrator.'
        raise RuntimeError(message)
    return data


def _access(token: str):
    identity = _request('GET', '/auth/v1/user', token)
    member = _request('POST', '/rest/v1/rpc/synapse_access', token, {})
    if member.get('allowed') is not True or member.get('user_id') != identity.get('id'):
        raise RuntimeError('This email has not been approved for Synapse.')
    return identity, member


def _persist(session: dict, config_path=None) -> dict:
    identity, member = _access(session['access_token'])
    cfg = load_secure_config(config_path)
    cfg.update(
        AUTH_EMAIL=identity['email'],
        AUTH_ROLE=member['role'],
        AUTH_USER_ID=identity['id'],
        SUPABASE_ACCESS_TOKEN=session['access_token'],
        SUPABASE_REFRESH_TOKEN=session['refresh_token']
    )
    save_secure_config(cfg, config_path)
    return {
        'ok': True,
        'email': identity['email'],
        'role': member['role'],
        'user_id': identity['id'],
        'access_token': session['access_token'],
        'refresh_token': session['refresh_token'],
        'questions_app_url': QUESTIONS_APP_URL
    }


def sign_in_synapse(email: str, password: str, config_path=None) -> dict:
    if not str(email or '').strip() or not password:
        return {'ok': False, 'error': 'Enter your email and password.'}
    try:
        data = _request('POST', '/auth/v1/token?grant_type=password', payload={'email': email.strip().lower(), 'password': password})
        return _persist(data, config_path)
    except Exception as e:
        return {'ok': False, 'error': str(e)}


def sign_up_synapse(email: str, password: str, config_path=None) -> dict:
    clean_email = str(email or '').strip().lower()
    if not clean_email or not password:
        return {'ok': False, 'error': 'Enter your email and password.'}
    try:
        res = _request('POST', '/auth/v1/signup', payload={'email': clean_email, 'password': password})
        if res.get('access_token'):
            return _persist(res, config_path)
        signin = sign_in_synapse(clean_email, password, config_path=config_path)
        if signin.get('ok'):
            return signin
        return {
            'ok': False,
            'error': signin.get('error') or 'Registration completed on server, but authentication check failed. Please click "Sign In & Continue".'
        }
    except Exception as e:
        err_msg = str(e)
        if any(term in err_msg.lower() for term in ['already registered', 'already exists', 'user already']):
            signin = sign_in_synapse(clean_email, password, config_path=config_path)
            if signin.get('ok'):
                return signin
            return {'ok': False, 'error': 'Account already registered. Please enter your password and click "Sign In & Continue".'}
        return {'ok': False, 'error': err_msg}


def clear_session_from_config(config_path=None) -> None:
    cfg = load_secure_config(config_path)
    for k in ['AUTH_EMAIL', 'AUTH_ROLE', 'AUTH_USER_ID', 'SUPABASE_ACCESS_TOKEN', 'SUPABASE_REFRESH_TOKEN']:
        cfg[k] = ''
    cfg.pop('USER_SESSION', None)
    save_secure_config(cfg, config_path)


def sign_out_synapse(config_path=None) -> dict:
    try:
        cfg = load_secure_config(config_path)
        token = cfg.get('SUPABASE_ACCESS_TOKEN')
        if token:
            _request('POST', '/auth/v1/logout?scope=local', token, {})
        clear_session_from_config(config_path)
        return {'ok': True}
    except Exception as e:
        return {'ok': False, 'error': 'Could not finish sign-out: ' + str(e)}


def _expiry(token: str) -> float:
    try:
        p = token.split('.')[1]
        return float(json.loads(base64.urlsafe_b64decode(p + '=' * (-len(p) % 4)))['exp'])
    except Exception:
        return 0.0


def get_valid_session(config_path=None) -> dict:
    try:
        with _LOCK:
            cfg = load_secure_config(config_path)
            access = cfg.get('SUPABASE_ACCESS_TOKEN', '')
            refresh = cfg.get('SUPABASE_REFRESH_TOKEN', '')
            if not access or not refresh:
                return {'ok': False, 'error': 'Sign in to Synapse.'}
            if _expiry(access) < time.time() + 15:
                data = _request('POST', '/auth/v1/token?grant_type=refresh_token', payload={'refresh_token': refresh})
                return _persist(data, config_path)
            identity, member = _access(access)
            return {
                'ok': True,
                'email': identity['email'],
                'role': member['role'],
                'user_id': identity['id'],
                'access_token': access,
                'refresh_token': refresh,
                'questions_app_url': QUESTIONS_APP_URL
            }
    except Exception as e:
        return {'ok': False, 'error': 'Sign in again: ' + str(e)}


def accept_question_session(access_token: str, refresh_token: str, config_path=None) -> dict:
    try:
        with _LOCK:
            cfg = load_secure_config(config_path)
            identity, member = _access(access_token)
            if not cfg.get('AUTH_USER_ID') or identity['id'] != cfg['AUTH_USER_ID']:
                raise RuntimeError('The refreshed account does not match Synapse.')
            return _persist({'access_token': access_token, 'refresh_token': refresh_token}, config_path)
    except Exception as e:
        return {'ok': False, 'error': str(e)}
