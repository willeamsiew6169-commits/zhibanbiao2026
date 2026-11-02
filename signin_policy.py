"""签到限制设置：由 Supabase 持久化，首次部署须执行配套 SQL。"""
from db import db_query

POLICY_KEYS = {
    'duty': 'signin_require_duty_assignment',
    'cleaning': 'signin_require_cleaning_signup',
    'meal': 'signin_require_meal_signup',
}

def get_signin_policies():
    rows = db_query('SELECT key, enabled FROM signin_policy_settings WHERE key = ANY(%s)',
                    (list(POLICY_KEYS.values()),), fetchall=True) or []
    values = {row['key']: bool(row['enabled']) for row in rows}
    return {role: values.get(key, False) for role, key in POLICY_KEYS.items()}

def set_signin_policy(role, enabled):
    if role not in POLICY_KEYS:
        raise ValueError('Unknown signin policy')
    db_query('''INSERT INTO signin_policy_settings (key, enabled, updated_at)
                VALUES (%s, %s, now())
                ON CONFLICT (key) DO UPDATE
                SET enabled=EXCLUDED.enabled, updated_at=now()''',
             (POLICY_KEYS[role], bool(enabled)))
