"""Phase 2: authenticated queue claim/ack endpoints. Register on existing schedule_bp.
Never store the agent secret in Supabase or expose it to the browser.
"""
import os
import hmac
import logging
from datetime import datetime, timezone
from flask import request, jsonify
from db import get_conn

LOG = logging.getLogger(__name__)

def register_whatsapp_agent_api(bp):
    def authenticated():
        expected = os.getenv('WHATSAPP_AGENT_TOKEN', '').strip()
        supplied = request.headers.get('Authorization', '')
        return bool(expected) and hmac.compare_digest(supplied, 'Bearer ' + expected)

    @bp.route('/schedule/api/whatsapp-agent/notice-1900/tick', methods=['POST'])
    def wa_notice_1900_tick():
        if not authenticated():
            return jsonify(error='unauthorized'), 401
        try:
            from schedule.services.whatsapp_notice_1900 import enqueue_tomorrow_signup_notice
            # Reuse the exact message currently displayed on the admin signup-notice page.
            from schedule.schedule_web import build_signup_notice_whatsapp
            result = enqueue_tomorrow_signup_notice(build_signup_notice_whatsapp)
            return jsonify(result)
        except Exception:
            LOG.exception('WhatsApp 19:00 signup notice enqueue failed')
            return jsonify(error='notice_enqueue_error'), 500

    @bp.route('/schedule/api/whatsapp-agent/roster-2200/tick', methods=['POST'])
    def wa_roster_2200_tick():
        if not authenticated():
            return jsonify(error='unauthorized'), 401
        try:
            from schedule.services.whatsapp_publish_2200 import run_2200_tick
            return jsonify(run_2200_tick())
        except Exception:
            LOG.exception('WhatsApp 22:00 roster automation failed')
            return jsonify(error='roster_2200_error'), 500

    @bp.route('/schedule/api/whatsapp-agent/morning/tick', methods=['POST'])
    def wa_morning_tick():
        if not authenticated():
            return jsonify(error='unauthorized'), 401
        try:
            from schedule.services.whatsapp_morning_summary import run_morning_tick
            return jsonify(run_morning_tick())
        except Exception:
            LOG.exception('WhatsApp morning summary enqueue failed')
            return jsonify(error='morning_enqueue_error'), 500

    @bp.route('/schedule/api/whatsapp-agent/claim', methods=['POST'])
    def wa_agent_claim():
        if not authenticated():
            return jsonify(error='unauthorized'), 401
        agent_id = str((request.get_json(silent=True) or {}).get('agent_id', 'windows-agent'))[:80]
        try:
            with get_conn() as conn:
                with conn.cursor() as cur:
                    # Only test jobs are enabled in phase 2. No live volunteer messages.
                    cur.execute('''
                        SELECT id, body, destination FROM whatsapp_notification_jobs
                        WHERE status = 'pending' AND message_type = 'test'
                          AND destination = 'TEST_GROUP_ONLY'
                        ORDER BY id ASC LIMIT 1 FOR UPDATE SKIP LOCKED
                    ''')
                    row = cur.fetchone()
                    if not row:
                        return jsonify(job=None)
                    job = dict(row) if isinstance(row, dict) else dict(zip(['id','body','destination'], row))
                    cur.execute('''
                        UPDATE whatsapp_notification_jobs
                        SET status='sending', claimed_at=NOW(), agent_id=%s,
                            attempts=COALESCE(attempts,0)+1
                        WHERE id=%s
                    ''', (agent_id, job['id']))
                conn.commit()
            return jsonify(job=job)
        except Exception:
            LOG.exception('WhatsApp claim failed')
            return jsonify(error='queue_error'), 500

    @bp.route('/schedule/api/whatsapp-agent/ack', methods=['POST'])
    def wa_agent_ack():
        if not authenticated():
            return jsonify(error='unauthorized'), 401
        data = request.get_json(silent=True) or {}
        status = data.get('status')
        if status not in ('sent', 'uncertain', 'failed'):
            return jsonify(error='invalid_status'), 400
        try:
            job_id = int(data['id'])
        except (ValueError, TypeError, KeyError):
            return jsonify(error='invalid_id'), 400
        agent_id = str(data.get('agent_id', 'windows-agent'))[:80]
        error = str(data.get('error', ''))[:1000]
        try:
            with get_conn() as conn:
                with conn.cursor() as cur:
                    cur.execute('''
                        UPDATE whatsapp_notification_jobs
                        SET status=%s, error_message=%s,
                            sent_at=CASE WHEN %s='sent' THEN NOW() ELSE sent_at END
                        WHERE id=%s AND status='sending' AND agent_id=%s
                        RETURNING id
                    ''', (status, error, status, job_id, agent_id))
                    updated = cur.fetchone()
                conn.commit()
            if not updated:
                return jsonify(error='not_claimed_by_agent'), 409
            return jsonify(ok=True)
        except Exception:
            LOG.exception('WhatsApp ack failed')
            return jsonify(error='queue_error'), 500
