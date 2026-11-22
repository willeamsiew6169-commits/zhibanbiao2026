"""弘法当天二维码：服务端 HMAC 签名；无新增数据库写入。"""
import base64
import hashlib
import hmac
import io
from flask import current_app

def qr_token(event_id, day):
    secret=current_app.secret_key
    if not secret:
        raise RuntimeError('未配置固定 SECRET_KEY，无法启用二维码签到')
    key=secret.encode('utf-8') if isinstance(secret,str) else secret
    return hmac.new(key,f'hongfa-qr-v1:{event_id}:{day.isoformat()}'.encode(),hashlib.sha256).hexdigest()

def valid_qr(event, token):
    return bool(token) and len(token)==64 and hmac.compare_digest(qr_token(event['id'],event['event_date']),token)

def qr_png_data(url):
    import qrcode
    img=qrcode.make(url)
    buf=io.BytesIO()
    img.save(buf,format='PNG')
    return base64.b64encode(buf.getvalue()).decode('ascii')
