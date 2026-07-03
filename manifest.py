# manifest.py

from flask import Blueprint

manifest_bp = Blueprint(
    "manifest",
    __name__
)

@manifest_bp.route("/admin-manifest.json")
def manifest():
    return {
        "name": "蕉赖观音堂管理员",
        "short_name": "管理员",
        "start_url": "/admin-home",
        "display": "standalone",
        "background_color": "#7a0000",
        "theme_color": "#7a0000",
        "icons": [
            {
                "src": "/static/icon.png",
                "sizes": "512x512",
                "type": "image/png"
            }
        ]
    }

@manifest_bp.route("/member-manifest.json")
def member_manifest():
    return {
        "name": "佛友月费查询",
        "short_name": "佛友查询",
        "start_url": "/member/",
        "display": "standalone",
        "background_color": "#ffffff",
        "theme_color": "#1976d2",
        "icons": [{
            "src": "/static/member_icon.png?v=1",
            "sizes": "512x512",
            "type": "image/png"
        }]
    }


@manifest_bp.route("/schedule-manifest.json")
def schedule_manifest():
    return {
        "name": "观音堂排班系统",
        "short_name": "排班",
        "start_url": "/schedule/",
        "display": "standalone",
        "background_color": "#ffffff",
        "theme_color": "#43a047",
        "icons": [{
            "src": "/static/schedule_icon.png?v=1",
            "sizes": "512x512",
            "type": "image/png"
        }]
    }

@manifest_bp.route("/volunteer-manifest.json")
def volunteer_manifest():
    return {
        "name": "观音堂义工报名",
        "short_name": "义工报名",
        "start_url": "/volunteer",
        "display": "standalone",
        "background_color": "#ffffff",
        "theme_color": "#00a86b",
        "icons": [{
            "src": "/static/volunteer_icon.png?v=1",
            "sizes": "512x512",
            "type": "image/png"
        }]
    }