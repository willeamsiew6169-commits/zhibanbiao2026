# manifest.py

from flask import Blueprint, jsonify


manifest_bp = Blueprint(
    "manifest",
    __name__
)


# =========================================================
# 1. 管理员系统
# =========================================================

@manifest_bp.route("/admin-manifest.json")
def admin_manifest():
    return jsonify({
        "name": "蕉赖观音堂管理员",
        "short_name": "管理员",
        "start_url": "/admin-home",
        "scope": "/",
        "display": "standalone",
        "background_color": "#7a0000",
        "theme_color": "#7a0000",

        "icons": [
            {
                "src": "/static/icon_192.png?v=5",
                "sizes": "192x192",
                "type": "image/png",
                "purpose": "any"
            },
            {
                "src": "/static/icon_512.png?v=5",
                "sizes": "512x512",
                "type": "image/png",
                "purpose": "any"
            }
        ]
    })


# =========================================================
# 2. 佛友月费查询
# =========================================================

@manifest_bp.route("/member-manifest.json")
def member_manifest():
    return jsonify({
        "name": "佛友月费查询",
        "short_name": "佛友查询",
        "start_url": "/member/query-login",
        "scope": "/member/",
        "display": "standalone",
        "background_color": "#ffffff",
        "theme_color": "#1976d2",

        "icons": [
            {
                "src": "/static/member_icon_192.png?v=5",
                "sizes": "192x192",
                "type": "image/png",
                "purpose": "any"
            },
            {
                "src": "/static/member_icon_512.png?v=5",
                "sizes": "512x512",
                "type": "image/png",
                "purpose": "any"
            }
        ]
    })


# =========================================================
# 3. 排班系统
# =========================================================

@manifest_bp.route("/schedule-manifest.json")
def schedule_manifest():
    return jsonify({
        "name": "观音堂排班系统",
        "short_name": "排班",
        "start_url": "/schedule/",
        "scope": "/schedule/",
        "display": "standalone",
        "background_color": "#ffffff",
        "theme_color": "#43a047",

        "icons": [
            {
                "src": "/static/schedule_icon_192.png?v=5",
                "sizes": "192x192",
                "type": "image/png",
                "purpose": "any"
            },
            {
                "src": "/static/schedule_icon_512.png?v=5",
                "sizes": "512x512",
                "type": "image/png",
                "purpose": "any"
            }
        ]
    })


# =========================================================
# 4. 义工报名
# =========================================================

@manifest_bp.route("/volunteer-manifest.json")
def volunteer_manifest():
    return jsonify({
        "name": "观音堂义工报名",
        "short_name": "义工报名",
        "start_url": "/volunteer",
        "scope": "/volunteer",
        "display": "standalone",
        "background_color": "#ffffff",
        "theme_color": "#00a86b",

        "icons": [
            {
                "src": "/static/volunteer_icon_192.png?v=5",
                "sizes": "192x192",
                "type": "image/png",
                "purpose": "any"
            },
            {
                "src": "/static/volunteer_icon_512.png?v=5",
                "sizes": "512x512",
                "type": "image/png",
                "purpose": "any"
            }
        ]
    })


# =========================================================
# 5. 藏经阁系统
# =========================================================

@manifest_bp.route("/library-manifest.json")
def library_manifest():
    return jsonify({
        "name": "藏经阁系统",
        "short_name": "藏经阁",
        "start_url": "/library/",
        "scope": "/library/",
        "display": "standalone",
        "background_color": "#ffffff",
        "theme_color": "#8B4513",

        "icons": [
            {
                "src": "/static/library_icon_192.png?v=5",
                "sizes": "192x192",
                "type": "image/png",
                "purpose": "any"
            },
            {
                "src": "/static/library_icon_512.png?v=5",
                "sizes": "512x512",
                "type": "image/png",
                "purpose": "any"
            }
        ]
    })


# =========================================================
# 6. 佛学班系统
# =========================================================

@manifest_bp.route("/dharma-class-manifest.json")
def dharma_class_manifest():
    return jsonify({
        "name": "蕉赖佛学班系统",
        "short_name": "佛学班",
        "start_url": "/class/",
        "scope": "/class/",
        "display": "standalone",
        "background_color": "#fffaf0",
        "theme_color": "#f6c54e",

        "icons": [
            {
                "src": "/static/dharma_icon_192.png?v=5",
                "sizes": "192x192",
                "type": "image/png",
                "purpose": "any"
            },
            {
                "src": "/static/dharma_icon_512.png?v=5",
                "sizes": "512x512",
                "type": "image/png",
                "purpose": "any"
            }
        ]
    })


# =========================================================
# 7. 财政系统
# =========================================================

@manifest_bp.route("/finance-manifest.json")
def finance_manifest():
    return jsonify({
        "name": "观音堂财政系统",
        "short_name": "财政",
        "start_url": "/finance/",
        "scope": "/finance/",
        "display": "standalone",
        "background_color": "#fffdf8",
        "theme_color": "#f4b400",

        "icons": [
            {
                "src": "/static/finance_icon_192.png?v=5",
                "sizes": "192x192",
                "type": "image/png",
                "purpose": "any"
            },
            {
                "src": "/static/finance_icon_512.png?v=5",
                "sizes": "512x512",
                "type": "image/png",
                "purpose": "any"
            }
        ]
    })