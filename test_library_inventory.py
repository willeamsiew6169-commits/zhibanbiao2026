"""Read-only inventory, legacy endpoint and Excel upload regressions."""
from io import BytesIO
import os
from pathlib import Path
import subprocess
import sys
from types import ModuleType
import unittest
from unittest.mock import MagicMock, Mock, patch

from flask import Blueprint, Flask
from openpyxl import Workbook
from werkzeug.security import generate_password_hash


ROOT = Path(__file__).resolve().parent
BLOCKED_DB = ModuleType("db")
BLOCKED_DB.get_conn = BLOCKED_DB.db_query = Mock(side_effect=AssertionError("No production database"))
with patch.dict(sys.modules, {"db": BLOCKED_DB}):
    import library_web as library
    from hongfa import library_bridge


ADMIN = {"id": 77, "volunteer_id": "CHE-1", "volunteer_name": "测试负责人", "role": "main_admin"}
BOOK = {"item_code": "EB0001", "name": "心经", "category": "经书", "balance": 80,
        "unit": "本", "location": "书架 A", "description": "示例说明"}


class InventoryTests(unittest.TestCase):
    def setUp(self):
        app = Flask(__name__)
        app.secret_key = "offline-library-test"
        app.testing = True
        app.register_blueprint(library.library_bp)
        self.client = app.test_client()
        self.connection = MagicMock()
        self.cursor = self.connection.__enter__.return_value.cursor.return_value.__enter__.return_value
        self.get_conn = self.hook(patch.object(library, "get_conn", return_value=self.connection))
        self.hook(patch.dict(sys.modules, {"db": BLOCKED_DB, "hongfa.library_bridge": library_bridge}))
        self.hook(patch.dict(os.environ, {"LIBRARY_LEGACY_STOCK_ENABLED": "",
                                         library_bridge.CONFIRM_FLAG: "", library_bridge.WRITE_FLAG: ""}))

    def hook(self, patcher):
        obj = patcher.start()
        self.addCleanup(patcher.stop)
        return obj

    def admin(self, *permissions):
        self.hook(patch.object(library, "get_library_admin", return_value=ADMIN))
        self.hook(patch.object(library, "library_admin_can", side_effect=set(permissions).__contains__))

    def test_public_catalog_search_filters_and_paginates_without_login(self):
        self.hook(patch.object(library, "get_library_admin", side_effect=AssertionError("Public lookup must not log in")))
        self.cursor.fetchall.side_effect = [[{"category": "经书"}], [BOOK] * 25]
        self.cursor.fetchone.return_value = {"updated": None}
        response = self.client.get("/library/search?q=心 经&category=经书&page=2&low=1")
        self.assertEqual(response.status_code, 200)
        body = response.get_data(as_text=True)
        self.assertIn("蕉赖观音堂｜藏经阁", body)
        self.assertIn("暂无更新记录", body)
        self.assertIn("负责人入口", body)
        self.assertIn("下一页", body)
        self.assertIn("上一页", body)
        self.assertEqual(body.count('class="inv-card"'), 24)
        self.assertNotIn("只看库存不足", body)
        self.assertNotIn("/library/out", body)
        sql, args = self.cursor.execute.call_args_list[1].args
        self.assertIn("COALESCE(keywords,'') ILIKE", sql)
        self.assertIn("is_active=true", sql)
        self.assertEqual(args, ("心 经", "%心 经%", "%心经%", "%心 经%", "%心 经%",
                                "经书", "经书", False, 25, 24))
        timestamp_sql, timestamp_args = self.cursor.execute.call_args_list[2].args
        self.assertIn("library_audit_logs WHERE action=%s", timestamp_sql)
        self.assertEqual(timestamp_args, (library.EXCEL_UPLOAD_AUDIT,))

    def test_detail_and_qr_do_not_disclose_personal_details_or_inactive_items(self):
        self.cursor.fetchone.return_value = BOOK
        qr = self.client.get("/library/scan/EB0001")
        self.assertEqual(qr.location, "/library/item/EB0001")
        body = self.client.get(qr.location).get_data(as_text=True)
        self.assertIn("示例说明", body)
        self.assertNotIn("<form", body)
        self.assertNotIn("义工编号", body)
        self.assertNotIn("负责人", body)
        self.assertIn("is_active=true", self.cursor.execute.call_args.args[0])
        self.cursor.fetchone.return_value = None
        self.assertEqual(self.client.get("/library/item/EB0002").status_code, 404)
        self.assertEqual(self.client.get("/library/api/item/EB0002").status_code, 404)

    def test_every_retired_route_is_blocked_before_database_and_session_checks(self):
        paths = {
            "/out": ("GET", "POST"), "/out/search": ("GET",), "/out/add": ("POST",),
            "/out/logout": ("GET",), "/out/cart/update": ("POST",), "/out/cart/delete": ("POST",),
            "/out/confirm": ("POST",), "/out/success": ("GET",), "/quick-out/EB0001": ("GET", "POST"),
            "/in": ("GET",), "/in/EB0001": ("GET", "POST"), "/batch-in": ("GET", "POST"),
            "/stocktake": ("GET", "POST"), "/records/cancel/3": ("POST",),
        }
        for path, methods in paths.items():
            for method in methods:
                with self.subTest(path=path, method=method):
                    reply = self.client.open("/library" + path, method=method, data={"quantity": "30"},
                                             headers={"X-Requested-With": "XMLHttpRequest"})
                    self.assertEqual(reply.status_code, 410)
        self.assertEqual(len(paths), len(library.RETIRED_STOCK_ENDPOINTS))
        self.cursor.execute.assert_not_called()
        with patch.dict(os.environ, {"LIBRARY_LEGACY_STOCK_ENABLED": "1"}):
            self.assertEqual(self.client.get("/library/out/logout").status_code, 302)

    def test_history_and_admin_only_catalog_are_backed_by_permissions(self):
        self.assertEqual(self.client.get("/library/records").status_code, 302)
        self.assertEqual(self.client.get("/library/?admin=1&low=1").status_code, 302)
        self.admin("can_view_records")
        self.cursor.fetchall.side_effect = [[], [BOOK]]
        self.cursor.fetchone.return_value = {"updated": "2026-11-23 20:30:01"}
        body = self.client.get("/library/?admin=1&low=1").get_data(as_text=True)
        self.assertIn("只看库存不足", body)
        self.assertIn("/library/records", body)
        self.assertIn("2026-11-23 20:30:01", body)
        self.assertTrue(self.cursor.execute.call_args_list[1].args[1][7])
        self.cursor.fetchall.side_effect = None
        self.cursor.fetchall.return_value = [dict(BOOK, id=1, status="active", transaction_type="out",
                                                   item_name="心经", quantity=2, volunteer_id="CHE-103")]
        history = self.client.get("/library/records").get_data(as_text=True)
        self.assertIn("CHE-103", history)
        self.assertNotIn("/library/records/cancel/", history)

    def test_admin_tiles_count_and_system_permission_isolation(self):
        self.admin("can_manage_items", "can_batch_out", "can_view_records", "can_view_audit_log")
        self.hook(patch.object(library_bridge, "pending_count", return_value=7))
        self.cursor.fetchone.side_effect = [{"variety": 8, "total": 100}, {"updated": None}]
        body = self.client.get("/library/admin/home").get_data(as_text=True)
        for label in ("上传库存 Excel", "弘法领取管理", "库存查询", "系统管理", "测试负责人", "7 待确认"):
            self.assertIn(label, body)
        for label in ("今日领取", "今日入库", "最近记录", "/library/stocktake"):
            self.assertNotIn(label, body)
        system = self.client.get("/library/admin/system").get_data(as_text=True)
        self.assertIn("/library/items", system)
        self.assertIn("/library/admin/audit", system)
        self.assertNotIn("/library/admin/users", system)

    def test_existing_login_and_denied_upload(self):
        self.cursor.fetchone.side_effect = [
            {"c": 1}, {"c": 1}, {"id": 77, "pin_hash": generate_password_hash("9876")}, ADMIN,
        ]
        self.assertIn('name="email"', self.client.get("/library/admin").get_data(as_text=True))
        reply = self.client.post("/library/admin", data={"email": "test@example.invalid", "pin": "9876"})
        self.assertEqual(reply.location, "/library/admin/home")
        with self.client.session_transaction() as user:
            self.assertEqual(user["library_admin_id"], 77)
        self.admin("can_view_records")
        self.hook(patch.object(library, "library_permission_denied", return_value=("denied", 403)))
        self.assertEqual(self.client.post("/library/upload").status_code, 403)

    @staticmethod
    def excel_upload(records=None, headers=None):
        workbook = Workbook()
        workbook.active.append(headers or ["项目名称", "期末余额"])
        for record in records if records is not None else [["心经", 42]]:
            workbook.active.append(record)
        data = BytesIO()
        workbook.save(data)
        data.seek(0)
        return {"file": (data, "fixture.xlsx")}

    def configure_upload_state(self, existing_rows, *, pending=False, mapped_rows=None):
        self.cursor.execute.reset_mock()
        self.cursor.fetchone.reset_mock()
        self.cursor.fetchall.reset_mock()
        self.connection.__enter__.return_value.commit.reset_mock()
        self.get_conn.reset_mock()
        self.cursor.fetchone.side_effect = [
            {"installed": True},
            {"exists": 1} if pending else None,
            {"installed": mapped_rows is not None},
        ]
        self.cursor.fetchall.side_effect = [existing_rows]
        if mapped_rows is not None:
            self.cursor.fetchall.side_effect = [existing_rows, mapped_rows]

    def test_upload_refuses_pending_transfers_without_inventory_or_audit_write(self):
        self.admin("can_manage_items")
        self.cursor.fetchone.side_effect = [{"installed": True}, {"exists": 1}]
        body = self.client.post("/library/upload", data=self.excel_upload()).get_data(as_text=True)
        self.assertIn("请先对账，再上传 Excel", body)
        sql = [call.args[0].lower() for call in self.cursor.execute.call_args_list]
        self.assertIn("pg_advisory_xact_lock", sql[0])
        self.assertFalse(any("insert" in statement for statement in sql))
        self.connection.__enter__.return_value.commit.assert_not_called()

    def test_successful_upload_uses_old_excel_mapping_and_audits_in_same_transaction(self):
        self.admin("can_manage_items")
        category = library.guess_category("心经")
        existing = {
            "id": 1, "item_code": "EB0001", "name": "心经", "category": category,
            "keywords": library.make_keywords("心经"), "unit": library.guess_unit("心经", category),
            "balance": 80, "min_balance": 0, "is_active": True,
        }
        self.configure_upload_state([existing])
        body = self.client.post("/library/upload", data=self.excel_upload()).get_data(as_text=True)
        self.assertIn("新增 0 项，更新 1 项，未变更 0 项", body)
        queries = self.cursor.execute.call_args_list
        upsert = next(call for call in queries if "insert into library_items" in call.args[0])
        self.assertIn("balance = excluded.balance", upsert.args[0])
        self.assertEqual(upsert.args[1][0], "EB0001")
        self.assertEqual(upsert.args[1][5], 42)
        audit = next(call for call in queries if "insert into library_audit_logs" in call.args[0])
        self.assertEqual(audit.args[1][4], library.EXCEL_UPLOAD_AUDIT)
        self.assertEqual(audit.args[1][12].adapted["updated"], 1)
        self.connection.__enter__.return_value.commit.assert_called_once()

    def test_upload_validation_matching_and_safe_inventory_diff(self):
        self.admin("can_manage_items")

        invalid_balances = [
            ["有效零库存", 0], ["空白余额", None], ["文字余额", "很多"],
            ["负数余额", -1], ["非法格式", "1.5"], ["小数数量", 2.5], ["", 4],
        ]
        body = self.client.post(
            "/library/upload",
            data=self.excel_upload(invalid_balances),
        ).get_data(as_text=True)
        for name in ("空白余额", "文字余额", "负数余额", "非法格式", "小数数量", "第8行"):
            self.assertIn(name, body)
        self.assertIn("余额不能为空", body)
        self.assertIn("余额不能为负数", body)
        self.assertIn("余额必须是非负整数", body)
        self.get_conn.assert_not_called()
        self.cursor.execute.assert_not_called()

        self.get_conn.reset_mock()
        duplicate_rows = [["DUP", "重复书", 0], ["dup", " 重复书 ", 1]]
        body = self.client.post(
            "/library/upload",
            data=self.excel_upload(duplicate_rows, ["书籍编号", "项目名称", "期末余额"]),
        ).get_data(as_text=True)
        self.assertIn("第2行《重复书》", body)
        self.assertIn("第3行《重复书》", body)
        self.assertIn("书名重复", body)
        self.assertIn("书籍编号重复", body)
        self.get_conn.assert_not_called()
        self.cursor.execute.assert_not_called()

        category = library.guess_category("心经")
        existing = {
            "id": 1, "item_code": "EB0001", "name": "心经", "category": category,
            "keywords": library.make_keywords("心经"), "unit": library.guess_unit("心经", category),
            "balance": 8, "min_balance": 0, "is_active": True,
        }
        self.configure_upload_state([existing])
        conflict_rows = [["EB0001", "大悲咒", 2], ["NEW1", "心经", 0]]
        body = self.client.post(
            "/library/upload",
            data=self.excel_upload(conflict_rows, ["书籍编号", "项目名称", "期末余额"]),
        ).get_data(as_text=True)
        self.assertIn("Excel 第2行《大悲咒》", body)
        self.assertIn("Excel 第3行《心经》", body)
        self.assertIn("已拒绝覆盖", body)
        self.assertIn("已拒绝匹配冲突", body)
        self.assertFalse(any("insert into library_items" in call.args[0].lower()
                             for call in self.cursor.execute.call_args_list))
        self.assertFalse(any("insert into library_audit_logs" in call.args[0].lower()
                             for call in self.cursor.execute.call_args_list))
        self.connection.__enter__.return_value.commit.assert_not_called()

        mismatched_category = dict(existing, category="书籍")
        mapped = {
            "library_item_id": 1, "item_code": "EB0001", "book_name": "心经", "category": "书籍",
        }
        self.configure_upload_state([mismatched_category], mapped_rows=[mapped])
        body = self.client.post(
            "/library/upload",
            data=self.excel_upload([["心经", 1]]),
        ).get_data(as_text=True)
        self.assertIn("已关联弘法领取", body)
        self.assertFalse(any("insert into library_items" in call.args[0].lower()
                             for call in self.cursor.execute.call_args_list))
        self.connection.__enter__.return_value.commit.assert_not_called()

        def existing_book(item_id, item_code, name, balance):
            category = library.guess_category(name)
            return {
                "id": item_id, "item_code": item_code, "name": name, "category": category,
                "keywords": library.make_keywords(name), "unit": library.guess_unit(name, category),
                "balance": balance, "min_balance": 0, "is_active": True,
            }

        existing_rows = [
            existing_book(1, "EB0001", "心经", 10),
            existing_book(2, "BHF0001", "白话佛法", 4),
            existing_book(3, "JING0009", "大悲咒", 5),
        ]
        self.configure_upload_state(existing_rows)
        reordered_rows = [[261, "大悲咒", 0], ["#NAME?", "心经", 10], [261, "新书", 7]]
        body = self.client.post(
            "/library/upload",
            data=self.excel_upload(
                reordered_rows,
                ["Row num", "项目名称", "July 2026 期末余额 Closing Balance"],
            ),
        ).get_data(as_text=True)
        self.assertIn("新增 1 项，更新 1 项，未变更 1 项", body)
        self.assertIn("另有 1 项现有书籍未出现在 Excel 中", body)
        self.assertIn("白话佛法", body)
        self.assertIn("库存已保留且未停用", body)
        upserts = [call.args[1] for call in self.cursor.execute.call_args_list
                   if "insert into library_items" in call.args[0].lower()]
        by_code = {args[0]: args for args in upserts}
        self.assertEqual(by_code["JING0009"][5], 0)
        self.assertEqual(by_code["BOOK0001"][1], "新书")
        self.assertNotIn("EB0001", by_code)
        self.assertNotIn("BHF0001", by_code)
        audit = next(call for call in self.cursor.execute.call_args_list
                     if "insert into library_audit_logs" in call.args[0].lower())
        self.assertEqual(audit.args[1][12].adapted["added"], 1)
        self.assertEqual(audit.args[1][12].adapted["updated"], 1)
        self.assertEqual(audit.args[1][12].adapted["unchanged"], 1)
        self.assertEqual(audit.args[1][12].adapted["missing"], 1)
        self.connection.__enter__.return_value.commit.assert_called_once()

    def test_phone_and_desktop_keep_two_admin_tiles_per_row(self):
        from playwright.sync_api import sync_playwright

        self.admin("can_manage_items", "can_batch_out", "can_view_records", "can_manage_permissions",
                   "can_view_audit_log")
        self.hook(patch.object(library_bridge, "pending_count", return_value=4))
        self.cursor.fetchone.side_effect = [
            {"variety": 8, "total": 100}, {"updated": "2026-11-23 20:30:01"},
        ]
        admin_html = self.client.get("/library/admin/home").get_data(as_text=True)
        self.cursor.fetchone.side_effect = None
        self.cursor.fetchone.return_value = {"updated": "2026-11-23 20:30:01"}
        titles = [
            dict(BOOK, name="心经", min_balance=10),
            dict(BOOK, name="白话佛法", balance=3, min_balance=5),
            dict(BOOK, name="礼佛大忏悔文", balance=0),
            dict(BOOK, name="礼佛卡", category="法物", balance=24, unit="张"),
            dict(BOOK, name="观音菩萨纪念册", item_code="EB0002", location="书架 B"),
            dict(BOOK, name="大悲咒", min_balance=10, location="书架 C"),
        ]
        self.cursor.fetchall.side_effect = [[{"category": "经书"}, {"category": "法物"}], titles]
        public_html = self.client.get("/library/").get_data(as_text=True)
        self.cursor.fetchone.return_value = titles[1]
        detail_html = self.client.get("/library/item/EB0001").get_data(as_text=True)
        system_html = self.client.get("/library/admin/system").get_data(as_text=True)

        transfer_app = Flask(__name__)
        transfer_app.secret_key = "offline-styles-fixture"
        transfer_bp = Blueprint("library", __name__, url_prefix="/library")
        library_bridge.register_library_confirmation(transfer_bp, lambda: ADMIN, lambda perm: True,
                                                    lambda admin: "admin", Mock())
        transfer_app.register_blueprint(transfer_bp)
        receipt = dict(id=1, book_name="心经", team_name="弘法第一组", quantity=20,
                       receiver_name="测试负责人", received_my="2026-11-27 18:30", status="pending")
        with patch.object(library_bridge, "installed", return_value=True), patch.object(
                library_bridge, "_history", return_value=[receipt]):
            transfer_html = transfer_app.test_client().get("/library/admin/hongfa-transfers").get_data(as_text=True)

        screenshots = ROOT / "output/library-ui"
        screenshots.mkdir(parents=True, exist_ok=True)
        sheets = {f"https://library.test/static/css/{name}": ROOT / "static/css" / name
                  for name in ("toolbox.css", "library-inventory.css")}
        with sync_playwright() as playwright:
            browser = playwright.chromium.launch()
            for width, columns in ((320, 1), (390, 1), (768, 2), (1100, 3)):
                with self.subTest(width=width):
                    page = browser.new_page(viewport={"width": width, "height": 800})
                    page.route("**/*", lambda route: route.fulfill(
                        path=sheets[route.request.url], content_type="text/css"
                    ) if route.request.url in sheets else route.abort())

                    def render(html):
                        page.set_content(html.replace("<head>", '<head><base href="https://library.test/">', 1))
                        self.assertEqual(page.locator("body").evaluate("el => getComputedStyle(el).backgroundColor"),
                                         "rgb(240, 245, 241)")
                        self.assertTrue(page.evaluate("document.documentElement.scrollWidth <= innerWidth"))

                    render(admin_html)
                    tiles = [node.bounding_box() for node in page.locator(".inv-tile").all()]
                    self.assertEqual(len(tiles), 4)
                    self.assertAlmostEqual(tiles[0]["y"], tiles[1]["y"], delta=1)
                    self.assertAlmostEqual(tiles[2]["y"], tiles[3]["y"], delta=1)
                    self.assertGreater(tiles[2]["y"], tiles[0]["y"])
                    if width in (390, 1100):
                        page.screenshot(path=screenshots / f"admin-{width}.png", full_page=True)
                    render(public_html)
                    title = page.locator("h1").bounding_box()
                    link = page.get_by_role("link", name="负责人入口").bounding_box()
                    self.assertGreater(link["x"], title["x"])
                    cards = page.locator("a.inv-card")
                    self.assertEqual(cards.count(), 6)
                    boxes = [node.bounding_box() for node in cards.all()]
                    self.assertEqual(sum(abs(box["y"] - boxes[0]["y"]) < 1 for box in boxes), columns)
                    self.assertEqual(cards.first.get_attribute("href"), "/library/item/EB0001")
                    self.assertEqual(page.get_by_label("书名或关键词").bounding_box()["height"],
                                     page.get_by_label("分类").bounding_box()["height"])
                    expected = {"库存充足": "rgb(22, 101, 52)", "库存不足": "rgb(154, 52, 18)",
                                "零库存": "rgb(153, 27, 27)"}
                    for status, color in expected.items():
                        label = page.locator(".inv-stock-label").get_by_text(status, exact=True).first
                        self.assertEqual(label.evaluate("el => getComputedStyle(el).color"), color)
                    if width in (390, 768, 1100):
                        page.screenshot(path=screenshots / f"inventory-{width}.png", full_page=True)
                    if width in (390, 1100):
                        for name, html in (("detail", detail_html), ("system", system_html), ("transfer", transfer_html)):
                            render(html)
                            self.assertEqual(page.locator(".inv-back").evaluate("el => getComputedStyle(el).position"),
                                             "fixed")
                            page.screenshot(path=screenshots / f"{name}-{width}.png", full_page=True)
                    page.close()
            browser.close()


if __name__ == "__main__":
    suite = unittest.defaultTestLoader.loadTestsFromModule(sys.modules[__name__])
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    status = int(not result.wasSuccessful())
    if "--all" in sys.argv:
        status |= subprocess.run([sys.executable, "-B", "test_hongfa_library_bridge.py", "--all"],
                                 cwd=ROOT).returncode
    sys.exit(status)
