Finance V6 Final Stable Candidate — Phase 1 Architecture Foundation
==================================================================

本包完成：
1. 新增 finance_common.py
   - 金额、马来西亚日期、月份、当前分会、当前财政用户
   - Month Close Lock 共用 Helper
   - JSON-safe 转换与请求 IP
2. 新增 finance_audit.py
   - write_finance_audit()
   - /finance/audit/
   - /finance/audit/<id>
3. app.py
   - finance_web_old -> finance_web
   - finance_month_end_old -> finance_month_end
   - 注册 finance_audit_bp
4. finance_web.py
   - 共用 Helper 改从 finance_common 导入
   - 作废记录写入 Audit Trail
5. finance_month_end.py
   - money / Month Lock 改从 finance_common 导入
   - Month Close 写入 Audit Trail

安装：
A. 先备份原文件。
B. 在 Supabase SQL Editor 执行 FINANCE_V6_ARCHITECTURE_SQL.sql。
C. 复制以下文件覆盖／新增：
   app.py
   finance_web.py
   finance_month_end.py
   finance_common.py
   finance_audit.py
D. 保留原本 member_web.py、helpers.py、db.py 和 toolbox.css。
E. 重新启动 Flask。

测试：
1. 财政登录后打开 /finance/audit/
2. 作废一笔未月结测试记录，Audit 应出现 finance_records / cancel。
3. 完成一个测试月份 Month Close，Audit 应出现 month_close / close。
4. 已月结月份仍须维持 Month Lock。

重要：
- 本包是最终版重构的第一阶段，不含历史 Excel Import 与 Finance Health。
- finance_audit_logs 尚无资料时，Audit 页面显示“暂无审计记录”属正常。
- Production 使用前请确认 app.secret_key 改为环境变量；本阶段未强制修改现有登录。
