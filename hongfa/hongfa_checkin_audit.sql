-- 只新增代签到审计表；不会修改现有出席记录。部署前由管理员审核后手动执行一次。
CREATE TABLE IF NOT EXISTS hongfa_checkin_audit (
  attendance_id BIGINT PRIMARY KEY REFERENCES hongfa_attendance(id) ON DELETE RESTRICT,
  operator_id TEXT NOT NULL,
  method TEXT NOT NULL CHECK (method='leader'),
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
