-- 仅新增 hongfa_* 表；先在 Supabase SQL Editor 审核后手动执行。
CREATE TABLE IF NOT EXISTS hongfa_teams (
 id BIGSERIAL PRIMARY KEY, name TEXT NOT NULL UNIQUE, active BOOLEAN NOT NULL DEFAULT TRUE,
 created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE TABLE IF NOT EXISTS hongfa_roles (
 id BIGSERIAL PRIMARY KEY, volunteer_id TEXT NOT NULL, team_id BIGINT REFERENCES hongfa_teams(id),
 role TEXT NOT NULL CHECK(role IN ('leader','secretary')),
 active BOOLEAN NOT NULL DEFAULT TRUE,
 CONSTRAINT hongfa_role_scope CHECK ((role='secretary' AND team_id IS NULL) OR (role='leader' AND team_id IS NOT NULL)),
 UNIQUE(volunteer_id,team_id,role)
);
CREATE UNIQUE INDEX IF NOT EXISTS hongfa_one_active_secretary_per_id ON hongfa_roles(volunteer_id) WHERE role='secretary' AND active;
CREATE TABLE IF NOT EXISTS hongfa_locations (
 id BIGSERIAL PRIMARY KEY, name TEXT NOT NULL UNIQUE, maps_url TEXT, active BOOLEAN NOT NULL DEFAULT TRUE
);
CREATE TABLE IF NOT EXISTS hongfa_events (
 id BIGSERIAL PRIMARY KEY, team_id BIGINT NOT NULL REFERENCES hongfa_teams(id),
 location_id BIGINT REFERENCES hongfa_locations(id), event_date DATE NOT NULL,
 starts_at TIME, ends_at TIME, status TEXT NOT NULL DEFAULT 'draft'
 CHECK(status IN ('draft','published','off','cancelled','completed')),
 note TEXT, code_hash TEXT, code_valid_from TIMESTAMPTZ, code_valid_until TIMESTAMPTZ,
 closed_at TIMESTAMPTZ, closed_by TEXT, created_by TEXT NOT NULL,
 created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
 UNIQUE(team_id,event_date),
 CHECK(status IN ('off','cancelled') OR (location_id IS NOT NULL AND starts_at IS NOT NULL AND ends_at IS NOT NULL))
);
CREATE TABLE IF NOT EXISTS hongfa_attendance (
 id BIGSERIAL PRIMARY KEY, event_id BIGINT NOT NULL REFERENCES hongfa_events(id),
 person_key TEXT NOT NULL, volunteer_id TEXT, guest_name TEXT, guest_phone TEXT,
 check_in_at TIMESTAMPTZ NOT NULL DEFAULT now(), check_out_at TIMESTAMPTZ,
 checkout_method TEXT CHECK(checkout_method IN ('self','leader')),
 UNIQUE(event_id,person_key),
 CHECK((volunteer_id IS NOT NULL AND guest_name IS NULL) OR (volunteer_id IS NULL AND guest_name IS NOT NULL))
);
-- 跨活动的同日唯一签到，在服务层事务锁定身份，正式签到实现前需再加数据库级约束/日期字段。
CREATE TABLE IF NOT EXISTS hongfa_items (
 id BIGSERIAL PRIMARY KEY, title TEXT NOT NULL UNIQUE, active BOOLEAN NOT NULL DEFAULT TRUE
);
CREATE TABLE IF NOT EXISTS hongfa_stock_movements (
 id BIGSERIAL PRIMARY KEY, team_id BIGINT NOT NULL REFERENCES hongfa_teams(id),
 item_id BIGINT NOT NULL REFERENCES hongfa_items(id), event_id BIGINT REFERENCES hongfa_events(id),
 quantity INTEGER NOT NULL CHECK(quantity<>0), movement_type TEXT NOT NULL
 CHECK(movement_type IN ('received','distributed','adjustment')),
 created_by TEXT NOT NULL, created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
 request_key TEXT UNIQUE,
 CHECK((movement_type='received' AND quantity>0) OR (movement_type='distributed' AND quantity<0) OR movement_type='adjustment')
);
CREATE INDEX IF NOT EXISTS hongfa_event_day_idx ON hongfa_events(event_date,status);
CREATE INDEX IF NOT EXISTS hongfa_stock_team_idx ON hongfa_stock_movements(team_id,item_id);
-- 示例：由管理员在审核义工编号后，手动授权（不要直接执行示例）。
-- INSERT INTO hongfa_teams(name) VALUES ('俊达组'),('Alex组'),('诗霞组') ON CONFLICT DO NOTHING;
-- INSERT INTO hongfa_roles(volunteer_id,team_id,role) VALUES ('CHE-208',NULL,'secretary');
