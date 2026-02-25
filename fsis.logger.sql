-- =====================================================
-- BFP INSPECTION MANAGEMENT SYSTEM — SUPABASE SQL SCHEMA
-- Run this in Supabase Dashboard → SQL Editor
-- =====================================================

-- Enable extensions
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- ========================
-- TABLE: inspection_logbook
-- ========================
CREATE TABLE IF NOT EXISTS inspection_logbook (
  id              UUID DEFAULT uuid_generate_v4() PRIMARY KEY,
  io_number       VARCHAR(50) NOT NULL,
  owner_name      VARCHAR(255) NOT NULL,
  business_name   VARCHAR(255) NOT NULL,
  address         TEXT NOT NULL,
  date_inspected  DATE NOT NULL,
  fsic_number     VARCHAR(50) NOT NULL,
  inspected_by    VARCHAR(255),
  latitude        DECIMAL(10,8) NULL,
  longitude       DECIMAL(11,8) NULL,
  photo_url       TEXT NULL,
  photo_taken_at  TEXT NULL,
  created_at      TIMESTAMPTZ DEFAULT NOW() NOT NULL,
  updated_at      TIMESTAMPTZ DEFAULT NOW() NOT NULL
);

-- Schema upgrades for existing projects (safe to run multiple times)
ALTER TABLE inspection_logbook
  ADD COLUMN IF NOT EXISTS latitude DECIMAL(10,8) NULL;
ALTER TABLE inspection_logbook
  ADD COLUMN IF NOT EXISTS longitude DECIMAL(11,8) NULL;
ALTER TABLE inspection_logbook
  ADD COLUMN IF NOT EXISTS photo_url TEXT NULL;
ALTER TABLE inspection_logbook
  ADD COLUMN IF NOT EXISTS photo_taken_at TEXT NULL;

-- Auto-update updated_at on row change
CREATE OR REPLACE FUNCTION update_updated_at()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = NOW();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_inspection_updated_at ON inspection_logbook;
CREATE TRIGGER trg_inspection_updated_at
  BEFORE UPDATE ON inspection_logbook
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();

-- Index for geospatial queries on latitude/longitude
CREATE INDEX IF NOT EXISTS idx_inspection_lat_long ON inspection_logbook (latitude, longitude);

-- Note: For PostgreSQL with PostGIS extension, you can alternatively use:
-- ALTER TABLE inspection_logbook ADD COLUMN IF NOT EXISTS location GEOGRAPHY(Point, 4326);
-- CREATE INDEX IF NOT EXISTS idx_inspection_location ON inspection_logbook USING GIST (location);

-- ==============================
-- TABLE: fsec_building_plan_logbook
-- ==============================
CREATE TABLE IF NOT EXISTS fsec_building_plan_logbook (
  id               UUID DEFAULT uuid_generate_v4() PRIMARY KEY,
  owner_name       VARCHAR(255) NOT NULL,
  proposed_project VARCHAR(255) NOT NULL,
  address          TEXT NOT NULL,
  date             DATE NOT NULL,
  contact_number   VARCHAR(30) NOT NULL,
  created_at       TIMESTAMPTZ DEFAULT NOW() NOT NULL,
  updated_at       TIMESTAMPTZ DEFAULT NOW() NOT NULL
);

DROP TRIGGER IF EXISTS trg_fsec_updated_at ON fsec_building_plan_logbook;
CREATE TRIGGER trg_fsec_updated_at
  BEFORE UPDATE ON fsec_building_plan_logbook
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();

-- ========================
-- TABLE: app_users (for login / future auth)
-- ========================
CREATE TABLE IF NOT EXISTS app_users (
  id            UUID DEFAULT uuid_generate_v4() PRIMARY KEY,
  username      VARCHAR(80) UNIQUE NOT NULL,
  display_name  VARCHAR(120),
  password_hash TEXT NOT NULL,
  role          TEXT NOT NULL DEFAULT 'user',
  created_at    TIMESTAMPTZ DEFAULT NOW() NOT NULL,
  updated_at    TIMESTAMPTZ DEFAULT NOW() NOT NULL,
  CHECK (role IN ('user', 'admin'))
);

DROP TRIGGER IF EXISTS trg_app_users_updated_at ON app_users;
CREATE TRIGGER trg_app_users_updated_at
  BEFORE UPDATE ON app_users
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();

-- Login helper: validate username/password against app_users
DROP FUNCTION IF EXISTS app_login(TEXT, TEXT);
CREATE FUNCTION app_login(p_username TEXT, p_password TEXT)
RETURNS TABLE (
  id UUID,
  username TEXT,
  display_name TEXT,
  role TEXT
)
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
DECLARE
  u app_users%ROWTYPE;
BEGIN
  SELECT u1.* INTO u
  FROM app_users AS u1
  WHERE u1.username = p_username;

  IF NOT FOUND THEN
    RETURN;
  END IF;

  -- Plain-text comparison (no pgcrypto on this project)
  IF u.password_hash IS NULL OR u.password_hash <> p_password THEN
    RETURN;
  END IF;

  RETURN QUERY
    SELECT
      u.id::uuid,
      u.username::text,
      u.display_name::text,
      u.role::text;
END;
$$;

-- ========================
-- ROW LEVEL SECURITY (RLS)
-- ========================
ALTER TABLE inspection_logbook ENABLE ROW LEVEL SECURITY;
ALTER TABLE fsec_building_plan_logbook ENABLE ROW LEVEL SECURITY;
ALTER TABLE app_users ENABLE ROW LEVEL SECURITY;

-- Allow anon key full access (app uses anon key without auth)
DROP POLICY IF EXISTS "Allow anon all inspection_logbook" ON inspection_logbook;
CREATE POLICY "Allow anon all inspection_logbook" ON inspection_logbook
  FOR ALL TO anon USING (true) WITH CHECK (true);

DROP POLICY IF EXISTS "Allow anon all fsec_building_plan_logbook" ON fsec_building_plan_logbook;
CREATE POLICY "Allow anon all fsec_building_plan_logbook" ON fsec_building_plan_logbook
  FOR ALL TO anon USING (true) WITH CHECK (true);

-- app_users: no anon access (keeps password_hash private).

-- ========================
-- ADMIN: secret + RPCs for dashboard
-- ========================
CREATE TABLE IF NOT EXISTS app_admin_secret (
  id    INT PRIMARY KEY DEFAULT 1,
  secret TEXT NOT NULL,
  CONSTRAINT single_row CHECK (id = 1)
);

ALTER TABLE app_admin_secret ENABLE ROW LEVEL SECURITY;
-- No policies: anon cannot read. Only SECURITY DEFINER functions can read.

-- Create new app user (called from admin dashboard with secret).
-- NOTE: For simplicity this stores the password as plain text in password_hash.
-- If you later enable pgcrypto, you can switch to crypt()/gen_salt().
DROP FUNCTION IF EXISTS create_app_user(TEXT, TEXT, TEXT, TEXT, TEXT);
CREATE FUNCTION create_app_user(
  admin_secret TEXT,
  p_username TEXT,
  p_display_name TEXT,
  p_password_plain TEXT,
  p_role TEXT DEFAULT 'user'
)
RETURNS UUID
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
DECLARE
  v_secret TEXT;
  v_id UUID;
BEGIN
  IF p_username IS NULL OR length(trim(p_username)) = 0 THEN
    RAISE EXCEPTION 'Username is required';
  END IF;
  IF p_password_plain IS NULL OR length(p_password_plain) < 4 THEN
    RAISE EXCEPTION 'Password must be at least 4 characters';
  END IF;
  IF p_role NOT IN ('user', 'admin') THEN
    RAISE EXCEPTION 'Role must be user or admin';
  END IF;

  SELECT secret INTO v_secret FROM app_admin_secret WHERE id = 1 LIMIT 1;
  IF v_secret IS NULL THEN
    RAISE EXCEPTION 'Admin secret not configured';
  END IF;
  IF v_secret <> admin_secret THEN
    RAISE EXCEPTION 'Invalid admin secret';
  END IF;

  INSERT INTO app_users (username, display_name, password_hash, role)
  VALUES (
    trim(p_username),
    nullif(trim(p_display_name), ''),
    p_password_plain,
    p_role
  )
  RETURNING id INTO v_id;
  RETURN v_id;
END;
$$;

DROP FUNCTION IF EXISTS update_app_user(TEXT, UUID, TEXT, TEXT, TEXT, TEXT);
CREATE FUNCTION update_app_user(
  admin_secret TEXT,
  p_user_id UUID,
  p_username TEXT,
  p_display_name TEXT,
  p_password_plain TEXT,
  p_role TEXT
)
RETURNS VOID
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
DECLARE
  v_secret TEXT;
BEGIN
  IF p_user_id IS NULL THEN
    RAISE EXCEPTION 'User id is required';
  END IF;

  IF p_role IS NOT NULL AND p_role NOT IN ('user', 'admin') THEN
    RAISE EXCEPTION 'Role must be user or admin';
  END IF;

  SELECT secret INTO v_secret FROM app_admin_secret WHERE id = 1 LIMIT 1;
  IF v_secret IS NULL OR v_secret <> admin_secret THEN
    RAISE EXCEPTION 'Invalid admin secret';
  END IF;

  UPDATE app_users
  SET
    username = COALESCE(NULLIF(trim(p_username), ''), username),
    display_name = NULLIF(trim(COALESCE(p_display_name, display_name::text)), ''),
    password_hash = COALESCE(NULLIF(p_password_plain, ''), password_hash),
    role = COALESCE(p_role, role),
    updated_at = NOW()
  WHERE id = p_user_id;
END;
$$;

DROP FUNCTION IF EXISTS list_app_users(TEXT);
-- List users for admin dashboard (id, username, display_name, role, created_at only; no password).
CREATE FUNCTION list_app_users(admin_secret TEXT)
RETURNS TABLE (
  id UUID,
  username TEXT,
  display_name TEXT,
  role TEXT,
  created_at TIMESTAMPTZ
)
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
DECLARE
  v_secret TEXT;
BEGIN
  SELECT app_admin_secret.secret INTO v_secret FROM app_admin_secret WHERE app_admin_secret.id = 1 LIMIT 1;
  IF v_secret IS NULL OR v_secret <> admin_secret THEN
    RAISE EXCEPTION 'Invalid admin secret';
  END IF;
  RETURN QUERY
    SELECT
      u.id::uuid        AS id,
      u.username::text  AS username,
      u.display_name::text AS display_name,
      u.role::text      AS role,
      u.created_at::timestamptz AS created_at
    FROM app_users u
    ORDER BY u.created_at DESC;
END;
$$;
