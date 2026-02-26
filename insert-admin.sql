-- =====================================================
-- INSERT FIRST ADMIN ACCOUNT + SET ADMIN SECRET
-- Run this ONCE in Supabase Dashboard → SQL Editor
-- after running fsis.logger.sql
-- =====================================================

-- 1) Set the admin secret (required for admin dashboard to create accounts).
--    Replace 'your-admin-secret-key' with a long random string and keep it safe.
INSERT INTO app_admin_secret (id, secret)
VALUES (1, 'your-admin-secret-key')
ON CONFLICT (id) DO UPDATE SET secret = EXCLUDED.secret;

-- 2) Insert the first admin user (username: admin, password: Admin@123).
--    Since pgcrypto is not used here, password_hash stores the plain password.
INSERT INTO app_users (username, display_name, password_hash, role)
VALUES (
  'admin',
  'Administrator',
  'Admin@123',
  'admin'
)
ON CONFLICT (username) DO NOTHING;

-- To use: log in at index.html with username "admin" and password "Admin@123"
-- (once login is wired to app_users). Then open admin.html and enter the
-- same secret you set above ('your-admin-secret-key') to manage accounts.
