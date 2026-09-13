BEGIN;
DROP TABLE IF EXISTS activation_key_devices;
DROP TABLE IF EXISTS activation_keys;
DROP TABLE IF EXISTS account_device_bindings;
DROP TABLE IF EXISTS account_sessions;
ALTER TABLE users DROP COLUMN IF EXISTS password_hash;
COMMIT;
