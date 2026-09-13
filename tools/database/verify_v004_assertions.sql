\set ON_ERROR_STOP on

DO $$
DECLARE
    active_device_constraint text;
    activation_constraint text;
BEGIN
    IF to_regclass('public.account_sessions') IS NULL
       OR to_regclass('public.account_device_bindings') IS NULL
       OR to_regclass('public.activation_keys') IS NULL
       OR to_regclass('public.activation_key_devices') IS NULL THEN
        RAISE EXCEPTION 'V004 required tables are missing';
    END IF;

    SELECT conname INTO active_device_constraint
    FROM pg_constraint
    WHERE conrelid = 'public.account_device_bindings'::regclass
      AND contype = 'u';
    IF active_device_constraint IS NULL THEN
        RAISE EXCEPTION 'V004 must enforce one active device binding per account';
    END IF;

    SELECT conname INTO activation_constraint
    FROM pg_constraint
    WHERE conrelid = 'public.activation_key_devices'::regclass
      AND contype = 'u';
    IF activation_constraint IS NULL THEN
        RAISE EXCEPTION 'V004 must enforce one activation per key/device';
    END IF;

    IF NOT EXISTS (
        SELECT 1 FROM pg_indexes
        WHERE schemaname = 'public'
          AND tablename = 'activation_keys'
          AND indexdef ILIKE '%key_hash%'
    ) THEN
        RAISE EXCEPTION 'V004 key hash lookup index is missing';
    END IF;
END $$;

SELECT 'V004_ASSERTIONS_PASS' AS result;
