DO $$
BEGIN
    IF to_regclass('public.admin_integration_replay_nonces') IS NULL THEN
        RAISE EXCEPTION 'V003 replay nonce table is missing';
    END IF;
    IF NOT EXISTS (
        SELECT 1 FROM pg_indexes
        WHERE schemaname = 'public' AND indexname = 'ix_admin_integration_replay_nonces_expiry'
    ) THEN
        RAISE EXCEPTION 'V003 replay nonce expiry index is missing';
    END IF;
END $$;
