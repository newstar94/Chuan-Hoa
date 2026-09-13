DO $$
DECLARE
    table_count integer;
BEGIN
    SELECT count(*) INTO table_count
    FROM information_schema.tables
    WHERE table_schema = 'public'
      AND table_name IN ('admin_integration_idempotency', 'admin_integration_audit');
    IF table_count <> 2 THEN
        RAISE EXCEPTION 'V002 expected two integration tables, got %', table_count;
    END IF;
    IF NOT EXISTS (
        SELECT 1 FROM pg_indexes
        WHERE schemaname = 'public' AND indexname = 'ix_admin_integration_audit_target'
    ) THEN
        RAISE EXCEPTION 'V002 audit target index is missing';
    END IF;
END $$;
