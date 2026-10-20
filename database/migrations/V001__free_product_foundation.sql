-- Chuẩn Hóa free-product database foundation.
-- No users, sessions, trials, billing, entitlements, activation keys or
-- BiddingFlow integration tables are created in a new installation.
CREATE TABLE IF NOT EXISTS schema_migrations (
    version text PRIMARY KEY,
    applied_at_utc timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS audit_events (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    event_type text NOT NULL,
    occurred_at_utc timestamptz NOT NULL DEFAULT now(),
    details jsonb NOT NULL DEFAULT '{}'::jsonb
);

CREATE TABLE outbox_messages (
    sequence_id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    message_id uuid NOT NULL UNIQUE,
    aggregate_type text NOT NULL,
    aggregate_id text NOT NULL,
    event_type text NOT NULL,
    payload jsonb NOT NULL,
    headers jsonb NOT NULL,
    occurred_at_utc timestamptz NOT NULL,
    available_at_utc timestamptz NOT NULL,
    claimed_at_utc timestamptz,
    claim_token uuid,
    completed_at_utc timestamptz,
    dead_lettered_at_utc timestamptz,
    attempt_count integer NOT NULL DEFAULT 0 CHECK (attempt_count >= 0),
    last_error_code text,
    idempotency_key text NOT NULL UNIQUE,
    CHECK (jsonb_typeof(payload) = 'object'),
    CHECK (jsonb_typeof(headers) = 'object'),
    CHECK ((claimed_at_utc IS NULL) = (claim_token IS NULL)),
    CHECK (completed_at_utc IS NULL OR dead_lettered_at_utc IS NULL)
);

CREATE INDEX ix_outbox_messages_pending
    ON outbox_messages(available_at_utc, sequence_id)
    WHERE completed_at_utc IS NULL AND dead_lettered_at_utc IS NULL;

CREATE TABLE idempotency_records (
    id uuid PRIMARY KEY,
    scope_key text NOT NULL,
    idempotency_key_hash bytea NOT NULL,
    request_hash bytea NOT NULL,
    method text NOT NULL,
    path text NOT NULL,
    status text NOT NULL CHECK (status IN ('IN_PROGRESS', 'COMPLETED', 'FAILED_RETRYABLE')),
    owner_token uuid NOT NULL,
    response_status integer,
    response_content_type text,
    response_headers jsonb,
    response_body bytea,
    last_error_code text,
    expires_at_utc timestamptz NOT NULL,
    created_at_utc timestamptz NOT NULL,
    completed_at_utc timestamptz,
    UNIQUE (scope_key, idempotency_key_hash),
    CHECK (octet_length(idempotency_key_hash) = 32),
    CHECK (octet_length(request_hash) = 32),
    CHECK (expires_at_utc > created_at_utc),
    CHECK (response_status IS NULL OR response_status BETWEEN 100 AND 599),
    CHECK (response_headers IS NULL OR jsonb_typeof(response_headers) = 'object'),
    CHECK (
        (status = 'COMPLETED' AND response_status IS NOT NULL AND response_content_type IS NOT NULL AND response_headers IS NOT NULL AND response_body IS NOT NULL AND completed_at_utc IS NOT NULL)
        OR
        (status <> 'COMPLETED' AND response_status IS NULL AND response_content_type IS NULL AND response_headers IS NULL AND response_body IS NULL)
    )
);

CREATE INDEX ix_idempotency_records_expiry
    ON idempotency_records(expires_at_utc);

INSERT INTO schema_migrations(version) VALUES ('V001_free_product_foundation')
ON CONFLICT (version) DO NOTHING;
