BEGIN;

CREATE TABLE admin_integration_idempotency (
    id uuid PRIMARY KEY,
    client_id text NOT NULL,
    idempotency_key text NOT NULL,
    request_hash bytea NOT NULL,
    response_status integer NOT NULL CHECK (response_status BETWEEN 100 AND 599),
    response_body jsonb NOT NULL CHECK (jsonb_typeof(response_body) = 'object'),
    created_at_utc timestamptz NOT NULL,
    UNIQUE (client_id, idempotency_key),
    CHECK (octet_length(request_hash) = 32)
);

CREATE TABLE admin_integration_audit (
    sequence_id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    event_id uuid NOT NULL UNIQUE,
    occurred_at_utc timestamptz NOT NULL,
    source_application text NOT NULL,
    external_actor_id text NOT NULL,
    target_type text NOT NULL,
    target_id text NOT NULL,
    action_code text NOT NULL,
    result_code text NOT NULL,
    correlation_id uuid NOT NULL,
    metadata jsonb NOT NULL CHECK (jsonb_typeof(metadata) = 'object')
);

CREATE INDEX ix_admin_integration_audit_target
    ON admin_integration_audit(target_type, target_id, occurred_at_utc DESC);

COMMIT;
