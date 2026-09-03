BEGIN;

CREATE EXTENSION IF NOT EXISTS btree_gist;

CREATE TYPE account_status AS ENUM ('PENDING_VERIFICATION', 'ACTIVE', 'SUSPENDED', 'REVOKED');
CREATE TYPE device_status AS ENUM ('PENDING_VERIFICATION', 'ACTIVE', 'SUSPENDED', 'REVOKED', 'BLOCKED');
CREATE TYPE device_assurance AS ENUM ('TPM_ATTESTED', 'TPM_BACKED', 'SOFTWARE_KEY', 'UNKNOWN');
CREATE TYPE trial_kind AS ENUM ('LAUNCH', 'PERSONAL');
CREATE TYPE trial_campaign_status AS ENUM ('DRAFT', 'SCHEDULED', 'ACTIVE', 'ENDED', 'TERMINATED');
CREATE TYPE trial_grant_status AS ENUM ('ACTIVE', 'EXPIRED', 'CONVERTED', 'REVOKED');
CREATE TYPE entitlement_source AS ENUM ('PURCHASE', 'ORGANIZATION_CONTRACT', 'MANUAL_GRANT', 'LAUNCH_TRIAL', 'PERSONAL_TRIAL', 'ADMIN_EXTENSION');
CREATE TYPE entitlement_status AS ENUM ('SCHEDULED', 'ACTIVE', 'EXPIRED', 'REVOKED');
CREATE TYPE offer_status AS ENUM ('DRAFT', 'SCHEDULED', 'PUBLISHED', 'RETIRED');
CREATE TYPE quote_status AS ENUM ('ACTIVE', 'CONSUMED', 'EXPIRED', 'CANCELLED');
CREATE TYPE order_status AS ENUM ('CREATED', 'PENDING_PROVIDER', 'AWAITING_PAYMENT', 'VERIFIED_PAID', 'ACTIVATED', 'EXPIRED', 'CANCELLED', 'FAILED', 'PARTIALLY_REFUNDED', 'REFUNDED', 'MANUAL_REVIEW');
CREATE TYPE approval_status AS ENUM ('DRAFT', 'SUBMITTED', 'APPROVED', 'REJECTED', 'EXPIRED', 'CANCELLED', 'EXECUTING', 'EXECUTED', 'FAILED');

CREATE TABLE organizations (
    id uuid PRIMARY KEY,
    display_name text NOT NULL,
    status text NOT NULL CHECK (status IN ('ACTIVE', 'SUSPENDED', 'CLOSED')),
    version bigint NOT NULL DEFAULT 1 CHECK (version > 0),
    created_at_utc timestamptz NOT NULL,
    updated_at_utc timestamptz NOT NULL
);

CREATE TABLE users (
    id uuid PRIMARY KEY,
    normalized_email text NOT NULL UNIQUE,
    display_name text NOT NULL,
    status account_status NOT NULL,
    email_verified_at_utc timestamptz,
    version bigint NOT NULL DEFAULT 1 CHECK (version > 0),
    created_at_utc timestamptz NOT NULL,
    updated_at_utc timestamptz NOT NULL
);

CREATE TABLE identities (
    id uuid PRIMARY KEY,
    user_id uuid NOT NULL REFERENCES users(id),
    issuer text NOT NULL,
    provider_subject text NOT NULL,
    created_at_utc timestamptz NOT NULL,
    UNIQUE (issuer, provider_subject)
);

CREATE TABLE organization_memberships (
    id uuid PRIMARY KEY,
    organization_id uuid NOT NULL REFERENCES organizations(id),
    user_id uuid NOT NULL REFERENCES users(id),
    role_code text NOT NULL,
    status text NOT NULL CHECK (status IN ('INVITED', 'ACTIVE', 'SUSPENDED', 'REVOKED')),
    created_at_utc timestamptz NOT NULL,
    revoked_at_utc timestamptz,
    UNIQUE (organization_id, user_id, role_code)
);

CREATE TABLE devices (
    id uuid PRIMARY KEY,
    user_id uuid NOT NULL REFERENCES users(id),
    organization_id uuid REFERENCES organizations(id),
    display_name text NOT NULL,
    public_key_jwk jsonb NOT NULL,
    public_key_thumbprint text NOT NULL UNIQUE,
    assurance device_assurance NOT NULL,
    status device_status NOT NULL,
    version bigint NOT NULL DEFAULT 1 CHECK (version > 0),
    enrolled_at_utc timestamptz NOT NULL,
    last_seen_at_utc timestamptz,
    revoked_at_utc timestamptz,
    revocation_reason text,
    CHECK (jsonb_typeof(public_key_jwk) = 'object')
);

CREATE INDEX ix_devices_user_status ON devices(user_id, status);
CREATE INDEX ix_devices_organization_status ON devices(organization_id, status) WHERE organization_id IS NOT NULL;

CREATE TABLE device_challenges (
    id uuid PRIMARY KEY,
    user_id uuid NOT NULL REFERENCES users(id),
    nonce_hash bytea NOT NULL UNIQUE,
    expires_at_utc timestamptz NOT NULL,
    consumed_at_utc timestamptz,
    created_at_utc timestamptz NOT NULL,
    CHECK (expires_at_utc > created_at_utc)
);

CREATE TABLE refresh_token_families (
    id uuid PRIMARY KEY,
    user_id uuid NOT NULL REFERENCES users(id),
    device_id uuid NOT NULL REFERENCES devices(id),
    current_token_hash bytea NOT NULL UNIQUE,
    rotated_at_utc timestamptz,
    revoked_at_utc timestamptz,
    reuse_detected_at_utc timestamptz,
    created_at_utc timestamptz NOT NULL
);

CREATE TABLE products (
    id uuid PRIMARY KEY,
    code text NOT NULL UNIQUE,
    display_name text NOT NULL,
    status text NOT NULL CHECK (status IN ('DRAFT', 'ACTIVE', 'RETIRED')),
    version bigint NOT NULL DEFAULT 1 CHECK (version > 0),
    created_at_utc timestamptz NOT NULL
);

CREATE TABLE features (
    id uuid PRIMARY KEY,
    product_id uuid NOT NULL REFERENCES products(id),
    code text NOT NULL,
    display_name text NOT NULL,
    UNIQUE (product_id, code)
);

CREATE TABLE trial_campaigns (
    id uuid PRIMARY KEY,
    product_id uuid NOT NULL REFERENCES products(id),
    kind trial_kind NOT NULL CHECK (kind = 'LAUNCH'),
    status trial_campaign_status NOT NULL,
    starts_at_utc timestamptz NOT NULL,
    ends_at_utc timestamptz NOT NULL,
    published_at_utc timestamptz,
    terminated_at_utc timestamptz,
    version bigint NOT NULL DEFAULT 1 CHECK (version > 0),
    created_at_utc timestamptz NOT NULL,
    CHECK (ends_at_utc > starts_at_utc),
    EXCLUDE USING gist (
        product_id WITH =,
        tstzrange(starts_at_utc, ends_at_utc, '[)') WITH &&
    ) WHERE (status IN ('SCHEDULED', 'ACTIVE'))
);

CREATE TABLE trial_grants (
    id uuid PRIMARY KEY,
    user_id uuid NOT NULL REFERENCES users(id),
    product_id uuid NOT NULL REFERENCES products(id),
    campaign_id uuid REFERENCES trial_campaigns(id),
    kind trial_kind NOT NULL,
    status trial_grant_status NOT NULL,
    starts_at_utc timestamptz NOT NULL,
    ends_at_utc timestamptz NOT NULL,
    source_reference text NOT NULL,
    created_at_utc timestamptz NOT NULL,
    revoked_at_utc timestamptz,
    converted_at_utc timestamptz,
    CHECK (ends_at_utc > starts_at_utc),
    UNIQUE (user_id, product_id)
);

CREATE TABLE entitlement_grants (
    id uuid PRIMARY KEY,
    user_id uuid REFERENCES users(id),
    organization_id uuid REFERENCES organizations(id),
    product_id uuid NOT NULL REFERENCES products(id),
    source entitlement_source NOT NULL,
    status entitlement_status NOT NULL,
    effective_from_utc timestamptz NOT NULL,
    effective_until_utc timestamptz NOT NULL,
    feature_codes text[] NOT NULL,
    source_reference text NOT NULL,
    reason text,
    created_at_utc timestamptz NOT NULL,
    revoked_at_utc timestamptz,
    CHECK ((user_id IS NULL) <> (organization_id IS NULL)),
    CHECK (effective_until_utc > effective_from_utc),
    CHECK (cardinality(feature_codes) > 0)
);

CREATE INDEX ix_entitlement_user_period ON entitlement_grants(user_id, effective_from_utc, effective_until_utc) WHERE user_id IS NOT NULL;
CREATE INDEX ix_entitlement_org_period ON entitlement_grants(organization_id, effective_from_utc, effective_until_utc) WHERE organization_id IS NOT NULL;

CREATE TABLE license_leases (
    id uuid PRIMARY KEY,
    user_id uuid NOT NULL REFERENCES users(id),
    organization_id uuid REFERENCES organizations(id),
    device_id uuid NOT NULL REFERENCES devices(id),
    client_release_id text NOT NULL,
    source entitlement_source NOT NULL,
    feature_codes text[] NOT NULL,
    key_id text NOT NULL,
    jti text NOT NULL UNIQUE,
    issued_at_utc timestamptz NOT NULL,
    not_before_utc timestamptz NOT NULL,
    expires_at_utc timestamptz NOT NULL,
    revoked_at_utc timestamptz,
    CHECK (expires_at_utc > not_before_utc),
    CHECK (not_before_utc >= issued_at_utc)
);

CREATE TABLE execution_grants (
    id uuid PRIMARY KEY,
    user_id uuid NOT NULL REFERENCES users(id),
    organization_id uuid REFERENCES organizations(id),
    device_id uuid NOT NULL REFERENCES devices(id),
    command_id text NOT NULL,
    client_release_id text NOT NULL,
    document_fingerprint text NOT NULL,
    operation_scope text NOT NULL,
    nonce_hash bytea NOT NULL UNIQUE,
    jti text NOT NULL UNIQUE,
    key_id text NOT NULL,
    issued_at_utc timestamptz NOT NULL,
    not_before_utc timestamptz NOT NULL,
    expires_at_utc timestamptz NOT NULL,
    consumed_at_utc timestamptz,
    CHECK (expires_at_utc > not_before_utc),
    CHECK (not_before_utc >= issued_at_utc)
);

CREATE TABLE offers (
    id uuid PRIMARY KEY,
    product_id uuid NOT NULL REFERENCES products(id),
    version integer NOT NULL CHECK (version > 0),
    audience text NOT NULL,
    channel text NOT NULL,
    currency char(3) NOT NULL,
    amount_minor bigint NOT NULL CHECK (amount_minor >= 0),
    billing_term_seconds bigint NOT NULL CHECK (billing_term_seconds > 0),
    feature_codes text[] NOT NULL,
    status offer_status NOT NULL,
    effective_from_utc timestamptz NOT NULL,
    effective_until_utc timestamptz,
    published_at_utc timestamptz,
    retired_at_utc timestamptz,
    created_at_utc timestamptz NOT NULL,
    CHECK (effective_until_utc IS NULL OR effective_until_utc > effective_from_utc),
    UNIQUE (product_id, version),
    EXCLUDE USING gist (
        product_id WITH =,
        audience WITH =,
        channel WITH =,
        currency WITH =,
        tstzrange(effective_from_utc, COALESCE(effective_until_utc, 'infinity'::timestamptz), '[)') WITH &&
    ) WHERE (status = 'PUBLISHED')
);

CREATE TABLE quotes (
    id uuid PRIMARY KEY,
    user_id uuid NOT NULL REFERENCES users(id),
    organization_id uuid REFERENCES organizations(id),
    offer_id uuid NOT NULL REFERENCES offers(id),
    offer_version integer NOT NULL,
    product_id uuid NOT NULL REFERENCES products(id),
    currency char(3) NOT NULL,
    amount_minor bigint NOT NULL CHECK (amount_minor >= 0),
    billing_term_seconds bigint NOT NULL CHECK (billing_term_seconds > 0),
    feature_codes text[] NOT NULL,
    status quote_status NOT NULL,
    integrity_binding bytea NOT NULL,
    created_at_utc timestamptz NOT NULL,
    expires_at_utc timestamptz NOT NULL,
    consumed_at_utc timestamptz,
    CHECK (expires_at_utc > created_at_utc)
);

CREATE TABLE orders (
    id uuid PRIMARY KEY,
    user_id uuid NOT NULL REFERENCES users(id),
    organization_id uuid REFERENCES organizations(id),
    quote_id uuid NOT NULL UNIQUE REFERENCES quotes(id),
    stable_order_code text NOT NULL UNIQUE,
    currency char(3) NOT NULL,
    amount_minor bigint NOT NULL CHECK (amount_minor >= 0),
    status order_status NOT NULL,
    version bigint NOT NULL DEFAULT 1 CHECK (version > 0),
    created_at_utc timestamptz NOT NULL,
    updated_at_utc timestamptz NOT NULL
);

CREATE TABLE payment_events (
    id uuid PRIMARY KEY,
    order_id uuid NOT NULL REFERENCES orders(id),
    provider text NOT NULL,
    provider_event_id text NOT NULL,
    event_type text NOT NULL,
    signature_valid boolean NOT NULL,
    merchant_reference text,
    currency char(3),
    amount_minor bigint,
    payload_hash bytea NOT NULL,
    received_at_utc timestamptz NOT NULL,
    processed_at_utc timestamptz,
    processing_result text,
    UNIQUE (provider, provider_event_id)
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

CREATE TABLE approval_requests (
    id uuid PRIMARY KEY,
    environment text NOT NULL,
    organization_id uuid REFERENCES organizations(id),
    action_code text NOT NULL,
    target_type text NOT NULL,
    target_id text NOT NULL,
    target_version text NOT NULL,
    payload_hash bytea NOT NULL,
    requester_user_id uuid NOT NULL REFERENCES users(id),
    approver_user_id uuid REFERENCES users(id),
    status approval_status NOT NULL,
    reason text NOT NULL,
    impact_summary text NOT NULL,
    created_at_utc timestamptz NOT NULL,
    expires_at_utc timestamptz NOT NULL,
    decided_at_utc timestamptz,
    executed_at_utc timestamptz,
    correlation_id uuid NOT NULL,
    CHECK (expires_at_utc > created_at_utc),
    CHECK (approver_user_id IS NULL OR approver_user_id <> requester_user_id)
);

CREATE TABLE audit_events (
    sequence_id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    event_id uuid NOT NULL UNIQUE,
    occurred_at_utc timestamptz NOT NULL,
    actor_user_id uuid REFERENCES users(id),
    organization_id uuid REFERENCES organizations(id),
    environment text NOT NULL,
    action_code text NOT NULL,
    target_type text NOT NULL,
    target_id text NOT NULL,
    reason text,
    before_hash bytea,
    after_hash bytea,
    correlation_id uuid NOT NULL,
    previous_event_hash bytea,
    event_hash bytea NOT NULL
);

CREATE OR REPLACE FUNCTION reject_immutable_update() RETURNS trigger LANGUAGE plpgsql AS $$
BEGIN
    RAISE EXCEPTION 'immutable table % does not permit update or delete', TG_TABLE_NAME
        USING ERRCODE = '55000';
END;
$$;

CREATE TRIGGER audit_events_immutable
BEFORE UPDATE OR DELETE ON audit_events
FOR EACH ROW EXECUTE FUNCTION reject_immutable_update();

CREATE TRIGGER payment_events_immutable
BEFORE UPDATE OR DELETE ON payment_events
FOR EACH ROW EXECUTE FUNCTION reject_immutable_update();

COMMIT;
