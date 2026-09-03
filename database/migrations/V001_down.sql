BEGIN;

DROP TABLE IF EXISTS audit_events;
DROP TABLE IF EXISTS approval_requests;
DROP TABLE IF EXISTS idempotency_records;
DROP TABLE IF EXISTS outbox_messages;
DROP TABLE IF EXISTS payment_events;
DROP TABLE IF EXISTS orders;
DROP TABLE IF EXISTS quotes;
DROP TABLE IF EXISTS offers;
DROP TABLE IF EXISTS execution_grants;
DROP TABLE IF EXISTS license_leases;
DROP TABLE IF EXISTS entitlement_grants;
DROP TABLE IF EXISTS trial_grants;
DROP TABLE IF EXISTS trial_campaigns;
DROP TABLE IF EXISTS features;
DROP TABLE IF EXISTS products;
DROP TABLE IF EXISTS refresh_token_families;
DROP TABLE IF EXISTS device_challenges;
DROP TABLE IF EXISTS devices;
DROP TABLE IF EXISTS organization_memberships;
DROP TABLE IF EXISTS identities;
DROP TABLE IF EXISTS users;
DROP TABLE IF EXISTS organizations;

DROP FUNCTION IF EXISTS reject_immutable_update();

DROP TYPE IF EXISTS approval_status;
DROP TYPE IF EXISTS order_status;
DROP TYPE IF EXISTS quote_status;
DROP TYPE IF EXISTS offer_status;
DROP TYPE IF EXISTS entitlement_status;
DROP TYPE IF EXISTS entitlement_source;
DROP TYPE IF EXISTS trial_grant_status;
DROP TYPE IF EXISTS trial_campaign_status;
DROP TYPE IF EXISTS trial_kind;
DROP TYPE IF EXISTS device_assurance;
DROP TYPE IF EXISTS device_status;
DROP TYPE IF EXISTS account_status;

COMMIT;
