"""Constants for the Google Outlook Contacts Sync integration."""

DOMAIN = "google_outlook_contacts_sync"

# Config entry keys
CONF_MS_CLIENT_ID = "ms_client_id"
CONF_MS_CLIENT_SECRET = "ms_client_secret"
CONF_GOOGLE_CLIENT_ID = "google_client_id"
CONF_GOOGLE_CLIENT_SECRET = "google_client_secret"
CONF_SYNC_INTERVAL_HOURS = "sync_interval_hours"
CONF_DELETE_REMOVED = "delete_removed"
CONF_AUTO_REMOVE_DUPLICATES = "auto_remove_duplicates"

# Defaults
DEFAULT_SYNC_INTERVAL_HOURS = 24
DEFAULT_DELETE_REMOVED = False
# Destructive: when on, duplicate Outlook contacts (matched by name) are deleted
# automatically on every sync, keeping one per group. Off by default.
DEFAULT_AUTO_REMOVE_DUPLICATES = False

# Birthday placeholder year for contacts without a year (Outlook convention).
BIRTHDAY_PLACEHOLDER_YEAR = 1604

# Microsoft Graph / O365 OAuth scopes
# MSAL (used internally by O365) reserves and adds offline_access, openid, and
# profile automatically — passing them explicitly raises a ValueError.
MS_SCOPES = [
    "Contacts.ReadWrite",
    "User.Read",
]

# Google People API scope
GOOGLE_SCOPE = "https://www.googleapis.com/auth/contacts.readonly"
GOOGLE_API_SERVICE = "people"
GOOGLE_API_VERSION = "v1"

# HA HTTP view paths for OAuth callbacks
MS_AUTH_CALLBACK_PATH = "/api/google_outlook_contacts_sync/auth"
MS_AUTH_CALLBACK_NAME = "api:google_outlook_contacts_sync:auth"
GOOGLE_AUTH_CALLBACK_PATH = "/api/google_outlook_contacts_sync/google_auth"
GOOGLE_AUTH_CALLBACK_NAME = "api:google_outlook_contacts_sync:google_auth"

# Storage keys
STORAGE_KEY_MAPPING = f"{DOMAIN}.mapping"
STORAGE_VERSION = 1

# Services
SERVICE_SYNC_NOW = "sync_now"
SERVICE_FULL_SYNC = "full_sync"

# hass.data sub-keys
DATA_COORDINATOR = "coordinator"
DATA_DRY_RUN_COORDINATOR = "dry_run_coordinator"
DATA_STORE = "store"
DATA_MS_OAUTH_FLOWS = "ms_oauth_flows"
DATA_GOOGLE_OAUTH_FLOWS = "google_oauth_flows"
DATA_REAUTH_PROVIDER = "reauth_provider"

# Device metadata (groups all entities under one HA device)
DEVICE_NAME = "Google Outlook Contacts Sync"
DEVICE_MANUFACTURER = "N1k4G"
DEVICE_MODEL = "Contacts Sync"
