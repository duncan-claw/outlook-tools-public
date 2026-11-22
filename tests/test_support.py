"""Shared synthetic, non-production configuration for offline tests."""
from outlook import config
from spam_guard import sender_identity

config.ACCOUNT = "owner@example.com"
config.GRAPH_CLIENT_ID = "test-app-id"
config.HISTORY_CUTOFF = "2000-01-17T00:00:00+00:00"
sender_identity.FREEMAIL = frozenset({"example.com"})
