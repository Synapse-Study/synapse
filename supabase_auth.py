"""
supabase_auth.py
================
Re-exports Synapse authentication routines from synapse_auth for full backwards compatibility.
"""
from synapse_auth import (
    SUPABASE_URL,
    SUPABASE_PUBLISHABLE_KEY,
    QUESTIONS_APP_URL,
    SECRET_KEYS,
    get_config_path,
    load_secure_config,
    save_secure_config,
    sign_in_synapse,
    sign_up_synapse,
    clear_session_from_config,
    sign_out_synapse,
    get_valid_session,
    accept_question_session,
)

__all__ = [
    "SUPABASE_URL",
    "SUPABASE_PUBLISHABLE_KEY",
    "QUESTIONS_APP_URL",
    "SECRET_KEYS",
    "get_config_path",
    "load_secure_config",
    "save_secure_config",
    "sign_in_synapse",
    "sign_up_synapse",
    "clear_session_from_config",
    "sign_out_synapse",
    "get_valid_session",
    "accept_question_session",
]
