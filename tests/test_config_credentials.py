"""Google credentials from env vars, in all the shapes hosting panels tend to produce.
Uses a throwaway RSA key generated on the fly (no real secrets)."""
import json

import pytest
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from google.oauth2 import service_account

from mazboot_bot import config

KEY = rsa.generate_private_key(public_exponent=65537, key_size=2048).private_bytes(
    serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption()).decode()
EMAIL = "bot@example-project.iam.gserviceaccount.com"
INFO = {"type": "service_account", "project_id": "example-project", "private_key_id": "abc123",
        "private_key": KEY, "client_email": EMAIL, "client_id": "1",
        "token_uri": "https://oauth2.googleapis.com/token"}


@pytest.fixture(autouse=True)
def clean_env(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)   # no .env file here
    for name in list(__import__("os").environ):
        if any(h in name.upper() for h in ("GCP", "GOOGLE", "PRIVATE", "CLIENT")):
            monkeypatch.delenv(name)


def _assert_usable(cfg):
    info = json.loads(cfg.google_credentials_json)
    assert info["private_key"].strip() == KEY.strip()
    assert info["client_email"] == EMAIL
    service_account.Credentials.from_service_account_info(info)   # google-auth accepts it


def test_whole_json(monkeypatch):
    monkeypatch.setenv("GOOGLE_CREDENTIALS_JSON", json.dumps(INFO))
    cfg = config.load()
    _assert_usable(cfg)
    assert cfg.google_credentials_source == "GOOGLE_CREDENTIALS_JSON"


def test_whole_json_with_real_newlines_in_key(monkeypatch):
    raw = json.dumps(INFO).replace("\\n", "\n")   # panel turned \n into real line breaks
    monkeypatch.setenv("GOOGLE_CREDENTIALS_JSON", raw)
    _assert_usable(config.load())


def test_invalid_json_gives_clear_error(monkeypatch):
    monkeypatch.setenv("GOOGLE_CREDENTIALS_JSON", '{"type": "service_account", "private_key": "trunc')
    with pytest.raises(SystemExit, match="not valid JSON"):
        config.load()


@pytest.mark.parametrize("key_value", [
    KEY.replace("\n", "\\n"),                 # literal \n (as in the JSON file)
    KEY,                                      # real newlines
    KEY.replace("\n", " "),                   # newlines flattened to spaces by a one-line input
    '"' + KEY.replace("\n", "\\n") + '",',    # copied with the JSON quotes and trailing comma
])
def test_separate_gcp_vars(monkeypatch, key_value):
    monkeypatch.setenv("GCP_PRIVATE_KEY", key_value)
    monkeypatch.setenv("GCP_CLIENT_EMAIL", EMAIL)
    monkeypatch.setenv("GCP_PROJECT_ID", "example-project")
    cfg = config.load()
    _assert_usable(cfg)
    assert cfg.google_credentials_source.startswith("separate variables")


@pytest.mark.parametrize("prefix", ["GOOGLE_", "", "lower"])
def test_other_names(monkeypatch, prefix):
    if prefix == "lower":
        monkeypatch.setenv("private_key", KEY.replace("\n", "\\n"))
        monkeypatch.setenv("client_email", EMAIL)
    else:
        monkeypatch.setenv(prefix + "PRIVATE_KEY", KEY.replace("\n", "\\n"))
        monkeypatch.setenv(prefix + "CLIENT_EMAIL", EMAIL)
    _assert_usable(config.load())


def test_key_split_in_parts(monkeypatch):
    escaped = KEY.replace("\n", "\\n")
    third = len(escaped) // 3
    for i, part in enumerate([escaped[:third], escaped[third:2 * third], escaped[2 * third:]], 1):
        monkeypatch.setenv(f"GCP_PRIVATE_KEY_{i}", part)
    monkeypatch.setenv("GCP_CLIENT_EMAIL", EMAIL)
    _assert_usable(config.load())


def test_key_without_email_gives_clear_error(monkeypatch):
    monkeypatch.setenv("GCP_PRIVATE_KEY", KEY)
    with pytest.raises(SystemExit, match="client email"):
        config.load()


def test_nothing_set_falls_back_to_file():
    cfg = config.load()
    assert cfg.google_credentials_json == "" and cfg.google_credentials == "service_account.json"
