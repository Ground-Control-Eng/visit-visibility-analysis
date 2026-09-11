import sys
from pathlib import Path

import pytest
import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import config


def _base_config_dict():
    return {
        "sql": {
            "ice2": {"server": "s", "database": "d"},
            "visits_api_excel": {"path": "API Status.xlsx"},
        },
        "statuses": {},
        "email": {
            "trigger": {"sender_filter": "ci@hubscape.co.uk", "subject_contains": ""},
            "recipients": {"to": ["a@b.com"]},
        },
        "hubscape_columns": {"api_id_column": "External Visit API Id"},
        "run": {},
        "graph": {
            "tenant_id": "tenant-1",
            "client_id": "client-1",
            "mailbox": "alex.clark@ground-control.co.uk",
        },
    }


def _write_config(tmp_path, graph_overrides=None) -> Path:
    raw = _base_config_dict()
    if graph_overrides:
        raw["graph"].update(graph_overrides)
    path = tmp_path / "config.yaml"
    path.write_text(yaml.dump(raw), encoding="utf-8")
    return path


@pytest.fixture(autouse=True)
def _isolate_from_real_dotenv(monkeypatch, tmp_path):
    # Point .env loading at a path that doesn't exist, so these tests are never affected by a
    # real .env file (with a real secret) sitting in the actual project root.
    monkeypatch.setattr(config, "DEFAULT_ENV_PATH", tmp_path / "does-not-exist.env")
    monkeypatch.delenv("GRAPH_CLIENT_SECRET", raising=False)


def test_load_config_succeeds_without_graph_client_secret_env_var(tmp_path):
    path = _write_config(tmp_path)
    cfg = config.load_config(path)
    assert cfg.graph.client_secret == ""
    assert cfg.graph.tenant_id == "tenant-1"


def test_load_config_rejects_negative_max_retries(tmp_path):
    path = _write_config(tmp_path, {"max_retries": -1})
    with pytest.raises(config.ConfigError, match="max_retries"):
        config.load_config(path)


def test_load_config_rejects_non_positive_mail_search_page_size(tmp_path):
    path = _write_config(tmp_path, {"mail_search_page_size": 0})
    with pytest.raises(config.ConfigError, match="mail_search_page_size"):
        config.load_config(path)


def test_load_config_rejects_non_positive_mail_search_max_pages(tmp_path):
    path = _write_config(tmp_path, {"mail_search_max_pages": 0})
    with pytest.raises(config.ConfigError, match="mail_search_max_pages"):
        config.load_config(path)


def test_load_config_rejects_non_positive_max_inline_attachment_bytes(tmp_path):
    path = _write_config(tmp_path, {"max_inline_attachment_bytes": 0})
    with pytest.raises(config.ConfigError, match="max_inline_attachment_bytes"):
        config.load_config(path)


def test_load_config_rejects_non_positive_request_timeout_seconds(tmp_path):
    path = _write_config(tmp_path, {"request_timeout_seconds": 0})
    with pytest.raises(config.ConfigError, match="request_timeout_seconds"):
        config.load_config(path)
