import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.send_summary import EMAIL_SUBJECT_PREFIX, is_stale_pipeline_outbox_item

SUBJECT = f"{EMAIL_SUBJECT_PREFIX} Summary - 2026-09-05"
THRESHOLD = 3600


def test_stale_pipeline_item_is_flagged():
    now = datetime(2026, 9, 7, 12, 0, 0)
    created_at = now - timedelta(hours=2)
    assert is_stale_pipeline_outbox_item(SUBJECT, created_at, now, THRESHOLD) is True


def test_fresh_pipeline_item_is_not_flagged():
    now = datetime(2026, 9, 7, 12, 0, 0)
    created_at = now - timedelta(seconds=5)
    assert is_stale_pipeline_outbox_item(SUBJECT, created_at, now, THRESHOLD) is False


def test_non_pipeline_subject_is_never_flagged_regardless_of_age():
    now = datetime(2026, 9, 7, 12, 0, 0)
    created_at = now - timedelta(days=30)
    assert is_stale_pipeline_outbox_item("Some unrelated email", created_at, now, THRESHOLD) is False


def test_boundary_at_exact_threshold():
    now = datetime(2026, 9, 7, 12, 0, 0)
    created_at = now - timedelta(seconds=THRESHOLD)
    assert is_stale_pipeline_outbox_item(SUBJECT, created_at, now, THRESHOLD) is True


def test_tz_aware_inputs_both_sides_consistent():
    now = datetime(2026, 9, 7, 12, 0, 0, tzinfo=timezone.utc)
    created_at = now - timedelta(hours=2)
    assert is_stale_pipeline_outbox_item(SUBJECT, created_at, now, THRESHOLD) is True


def test_naive_vs_aware_mismatch_raises_typeerror():
    now = datetime(2026, 9, 7, 12, 0, 0)
    created_at = datetime(2026, 9, 7, 10, 0, 0, tzinfo=timezone.utc)
    with pytest.raises(TypeError):
        is_stale_pipeline_outbox_item(SUBJECT, created_at, now, THRESHOLD)
