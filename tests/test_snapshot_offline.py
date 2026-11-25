import csv
import json
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

from config import PROJECT_ROOT
from workflow.snapshot_offline import SnapshotOptions, SnapshotUsageError, run_offline_snapshot


FIXTURES = PROJECT_ROOT / "tests" / "fixtures"
SNAPSHOT = FIXTURES / "snapshot_offline_synthetic.json"
STRONG_IDS = FIXTURES / "snapshot_offline_strong_ids.csv"
BRANDS = FIXTURES / "snapshot_offline_brands.json"
THEME = FIXTURES / "snapshot_offline_theme.txt"
REVIEW = FIXTURES / "snapshot_offline_human_review.csv"
SECRET = "synthetic-contact@example.test"
ASIN = "SYNTHETIC1"


def _options(tmp_path: Path, **overrides) -> SnapshotOptions:
    values = dict(
        input_path=SNAPSHOT,
        output_dir=tmp_path / "out",
        strong_id_paths=[STRONG_IDS],
        brand_exclusion_path=BRANDS,
        theme_keywords_path=THEME,
        human_review_path=REVIEW,
        limit=10,
    )
    values.update(overrides)
    return SnapshotOptions(**values)


def _run(tmp_path: Path, **overrides):
    summary = run_offline_snapshot(_options(tmp_path, **overrides))
    output = tmp_path / "out"
    audit = list(csv.DictReader((output / "audit.csv").open(encoding="utf-8-sig")))
    shortlist = list(csv.DictReader((output / "review-shortlist.csv").open(encoding="utf-8-sig")))
    return summary, audit, shortlist


def _assert_no_contact_leak(directory: Path) -> None:
    for path in directory.rglob("*"):
        if not path.is_file():
            continue
        if path.suffix.lower() == ".xlsx":
            workbook = load_workbook(path, data_only=True)
            values = [
                str(cell)
                for sheet in workbook.worksheets
                for row in sheet.iter_rows(values_only=True)
                for cell in row
                if cell is not None
            ]
            text = "\n".join(values)
        else:
            text = path.read_text(encoding="utf-8", errors="ignore")
        assert SECRET not in text
        assert ASIN not in text
        assert "ready_to_send,yes" not in text.lower()


def test_snapshot_shortlist_reuses_filter_and_keeps_unknowns(tmp_path, monkeypatch):
    def fail_network(*_args, **_kwargs):
        raise AssertionError("snapshot replay called a YouTube client")

    monkeypatch.setattr("youtube.client.api_request", fail_network)
    monkeypatch.setattr("youtube.videos.api_request", fail_network)
    monkeypatch.setattr("youtube.channels.api_request", fail_network)

    summary, audit, shortlist = _run(tmp_path)
    audit_ids = {row["channel_id"] for row in audit}
    shortlist_ids = [row["channel_id"] for row in shortlist]

    assert summary["live_youtube_called"] is False
    assert summary["end_to_end_online_automation_passed"] is False
    assert shortlist_ids == [
        "UCsynth00000000000000002",
        "UCsynth00000000000000001",
        "UCsynth00000000000000016",
    ]
    assert "UCsynth00000000000000003" in audit_ids
    assert "UCsynth00000000000000003" not in shortlist_ids
    assert "UCsynth00000000000000012" in audit_ids
    assert "UCsynth00000000000000012" not in shortlist_ids
    assert "UCsynth00000000000000004" not in audit_ids
    assert "UCsynth00000000000000005" not in audit_ids
    assert "UCsynth00000000000000005" not in shortlist_ids

    assert all(row["ready_to_send"] == "no" for row in audit)
    assert all(row["creator_identity"] == "UNKNOWN" for row in audit)
    assert all(row["identity_verified"] == "no" for row in audit)
    assert all(row["audience_geo_status"] == "UNKNOWN" for row in audit)
    assert all(row["amazon_native_tag"] == "UNKNOWN" for row in audit)
    assert all(row["brand_channel_check"] == "UNKNOWN" for row in audit)
    assert all(row["theme_keyword_is_relevance"] == "no" for row in audit)
    assert all("验证通过" not in row["fit_reason"] for row in audit)

    passed = shortlist[0]
    assert passed["original_video_filter"] == "pass"
    assert passed["threshold_exception"] == ""
    assert passed["human_review_decision"] == "kept"
    assert passed["recommendation"] == "review_shortlist"

    exceptions = [row for row in shortlist if row["threshold_exception"]]
    assert [row["channel_id"] for row in exceptions] == [
        "UCsynth00000000000000001",
        "UCsynth00000000000000016",
    ]
    assert all(row["original_video_filter"] == "fail" for row in exceptions)
    assert all(
        row["threshold_exception"] == "human_retained_exception_auto_threshold_not_passed"
        for row in exceptions
    )

    dropped = next(row for row in audit if row["channel_id"] == "UCsynth00000000000000003")
    assert dropped["human_review_decision"] == "dropped"
    assert dropped["queue"] == "human_dropped_not_recommended"
    assert dropped["recommendation"] == "not_recommended"
    assert "cabinet lock" in dropped["theme_keywords_hit"]

    unnamed = next(row for row in audit if row["channel_id"] == "UCsynth00000000000000012")
    assert unnamed["human_review_decision"] == "not_listed"
    assert unnamed["queue"] == "pending_identity_check_not_recommended"
    assert unnamed["recommendation"] == "not_recommended"
    assert "品牌官方" in unnamed["taxonomy_preliminary"]
    assert unnamed["brand_channel_check"] == "UNKNOWN"

    counts = summary["counts"]
    assert counts["original_filter_qualified_videos"] == 11
    assert counts["original_filter_brand_excluded_videos"] == 1
    assert counts["deduped_existing_strong_id_channels"] == 1
    assert counts["audit_rows"] == 8
    assert counts["review_shortlist_rows"] == 3
    assert counts["threshold_exceptions_in_shortlist"] == 2
    assert counts["human_review_kept_labels"] == 4
    assert counts["human_dropped_in_audit"] == 1
    assert counts["pending_identity_check_not_recommended"] == 4
    assert counts["ready_to_send_rows"] == 0
    assert counts["theme_keyword_is_relevance"] == 0
    assert counts["subscriber_hidden"] == 1
    assert counts["subscriber_missing"] == 1
    assert counts["subscriber_known_zero"] == 1

    hidden = next(item for item in summary["subscriber_unknown_not_zero"] if item["status"] == "hidden")
    assert hidden["exported_value"] == "UNKNOWN"
    assert hidden["parser_would_write"] == "99999"
    assert "99999" not in {row["subscriber_count"] for row in audit}

    excluded = {row["channel_id"]: row["reason"] for row in summary["in_band_excluded"]}
    assert excluded["UCsynth00000000000000004"] == "brand_exclusion"
    assert excluded["UCsynth00000000000000005"].startswith("strong_channel_id:")
    assert excluded["UCsynth00000000000000013"] == "theme_not_matched"
    _assert_no_contact_leak(tmp_path / "out")

    workbook = load_workbook(tmp_path / "out" / "offline_snapshot" / "search_videos.xlsx")
    sheet = workbook.active
    headers = [cell.value for cell in sheet[1]]
    likes = headers.index("Likes")
    video_id = headers.index("Video ID")
    for record in sheet.iter_rows(min_row=2, values_only=True):
        if record[video_id] == "synthvid012":
            assert record[likes] in (None, "")
            break
    else:
        raise AssertionError("missing synthetic video row")


def test_human_review_gates_are_not_satisfied_by_the_old_suite(tmp_path):
    """Business acceptance: keyword and subscriber hits do not make a recommendation."""
    _summary, audit, shortlist = _run(tmp_path)
    dropped = "UCsynth00000000000000003"
    not_reviewed = "UCsynth00000000000000012"
    strong_kept = "UCsynth00000000000000005"
    shortlist_ids = {row["channel_id"] for row in shortlist}

    assert dropped not in shortlist_ids
    assert next(row["queue"] for row in audit if row["channel_id"] == dropped) == "human_dropped_not_recommended"
    assert not_reviewed not in shortlist_ids
    assert next(row["queue"] for row in audit if row["channel_id"] == not_reviewed) == (
        "pending_identity_check_not_recommended"
    )
    assert strong_kept not in shortlist_ids
    assert all(row["ready_to_send"] == "no" for row in shortlist)
    assert all(row["identity_verified"] == "no" for row in shortlist)
    exceptions = [
        row for row in shortlist
        if row["threshold_exception"] == "human_retained_exception_auto_threshold_not_passed"
    ]
    assert len(shortlist) == 3
    assert len(exceptions) == 2
    assert all(row["original_video_filter"] == "fail" for row in exceptions)


def test_limit_does_not_pad_and_repo_output_is_refused(tmp_path):
    summary, audit, shortlist = _run(tmp_path, limit=1)
    assert [row["channel_id"] for row in shortlist] == ["UCsynth00000000000000002"]
    assert len(audit) == 8
    assert summary["counts"]["review_shortlist_before_cap"] == 3
    assert summary["counts"]["shortlist_truncated"] == 2

    with pytest.raises(SnapshotUsageError):
        run_offline_snapshot(_options(tmp_path, output_dir=PROJECT_ROOT / "output"))


def test_snapshot_cli_stays_offline(tmp_path):
    output = tmp_path / "cli-out"
    completed = subprocess.run(
        [
            sys.executable,
            "main.py",
            "snapshot-input",
            "--input",
            str(SNAPSHOT),
            "--strong-ids",
            str(STRONG_IDS),
            "--brand-exclusion",
            str(BRANDS),
            "--theme-keywords-file",
            str(THEME),
            "--human-review",
            str(REVIEW),
            "--output-dir",
            str(output),
            "--limit",
            "10",
        ],
        cwd=PROJECT_ROOT,
        text=True,
        capture_output=True,
        check=False,
    )
    assert completed.returncode == 0, completed.stderr
    assert "不是端到端在线自动化通过" in completed.stdout
    payload = json.loads((output / "summary.json").read_text(encoding="utf-8"))
    assert payload["live_youtube_called"] is False
    assert (output / "report.md").exists()
    assert (output / "audit.csv").exists()
    assert (output / "review-shortlist.csv").exists()
    assert payload["counts"]["ready_to_send_rows"] == 0

    rejected = subprocess.run(
        [
            sys.executable,
            "main.py",
            "snapshot-input",
            "--input",
            str(SNAPSHOT),
            "--output-dir",
            str(output),
            "--limit",
            "11",
        ],
        cwd=PROJECT_ROOT,
        text=True,
        capture_output=True,
        check=False,
    )
    assert rejected.returncode == 2
