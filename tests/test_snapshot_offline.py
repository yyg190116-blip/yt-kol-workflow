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
SECRET = "synthetic-contact@example.test"
ASIN = "SYNTHETIC1"


def _options(tmp_path: Path, **overrides) -> SnapshotOptions:
    values = dict(
        input_path=SNAPSHOT,
        output_dir=tmp_path / "out",
        strong_id_paths=[STRONG_IDS],
        brand_exclusion_path=BRANDS,
        theme_keywords_path=THEME,
        limit=10,
    )
    values.update(overrides)
    return SnapshotOptions(**values)


def _run(tmp_path: Path, **overrides):
    summary = run_offline_snapshot(_options(tmp_path, **overrides))
    rows = list(csv.DictReader((tmp_path / "out" / "candidates.csv").open(encoding="utf-8-sig")))
    return summary, rows


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

    summary, rows = _run(tmp_path)
    ids = [row["channel_id"] for row in rows]

    assert summary["live_youtube_called"] is False
    assert summary["end_to_end_online_automation_passed"] is False
    assert ids == [
        "UCsynth00000000000000002",
        "UCsynth00000000000000012",
        "UCsynth00000000000000001",
        "UCsynth00000000000000014",
        "UCsynth00000000000000016",
        "UCsynth00000000000000017",
        "UCsynth00000000000000018",
        "UCsynth00000000000000003",
    ]
    assert all(row["ready_to_send"] == "no" for row in rows)
    assert all(row["audience_geo_status"] == "UNKNOWN" for row in rows)
    assert all(row["amazon_native_tag"] == "UNKNOWN" for row in rows)
    assert all(row["contact"] == "UNKNOWN" for row in rows)
    assert all(row["email_status"] == "UNKNOWN" for row in rows)
    assert all(row["contact_history"] == "UNKNOWN" for row in rows)
    assert all(row["brand_channel_check"] == "UNKNOWN" for row in rows)
    assert all(row["nox_status"] == "UNKNOWN" for row in rows)
    assert all(row["strong_id_match"] == "no" for row in rows)

    primary = rows[2]
    assert primary["creator_country"] == "US"
    assert primary["amazon_link_seen"] == "yes"
    assert primary["email_seen_in_public_text"] == "yes"
    assert primary["original_video_filter"] == "fail"
    assert primary["snapshot_review"] == "kept"
    assert primary["subscriber_band"] == "10K-50K"
    assert primary["subscriber_count"] == "20000"

    passed = rows[0]
    assert passed["video_id"] == "synthvid002"
    assert passed["original_video_filter"] == "pass"
    assert int(passed["video_views"]) > 10000

    weak = rows[1]
    assert "品牌官方" in weak["taxonomy_preliminary"]
    assert weak["brand_channel_check"] == "UNKNOWN"
    assert weak["video_engagement_rate"] == "UNKNOWN"
    assert weak["creator_country"] == "CA"

    extended = rows[-1]
    assert extended["subscriber_band"] == "50K-100K_extended"
    assert extended["snapshot_review"] == "dropped"
    assert extended["duration_note"] == "under_60s"
    assert extended["duration_seconds"] == "24"
    assert extended["original_video_filter"] == "fail"

    counts = summary["counts"]
    assert counts["original_filter_qualified_videos"] == 11
    assert counts["original_filter_brand_excluded_videos"] == 1
    assert counts["deduped_existing_strong_id_channels"] == 1
    assert counts["strong_id_index_size"] == 1
    assert counts["band_primary_channels"] == 8
    assert counts["band_extended_channels"] == 3
    assert counts["subscriber_hidden"] == 1
    assert counts["subscriber_missing"] == 1
    assert counts["subscriber_known_zero"] == 1
    assert counts["shortlist_written"] == 8
    assert counts["shortlist_truncated"] == 0

    hidden = next(item for item in summary["subscriber_unknown_not_zero"] if item["status"] == "hidden")
    assert hidden["exported_value"] == "UNKNOWN"
    assert hidden["parser_would_write"] == "99999"
    assert "99999" not in {row["subscriber_count"] for row in rows}
    assert "0" not in {row["subscriber_count"] for row in rows}

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


def test_limit_does_not_pad_and_repo_output_is_refused(tmp_path):
    summary, rows = _run(tmp_path, limit=1)
    assert [row["channel_id"] for row in rows] == ["UCsynth00000000000000002"]
    assert summary["counts"]["shortlist_before_cap"] == 8
    assert summary["counts"]["shortlist_truncated"] == 7

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
