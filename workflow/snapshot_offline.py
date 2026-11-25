#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Offline adapter for an already collected YouTube Data API snapshot.

The live workflow searches and enriches through YouTube, then optionally
writes Feishu. This command does none of that. It reads a local snapshot,
reuses the existing video parser, video filter, channel dedupe, channel
taxonomy and Excel exporter, and writes a capped offline shortlist.

It does not send mail, write Feishu or a CRM, call 17TRACK, start OAuth, or
read an API key. ``ready_to_send`` is fixed to ``no``.
"""

from __future__ import annotations

import csv
import html
import json
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Sequence, Tuple

from config import PROJECT_ROOT, BrandExclusion
from export.excel_exporter import ExcelExporter
from filter.channel_classifier import classify_channel
from filter.channel_dedup import deduplicate_channels
from filter.video_filter import filter_videos
from youtube.channels import _parse_channel_item
from youtube.videos import _parse_video_item

EMAIL_PATTERN = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
AMAZON_PATTERN = re.compile(
    r"(?i)(?:\bamzn\.to\b|\bamazon\.(?:com|co\.uk|ca|de|fr|it|es|in|com\.au|com\.mx|co\.jp)\b)"
)
STRONG_KEY_TYPES = {"channel_id", "channel_id_lower"}
NON_NUMERIC = {"unknown", "hidden", "none", "null", "na", "n/a"}

CANDIDATE_COLUMNS = [
    "channel_id",
    "channel_name",
    "channel_url",
    "creator_country",
    "audience_geo_status",
    "contact",
    "email_status",
    "email_seen_in_public_text",
    "subscriber_count",
    "subscriber_count_status",
    "subscriber_band",
    "hidden_subscribers",
    "brand_channel_check",
    "taxonomy_preliminary",
    "video_id",
    "video_url",
    "video_title",
    "video_published_at",
    "video_views",
    "video_view_status",
    "video_engagement_rate",
    "duration_seconds",
    "duration_hms",
    "duration_note",
    "search_query",
    "theme_keywords_hit",
    "original_video_filter",
    "original_filter_reason",
    "any_video_passed_original_filter",
    "fit_reason",
    "amazon_link_seen",
    "amazon_native_tag",
    "strong_id_match",
    "contact_history",
    "unsubscribe_status",
    "bounce_status",
    "brand_history",
    "nox_status",
    "snapshot_review",
    "snapshot_review_note",
    "human_review_decision",
    "human_review_source",
    "human_review_note",
    "queue",
    "recommendation",
    "creator_identity",
    "identity_verified",
    "threshold_exception",
    "theme_keyword_is_relevance",
    "ready_to_send",
    "evidence_type",
    "trial_bucket",
]

REUSED = [
    "youtube/videos.py:_parse_video_item",
    "youtube/channels.py:_parse_channel_item",
    "filter/video_filter.py:filter_videos",
    "filter/channel_dedup.py:deduplicate_channels",
    "filter/channel_classifier.py:classify_channel",
    "export/excel_exporter.py:ExcelExporter",
    "config.py:BrandExclusion",
]

NOT_TESTED = [
    "YouTube search.list / videos.list / channels.list live calls and pagination",
    "playlist recent-video metrics and activity refresh",
    "Feishu OAuth, setup, bitable writes and notifications",
    "17TRACK logistics and reminder dispatch",
    "email extraction as a usable contact, outreach generation and sending",
    "external LLM, crawler, CRM and paid APIs",
]


class SnapshotUsageError(ValueError):
    """Raised for a local input or output problem. No network is involved."""


@dataclass
class SnapshotOptions:
    input_path: Path
    output_dir: Path
    strong_id_paths: List[Path] = field(default_factory=list)
    brand_exclusion_path: Optional[Path] = None
    theme_keywords_path: Optional[Path] = None
    human_review_path: Optional[Path] = None
    limit: int = 10
    min_subscribers: int = 10000
    primary_max: int = 50000
    extended_max: int = 100000
    min_views: int = 10000
    min_engagement: float = 3.0
    filter_mode: str = "or"
    write_excel: bool = True
    repo_root: Path = PROJECT_ROOT


@dataclass
class StrongIdIndex:
    hits: Dict[str, List[Dict[str, str]]] = field(default_factory=dict)

    def contains(self, channel_id: str) -> bool:
        return _fold(channel_id) in self.hits

    def records_for(self, channel_id: str) -> List[Dict[str, str]]:
        return list(self.hits.get(_fold(channel_id), []))

    def __len__(self) -> int:
        return len(self.hits)


def handle_snapshot_input(args: Any) -> int:
    """CLI entry. Returns a process status and does not contact a network."""
    try:
        options = options_from_args(args)
        guard_output_dir(options.output_dir, options.repo_root)
        block_network()
        summary = run_offline_snapshot(options)
    except SnapshotUsageError as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 2
    counts = summary["counts"]
    print("离线 snapshot 试跑完成。这不是端到端在线自动化通过。")
    print(
        "搜索行 {search_rows}，视频 {unique_videos}，频道 {unique_channels}，"
        "原视频筛选通过 {original_filter_qualified_videos}，"
        "audit {audit_rows}，复核短名单 {review_shortlist_rows}，"
        "ready_to_send {ready_to_send_rows}。".format(**counts)
    )
    print(f"audit CSV: {summary['outputs']['audit_csv']}")
    print(f"复核短名单 CSV: {summary['outputs']['review_shortlist_csv']}")
    print(f"结构 JSON: {summary['outputs']['summary_json']}")
    print(f"报告: {summary['outputs']['report_md']}")
    return 0


def options_from_args(args: Any) -> SnapshotOptions:
    limit = int(getattr(args, "limit", 10))
    min_subscribers = int(getattr(args, "min_subscribers", 10000))
    primary_max = int(getattr(args, "primary_max", 50000))
    extended_max = int(getattr(args, "extended_max", 100000))
    if not 1 <= limit <= 10:
        raise SnapshotUsageError("limit 必须在 1 到 10 之间")
    if min_subscribers < 0 or primary_max < min_subscribers or extended_max < primary_max:
        raise SnapshotUsageError("订阅区间必须满足 0 <= min <= primary_max <= extended_max")
    brand = str(getattr(args, "brand_exclusion", "") or "").strip()
    theme = str(getattr(args, "theme_keywords_file", "") or "").strip()
    review = str(getattr(args, "human_review", "") or "").strip()
    return SnapshotOptions(
        input_path=Path(args.input),
        output_dir=Path(args.output_dir),
        strong_id_paths=[Path(path) for path in (getattr(args, "strong_ids", None) or [])],
        brand_exclusion_path=Path(brand) if brand else None,
        theme_keywords_path=Path(theme) if theme else None,
        human_review_path=Path(review) if review else None,
        limit=limit,
        min_subscribers=min_subscribers,
        primary_max=primary_max,
        extended_max=extended_max,
        min_views=int(getattr(args, "min_views", 10000)),
        min_engagement=float(getattr(args, "min_engagement", 3.0)),
        filter_mode=str(getattr(args, "filter_mode", "or") or "or"),
        write_excel=not bool(getattr(args, "no_excel", False)),
    )


def guard_output_dir(output_dir: Path, repo_root: Path) -> None:
    """Keep real trial outputs out of the git work tree."""
    resolved = output_dir.expanduser().resolve()
    root = repo_root.expanduser().resolve()
    if resolved == root or _is_relative_to(resolved, root):
        raise SnapshotUsageError(
            f"输出目录必须在仓库之外，避免真实候选进入公开提交: {resolved}"
        )


def block_network() -> None:
    """Fail closed if a reused module tries to open HTTP during this command."""
    import requests

    def _blocked(*_args: Any, **_kwargs: Any) -> None:
        raise RuntimeError("snapshot-input 禁止网络请求")

    requests.sessions.Session.request = _blocked  # type: ignore[method-assign]


def run_offline_snapshot(options: SnapshotOptions) -> Dict[str, Any]:
    guard_output_dir(options.output_dir, options.repo_root)
    snapshot = _load_json(options.input_path)
    if not isinstance(snapshot, dict):
        raise SnapshotUsageError("snapshot 根节点必须是 JSON 对象")

    strong_ids = load_strong_channel_ids(options.strong_id_paths)
    brand = _load_brand_exclusion(options.brand_exclusion_path)
    keywords = _load_keywords(options.theme_keywords_path)
    human_review = _load_human_review(options.human_review_path)
    primary_label, extended_label = _band_labels(options)

    search_rows, search_shape = _normalize_search(snapshot)
    videos, video_shape = _normalize_videos(snapshot)
    channels, channel_shape = _normalize_channels(snapshot)
    _attach_search_context(videos, search_rows)

    qualified, _ = filter_videos(
        videos,
        brand,
        min_views=options.min_views,
        min_engagement=options.min_engagement,
        filter_mode=options.filter_mode,
    )
    exact_strong_hits = {
        video.get("channel_id", "")
        for video in qualified
        if video.get("channel_id") and strong_ids.contains(video["channel_id"])
    }
    new_channels, existing_channels = deduplicate_channels(
        qualified,
        exact_strong_hits,
        source_keyword="",
    )

    review = _review_index(snapshot.get("description_review"))
    shortlist, excluded_in_band, fidelity = _build_shortlist(
        videos=videos,
        channels=channels,
        brand=brand,
        strong_ids=strong_ids,
        keywords=keywords,
        review=review,
        human_review=human_review,
        options=options,
        primary_label=primary_label,
        extended_label=extended_label,
    )
    audit_rows = list(shortlist)
    review_rows = [row for row in audit_rows if row.get("human_review_decision") == "kept"]
    written_review = review_rows[: options.limit]
    for row in audit_rows:
        row["strong_id_match"] = "no"
        row["ready_to_send"] = "no"
        row["creator_identity"] = "UNKNOWN"
        row["identity_verified"] = "no"

    output_dir = options.output_dir.expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    audit_path = output_dir / "audit.csv"
    shortlist_path = output_dir / "review-shortlist.csv"
    summary_path = output_dir / "summary.json"
    report_path = output_dir / "report.md"
    excel_dir = ""
    if options.write_excel:
        excel_dir = _export_excel(output_dir, videos, written_review)

    strong_id_hits = _snapshot_strong_hits(
        channels,
        strong_ids,
        options,
        primary_label,
        extended_label,
    )
    public_audit = [
        {column: row.get(column, "") for column in CANDIDATE_COLUMNS}
        for row in audit_rows
    ]
    public_review = [
        {column: row.get(column, "") for column in CANDIDATE_COLUMNS}
        for row in written_review
    ]
    counts = _counts(
        search_rows=search_rows,
        videos=videos,
        channels=channels,
        qualified=qualified,
        new_channels=new_channels,
        existing_channels=existing_channels,
        shortlist=audit_rows,
        written=public_review,
        strong_ids=strong_ids,
        fidelity=fidelity,
        brand=brand,
        keywords=keywords,
        options=options,
        primary_label=primary_label,
        extended_label=extended_label,
    )
    summary = {
        "schema_version": 1,
        "kind": "snapshot_offline_trial",
        "mode": "offline_snapshot",
        "live_youtube_called": False,
        "end_to_end_online_automation_passed": False,
        "ready_to_send": "no",
        "input_shapes": {
            "search": search_shape,
            "videos": video_shape,
            "channels": channel_shape,
        },
        "claimed_counts": snapshot.get("counts") if isinstance(snapshot.get("counts"), dict) else {},
        "search_preferences_are_not_audience": _search_preferences(snapshot),
        "counts": _acceptance_counts(counts, public_audit, public_review, human_review, keywords),
        "subscriber_unknown_not_zero": fidelity,
        "strong_id_hits": strong_id_hits,
        "in_band_excluded": excluded_in_band,
        "candidate_columns": list(CANDIDATE_COLUMNS),
        "audit": public_audit,
        "review_shortlist": public_review,
        "reused": list(REUSED),
        "adapter": [
            "workflow/snapshot_offline.py",
            "main.py snapshot-input",
            "flat or raw snapshot normalization",
            "hidden or missing subscriberCount stays UNKNOWN and is not written as 0",
            "strong channel-id exclusion wins over a human kept label",
            "human kept is the only review shortlist; dropped and not_listed stay in audit",
            "theme keyword hits are not relevance and do not auto-recommend",
            "audience, native tag, contact history and ready_to_send stay unverified",
        ],
        "not_tested": list(NOT_TESTED),
        "boundaries": _boundaries(),
        "outputs": {
            "audit_csv": str(audit_path),
            "review_shortlist_csv": str(shortlist_path),
            "summary_json": str(summary_path),
            "report_md": str(report_path),
            "excel_dir": excel_dir,
        },
        "command": _format_command(options),
    }
    summary = _scrub(summary)
    _write_csv(audit_path, summary["audit"])
    _write_csv(shortlist_path, summary["review_shortlist"])
    _write_json(summary_path, summary)
    report_path.write_text(render_report(summary), encoding="utf-8")
    return summary


def _load_human_review(path: Optional[Path]) -> Dict[str, Any]:
    """Load explicit kept/dropped labels. Any other decision is not a pass."""
    if path is None:
        return {"provided": False, "by_id": {}}
    if not path.is_file():
        raise SnapshotUsageError(f"找不到人工复核文件: {path}")
    by_id: Dict[str, Dict[str, str]] = {}
    with path.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        fields = set(reader.fieldnames or [])
        if not {"channel_id", "decision"} <= fields:
            raise SnapshotUsageError("人工复核 CSV 需要 channel_id 和 decision 列")
        for row in reader:
            channel_id = str(row.get("channel_id") or "").strip()
            decision = str(row.get("decision") or "").strip().lower()
            if not channel_id or decision not in {"kept", "dropped"}:
                continue
            by_id[_fold(channel_id)] = {
                "decision": decision,
                "reason": str(row.get("reason") or "").strip(),
                "source": str(row.get("source") or path.name).strip() or path.name,
            }
    return {"provided": True, "by_id": by_id, "path": str(path)}


def _human_decision(
    channel_id: str,
    human_review: Mapping[str, Any],
    snapshot_review: Mapping[str, Any],
) -> Tuple[str, str, str]:
    if human_review.get("provided"):
        record = (human_review.get("by_id") or {}).get(_fold(channel_id))
        if not record:
            return "not_listed", "", ""
        return record["decision"], record["source"], record["reason"]
    state, note = _review_for(channel_id, snapshot_review)
    if state in {"kept", "dropped"}:
        return state, "snapshot.description_review", note
    return "not_listed", "", ""


def _acceptance_counts(
    counts: Mapping[str, int],
    audit: Sequence[Mapping[str, Any]],
    review: Sequence[Mapping[str, Any]],
    human_review: Mapping[str, Any],
    keywords: Sequence[str],
) -> Dict[str, int]:
    merged = dict(counts)
    if human_review.get("provided"):
        labels = list((human_review.get("by_id") or {}).values())
        kept_labels = sum(1 for item in labels if item.get("decision") == "kept")
        dropped_labels = sum(1 for item in labels if item.get("decision") == "dropped")
    else:
        kept_labels = sum(1 for row in audit if row.get("human_review_decision") == "kept")
        dropped_labels = sum(1 for row in audit if row.get("human_review_decision") == "dropped")
    theme_hits = merged.get("theme_matched_videos") or 0
    unique_videos = merged.get("unique_videos") or 0
    merged.update(
        {
            "audit_rows": len(audit),
            "review_shortlist_rows": len(review),
            "review_shortlist_before_cap": sum(1 for row in audit if row.get("human_review_decision") == "kept"),
            "shortlist_before_cap": sum(1 for row in audit if row.get("human_review_decision") == "kept"),
            "shortlist_written": len(review),
            "shortlist_truncated": max(
                0,
                sum(1 for row in audit if row.get("human_review_decision") == "kept") - len(review),
            ),
            "ready_to_send_rows": sum(1 for row in audit if str(row.get("ready_to_send")) == "yes"),
            "human_dropped_in_audit": sum(1 for row in audit if row.get("human_review_decision") == "dropped"),
            "not_listed_in_audit": sum(1 for row in audit if row.get("human_review_decision") == "not_listed"),
            "pending_identity_check_not_recommended": sum(
                1 for row in audit if row.get("queue") == "pending_identity_check_not_recommended"
            ),
            "threshold_exceptions_in_shortlist": sum(1 for row in review if row.get("threshold_exception")),
            "human_review_kept_labels": kept_labels,
            "human_review_dropped_labels": dropped_labels,
            "human_review_file_provided": 1 if human_review.get("provided") else 0,
            "theme_keyword_is_relevance": 0,
            "theme_filter_discriminative": 0 if unique_videos and theme_hits / unique_videos >= 0.9 else 1,
            "original_filter_passed_in_shortlist": sum(
                1 for row in review if row.get("any_video_passed_original_filter") == "yes"
            ),
        }
    )
    if not keywords:
        merged["theme_filter_discriminative"] = 0
    return merged


def load_strong_channel_ids(paths: Sequence[Path]) -> StrongIdIndex:
    """Load explicit channel IDs. Name, handle and email keys are ignored."""
    index = StrongIdIndex()
    for path in paths:
        if not path.is_file():
            raise SnapshotUsageError(f"找不到强频道 ID 文件: {path}")
        suffix = path.suffix.lower()
        if suffix == ".json":
            _load_strong_json(index, path)
        else:
            _load_strong_csv(index, path)
    return index


def render_report(summary: Mapping[str, Any]) -> str:
    """Chinese trial report. It states the offline boundary in the opening."""
    counts = summary["counts"]
    lines = [
        "# 离线 snapshot 试跑报告",
        "",
        "本次只重放本地 snapshot，复用仓库里的解析、视频筛选、频道查重、分类和 Excel 导出。",
        "这不是端到端在线自动化通过，没有调用 YouTube，也没有发信或写飞书。",
        "`ready_to_send` 固定为 no。",
        "",
        "## 输入与重算",
        "",
        f"- 搜索行: {counts['search_rows']}",
        f"- 去重后视频: {counts['unique_videos']}",
        f"- 去重后频道: {counts['unique_channels']}",
        f"- 有视频详情的视频: {counts['videos_with_details']}",
        f"- 有频道详情的频道: {counts['channels_with_details']}",
        f"- 搜索里缺少视频详情: {counts['search_videos_missing_details']}",
        f"- 搜索里缺少频道详情: {counts['search_channels_missing_details']}",
        "",
        "## 筛选、查重、排除、输出",
        "",
        f"- 原视频筛选通过: {counts['original_filter_qualified_videos']}",
        f"- 原筛选品牌排除视频: {counts['original_filter_brand_excluded_videos']}",
        f"- 原筛选未达标视频: {counts['original_filter_below_threshold_videos']}",
        f"- 达标视频去重后的新频道: {counts['deduped_new_channels']}",
        f"- 达标视频里命中强频道 ID 的频道: {counts['deduped_existing_strong_id_channels']}",
        f"- 强频道 ID 索引（按频道 ID 折叠后）: {counts['strong_id_index_size']}",
        f"- snapshot 内命中强频道 ID 的频道: {counts['snapshot_channels_strong_id_hits']}",
        f"- 品牌排除频道: {counts['brand_excluded_channels']}",
        f"- 订阅隐藏: {counts['subscriber_hidden']}",
        f"- 订阅缺失或非数字: {counts['subscriber_missing']}",
        f"- 订阅确认为 0: {counts['subscriber_known_zero']}",
        f"- 主订阅区间频道: {counts['band_primary_channels']}",
        f"- 扩展订阅区间频道: {counts['band_extended_channels']}",
        f"- 主题命中视频 / 频道: {counts['theme_matched_videos']} / {counts['theme_matched_channels']}",
        f"- audit 行: {counts['audit_rows']}",
        f"- 人工 kept 标签: {counts['human_review_kept_labels']}",
        f"- 人工 dropped 标签: {counts['human_review_dropped_labels']}",
        f"- audit 中的 dropped: {counts['human_dropped_in_audit']}",
        f"- audit 中未进入人工复核: {counts['not_listed_in_audit']}",
        f"- 待身份核查且不推荐: {counts['pending_identity_check_not_recommended']}",
        f"- 复核短名单: {counts['review_shortlist_rows']}",
        f"- 短名单中的人工保留例外: {counts['threshold_exceptions_in_shortlist']}",
        f"- ready_to_send 为 yes 的行: {counts['ready_to_send_rows']}",
        "",
        "audit 是订阅区间内、且没有被强频道 ID 或品牌频道 ID 排除的机械结果。它不是推荐名单。",
        "只有人工 decision=kept 进入复核短名单。dropped 和 not_listed 不能因主题词或订阅区间回到推荐。",
        "UNKNOWN 不是通过。创作者身份 UNKNOWN 不会被写成已验证，ready_to_send 固定为 no。",
        "强频道 ID 和品牌频道 ID 优先于人工 kept。频道名像品牌，不能单独当成官方或达人结论。",
        "原视频筛选结果保持不变。人工保留但未过自动阈值的人带 human_retained_exception_auto_threshold_not_passed，不改成通过。",
        "历史强频道 ID 没有命中，只表示这张表里没有这条 ID，不表示从未联系。",
        "",
    ]
    if (
        counts.get("theme_filter_applied")
        and counts["unique_videos"]
        and counts["theme_matched_videos"] / counts["unique_videos"] >= 0.9
    ):
        lines.append(
            "主题词命中率很高，说明它没有区分力。命中主题词不是相关性结论，"
            "也不能把人工 dropped 的泛内容、非目标语言或地区、无关木工送回复核短名单。"
        )
        lines.append("")
    lines.extend(["### snapshot 内的强频道 ID 命中", ""])
    strong_hits = summary.get("strong_id_hits") or []
    if not strong_hits:
        lines.append("- 无")
    else:
        for row in strong_hits:
            lines.append(
                f"- `{row['channel_id']}` {row['channel_name']}: {row['subscriber_band']}; {row['reason']}"
            )
    lines.extend(["", "### 区间内被排除", ""])
    excluded = summary.get("in_band_excluded") or []
    if not excluded:
        lines.append("- 无")
    else:
        for row in excluded:
            lines.append(
                f"- `{row['channel_id']}` {row['channel_name']}: {row['reason']}"
            )
    lines.extend(["", "## 复核短名单", ""])
    candidates = summary.get("review_shortlist") or []
    if not candidates:
        lines.append("复核短名单为空。没有把 dropped、not_listed 或 UNKNOWN 补进去。")
    else:
        lines.append(
            "| channel_id | name | band | subs | original_filter | human_review | exception |"
        )
        lines.append("|---|---|---|---|---|---|---|")
        for row in candidates:
            lines.append(
                "| `{channel_id}` | {channel_name} | {subscriber_band} | {subscriber_count} | "
                "{original_video_filter} | {human_review_decision} | {threshold_exception} |".format(**row)
            )
        lines.append("")
        lines.append("这些行仍是待人工复核，ready_to_send=no。身份字段 UNKNOWN 不是验证通过。")
    lines.extend(["", "## 不推荐队列", ""])
    identity_rows = [
        row for row in (summary.get("audit") or [])
        if row.get("queue") != "review_shortlist"
    ]
    if not identity_rows:
        lines.append("- 无")
    else:
        for row in identity_rows:
            lines.append(
                f"- `{row['channel_id']}` {row['channel_name']}: {row['queue']}; "
                f"human_review={row['human_review_decision']}; creator_identity=UNKNOWN"
            )
    lines.extend(["", "## 边界", ""])
    for item in summary.get("boundaries") or []:
        lines.append(f"- {item}")
    lines.extend(["", "## 复用、适配器、未测", ""])
    lines.append("复用:")
    for item in summary.get("reused") or []:
        lines.append(f"- `{item}`")
    lines.append("")
    lines.append("适配器:")
    for item in summary.get("adapter") or []:
        lines.append(f"- {item}")
    lines.append("")
    lines.append("未测:")
    for item in summary.get("not_tested") or []:
        lines.append(f"- {item}")
    lines.extend(
        [
            "",
            "## 可重复命令",
            "",
            "```bash",
            str(summary.get("command") or ""),
            "```",
            "",
            "## 下一步只需验证",
            "",
            "- 人工观看代表视频，确认是否为独立创作者；分类词和频道名都不是品牌身份结论。",
            "- 受众地区、Amazon 原生商品标签、邮箱、联系史、退订和退信仍是 UNKNOWN，不能据此发信。",
            "- 强频道 ID 未命中的人，仍要在业务侧核验联系史。",
            "- 在线 YouTube、飞书、17TRACK 和发信都还没有跑，不能把本次记成线上通过。",
            "- 不需要为了凑满 10 人再搜索新候选。",
            "",
        ]
    )
    return _scrub("\n".join(lines))


def _build_shortlist(
    *,
    videos: Sequence[Dict[str, Any]],
    channels: Mapping[str, Dict[str, Any]],
    brand: BrandExclusion,
    strong_ids: StrongIdIndex,
    keywords: Sequence[str],
    review: Mapping[str, Any],
    human_review: Mapping[str, Any],
    options: SnapshotOptions,
    primary_label: str,
    extended_label: str,
) -> Tuple[List[Dict[str, str]], List[Dict[str, str]], List[Dict[str, str]]]:
    by_channel: Dict[str, List[Dict[str, Any]]] = {}
    for video in videos:
        by_channel.setdefault(video.get("channel_id", ""), []).append(video)

    fidelity: List[Dict[str, str]] = []
    excluded: List[Dict[str, str]] = []
    shortlist: List[Dict[str, Any]] = []
    for channel_id, channel in channels.items():
        if channel.get("subscriber_count_status") != "known":
            fidelity.append(
                {
                    "channel_id": channel_id,
                    "status": str(channel.get("subscriber_count_status") or "missing"),
                    "parser_would_write": str(channel.get("subscriber_parser_value")),
                    "exported_value": "UNKNOWN",
                }
            )
        band = _subscriber_band(channel, options, primary_label, extended_label)
        channel_videos = by_channel.get(channel_id, [])
        title = channel.get("channel_title") or _first_title(channel_videos)
        theme_videos = _theme_videos(channel_videos, keywords)
        if band is None:
            continue
        if brand.is_excluded(channel_id, title):
            excluded.append(
                {
                    "channel_id": channel_id,
                    "channel_name": title,
                    "subscriber_band": band,
                    "reason": "brand_exclusion",
                }
            )
            continue
        if strong_ids.contains(channel_id):
            actions = ",".join(dict.fromkeys(_strong_labels(strong_ids.records_for(channel_id))))
            excluded.append(
                {
                    "channel_id": channel_id,
                    "channel_name": title,
                    "subscriber_band": band,
                    "reason": f"strong_channel_id:{actions}",
                }
            )
            continue
        if keywords and not theme_videos:
            excluded.append(
                {
                    "channel_id": channel_id,
                    "channel_name": title,
                    "subscriber_band": band,
                    "reason": "theme_not_matched",
                }
            )
            continue
        pool = theme_videos or list(channel_videos)
        representative = _representative(pool)
        row = _candidate_row(
            channel=channel,
            representative=representative,
            channel_videos=channel_videos,
            band=band,
            keywords=keywords,
            review=review,
            human_review=human_review,
            extended_label=extended_label,
        )
        shortlist.append(row)

    shortlist.sort(key=_sort_key)
    return shortlist, excluded, fidelity


def _candidate_row(
    *,
    channel: Mapping[str, Any],
    representative: Mapping[str, Any],
    channel_videos: Sequence[Mapping[str, Any]],
    band: str,
    keywords: Sequence[str],
    review: Mapping[str, Any],
    human_review: Mapping[str, Any],
    extended_label: str,
) -> Dict[str, Any]:
    channel_id = str(channel.get("channel_id") or "")
    passed_any = any(bool(video.get("is_qualified")) for video in channel_videos)
    passed_rep = bool(representative.get("is_qualified"))
    hits = _keyword_hits(_match_text(representative), keywords)
    country = _creator_country(channel.get("country"))
    subs_status = str(channel.get("subscriber_count_status") or "missing")
    subs_value = channel.get("subscriber_count")
    view_status = str(representative.get("view_count_status") or "missing")
    view_value = representative.get("view_count_truth")
    engagement = representative.get("engagement_truth")
    duration_status = str(representative.get("duration_status") or "missing")
    review_state, review_note = _review_for(channel_id, review)
    decision, review_source, review_reason = _human_decision(channel_id, human_review, review)
    text = _match_text(representative) + " " + str(channel.get("channel_description") or "")
    taxonomy = classify_channel(
        str(channel.get("channel_title") or ""),
        str(channel.get("channel_description") or ""),
        channel_videos,
        " ".join(representative.get("search_queries") or []),
    )
    bucket = "original_filter" if passed_any else "below_original_threshold"
    if band == extended_label:
        bucket += "_extended"
    else:
        bucket += "_primary"
    if decision == "kept" and not passed_any:
        threshold_exception = "human_retained_exception_auto_threshold_not_passed"
        bucket = "human_retained_exception_auto_threshold_not_passed"
        if band == extended_label:
            bucket += "_extended"
    else:
        threshold_exception = ""
    if decision == "kept":
        queue = "review_shortlist"
        recommendation = "review_shortlist"
    elif decision == "dropped":
        queue = "human_dropped_not_recommended"
        recommendation = "not_recommended"
    else:
        queue = "pending_identity_check_not_recommended"
        recommendation = "not_recommended"
    fit = (
        f"主题关键词命中只是审计线索: {', '.join(hits) if hits else '无'}; "
        f"订阅区间 {band}; 原视频筛选 {'pass' if passed_rep else 'fail'}"
        f"({representative.get('filter_reason') or ''})。"
        f"人工复核={decision}。"
        + (f"{threshold_exception}。" if threshold_exception else "")
        + "创作者身份保持 UNKNOWN。受众、邮箱、原生标签和联系史未核验。"
    )
    return {
        "channel_id": channel_id,
        "channel_name": str(channel.get("channel_title") or ""),
        "channel_url": str(channel.get("channel_url") or ""),
        "creator_country": country,
        "audience_geo_status": "UNKNOWN",
        "contact": "UNKNOWN",
        "email_status": "UNKNOWN",
        "email_seen_in_public_text": "yes" if EMAIL_PATTERN.search(text) else "no",
        "subscriber_count": str(subs_value) if subs_status == "known" else "UNKNOWN",
        "subscriber_count_status": subs_status,
        "subscriber_band": band,
        "hidden_subscribers": _hidden_text(channel),
        "brand_channel_check": "UNKNOWN",
        "taxonomy_preliminary": taxonomy,
        "video_id": str(representative.get("video_id") or ""),
        "video_url": str(representative.get("video_url") or ""),
        "video_title": html.unescape(str(representative.get("title") or "")),
        "video_published_at": str(
            representative.get("published_at_raw") or representative.get("published_at") or ""
        ),
        "video_views": str(view_value) if view_status == "known" else "UNKNOWN",
        "video_view_status": view_status,
        "video_engagement_rate": f"{engagement:.2f}" if isinstance(engagement, float) else "UNKNOWN",
        "duration_seconds": (
            str(representative.get("duration_seconds_truth"))
            if duration_status == "known"
            else "UNKNOWN"
        ),
        "duration_hms": str(representative.get("duration_hms") or "") if duration_status == "known" else "UNKNOWN",
        "duration_note": "under_60s" if _is_short(representative) else "",
        "search_query": " | ".join(representative.get("search_queries") or []),
        "theme_keywords_hit": ", ".join(hits),
        "original_video_filter": "pass" if passed_rep else "fail",
        "original_filter_reason": str(representative.get("filter_reason") or ""),
        "any_video_passed_original_filter": "yes" if passed_any else "no",
        "fit_reason": fit,
        "amazon_link_seen": "yes" if AMAZON_PATTERN.search(text) else "no",
        "amazon_native_tag": "UNKNOWN",
        "strong_id_match": "no",
        "contact_history": "UNKNOWN",
        "unsubscribe_status": "UNKNOWN",
        "bounce_status": "UNKNOWN",
        "brand_history": "UNKNOWN",
        "nox_status": "UNKNOWN",
        "snapshot_review": review_state,
        "snapshot_review_note": review_note,
        "human_review_decision": decision,
        "human_review_source": review_source,
        "human_review_note": review_reason,
        "queue": queue,
        "recommendation": recommendation,
        "creator_identity": "UNKNOWN",
        "identity_verified": "no",
        "threshold_exception": threshold_exception,
        "theme_keyword_is_relevance": "no",
        "ready_to_send": "no",
        "evidence_type": "public_title_and_description_in_local_snapshot",
        "trial_bucket": bucket,
        "_views_sort": view_value if view_status == "known" else -1,
        "_passed_any": passed_any,
        "_extended": band == extended_label,
    }


def _normalize_search(snapshot: Mapping[str, Any]) -> Tuple[List[Dict[str, Any]], str]:
    rows: List[Dict[str, Any]] = []
    shapes = set()
    for response in snapshot.get("search_responses") or []:
        if not isinstance(response, dict):
            continue
        query = str(response.get("query") or (response.get("params") or {}).get("q") or "")
        for item in response.get("items") or []:
            if not isinstance(item, dict):
                continue
            raw = isinstance(item.get("id"), dict) or isinstance(item.get("snippet"), dict)
            shapes.add("raw" if raw else "flat")
            if raw:
                snippet = item.get("snippet") or {}
                video_id = ""
                if isinstance(item.get("id"), dict):
                    video_id = str(item["id"].get("videoId") or "")
                video_id = video_id or str(item.get("videoId") or "")
                rows.append(
                    {
                        "video_id": video_id,
                        "channel_id": str(snippet.get("channelId") or item.get("channelId") or ""),
                        "title": str(snippet.get("title") or item.get("title") or ""),
                        "description": str(snippet.get("description") or item.get("description") or ""),
                        "query": query,
                    }
                )
            else:
                rows.append(
                    {
                        "video_id": str(item.get("videoId") or ""),
                        "channel_id": str(item.get("channelId") or ""),
                        "title": str(item.get("title") or ""),
                        "description": str(item.get("description") or ""),
                        "query": query,
                    }
                )
    return rows, _shape_label(shapes)


def _normalize_videos(snapshot: Mapping[str, Any]) -> Tuple[List[Dict[str, Any]], str]:
    videos: List[Dict[str, Any]] = []
    seen = set()
    shapes = set()
    for response in snapshot.get("videos_responses") or []:
        if not isinstance(response, dict):
            continue
        for item in response.get("items") or []:
            if not isinstance(item, dict):
                continue
            raw_shape = isinstance(item.get("snippet"), dict) or isinstance(item.get("statistics"), dict)
            shapes.add("raw" if raw_shape else "flat")
            video = _normalize_one_video(item, raw_shape)
            video_id = video.get("video_id") or ""
            if not video_id or video_id in seen:
                continue
            seen.add(video_id)
            videos.append(video)
    return videos, _shape_label(shapes)


def _normalize_one_video(item: Mapping[str, Any], raw_shape: bool) -> Dict[str, Any]:
    view_status, view_value = _metric(item, "viewCount", raw_shape)
    like_status, like_value = _metric(item, "likeCount", raw_shape)
    comment_status, comment_value = _metric(item, "commentCount", raw_shape)
    duration_raw = _duration_raw(item, raw_shape)
    if raw_shape:
        parsed = _parse_video_item(dict(item))
    else:
        parsed = _parse_video_item(_flat_video_raw(item, view_status, view_value, like_status, like_value, comment_status, comment_value, duration_raw))
    parsed["view_count_status"] = view_status
    parsed["view_count_truth"] = view_value
    parsed["like_count_status"] = like_status
    parsed["like_count_truth"] = like_value
    parsed["comment_count_status"] = comment_status
    parsed["comment_count_truth"] = comment_value
    parsed["engagement_truth"] = _engagement(view_status, view_value, like_status, like_value, comment_status, comment_value)
    if duration_raw:
        parsed["duration_status"] = "known"
        parsed["duration_seconds_truth"] = parsed.get("duration_seconds")
    else:
        parsed["duration_status"] = "missing"
        parsed["duration_seconds_truth"] = None
    parsed["description"] = _video_description(item, raw_shape)
    parsed["search_queries"] = []
    return parsed


def _normalize_channels(snapshot: Mapping[str, Any]) -> Tuple[Dict[str, Dict[str, Any]], str]:
    channels: Dict[str, Dict[str, Any]] = {}
    shapes = set()
    for response in snapshot.get("channels_responses") or []:
        if not isinstance(response, dict):
            continue
        for item in response.get("items") or []:
            if not isinstance(item, dict):
                continue
            raw_shape = isinstance(item.get("snippet"), dict) or isinstance(item.get("statistics"), dict)
            shapes.add("raw" if raw_shape else "flat")
            channel = _normalize_one_channel(item, raw_shape)
            channel_id = channel.get("channel_id") or ""
            if channel_id and channel_id not in channels:
                channels[channel_id] = channel
    return channels, _shape_label(shapes)


def _normalize_one_channel(item: Mapping[str, Any], raw_shape: bool) -> Dict[str, Any]:
    hidden = _hidden_subscriber_flag(item, raw_shape)
    present, raw_value = _subscriber_raw(item, raw_shape)
    status, value = _public_count(raw_value, present=present, hidden=hidden)
    if raw_shape:
        parsed = _parse_channel_item(dict(item))
    else:
        parsed = _parse_channel_item(_flat_channel_raw(item, status, value, hidden))
    parsed["subscriber_parser_value"] = parsed.get("subscriber_count")
    parsed["subscriber_count_status"] = status
    parsed["subscriber_count"] = value if status == "known" else None
    parsed["hidden_subscriber_flag"] = hidden
    if not raw_shape and not item.get("description"):
        parsed["channel_description"] = parsed.get("channel_description") or ""
    return parsed


def _attach_search_context(videos: Sequence[Dict[str, Any]], search_rows: Sequence[Mapping[str, Any]]) -> None:
    by_id = {video.get("video_id"): video for video in videos}
    for row in search_rows:
        video = by_id.get(row.get("video_id"))
        if not video:
            continue
        query = str(row.get("query") or "")
        if query and query not in video["search_queries"]:
            video["search_queries"].append(query)
        if not video.get("description") and row.get("description"):
            video["description"] = row["description"]
        if query:
            video["搜索关键词"] = " | ".join(video["search_queries"])


def _load_strong_csv(index: StrongIdIndex, path: Path) -> None:
    with path.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        fieldnames = set(reader.fieldnames or [])
        if {"key_type", "key_value"} <= fieldnames:
            for row in reader:
                key_type = str(row.get("key_type") or "").strip()
                if key_type not in STRONG_KEY_TYPES:
                    continue
                evidence = str(row.get("evidence_strength") or "").strip()
                if evidence not in {"", "strong"}:
                    continue
                _add_strong_id(
                    index,
                    str(row.get("key_value") or ""),
                    source=path.name,
                    key_type=key_type,
                    action=str(row.get("action") or "").strip(),
                    evidence_strength=evidence or "strong",
                )
            return
        column = "channel_id" if "channel_id" in fieldnames else "channelId" if "channelId" in fieldnames else ""
        if not column:
            raise SnapshotUsageError(f"强频道 ID CSV 缺少 channel_id 或 key_type/key_value: {path}")
        for row in reader:
            _add_strong_id(
                index,
                str(row.get(column) or ""),
                source=path.name,
                key_type=column,
                action=str(row.get("action") or "").strip(),
                evidence_strength="explicit_column",
            )


def _load_strong_json(index: StrongIdIndex, path: Path) -> None:
    data = _load_json(path)
    values: Iterable[Any]
    if isinstance(data, dict):
        values = data.get("channel_ids") or []
    elif isinstance(data, list):
        values = data
    else:
        raise SnapshotUsageError(f"强频道 ID JSON 无法识别: {path}")
    for value in values:
        _add_strong_id(
            index,
            str(value or ""),
            source=path.name,
            key_type="channel_id",
            action="",
            evidence_strength="explicit_column",
        )


def _strong_labels(records: Sequence[Mapping[str, str]]) -> List[str]:
    return [str(item.get("action") or item.get("source") or "strong_channel_id") for item in records]


def _snapshot_strong_hits(
    channels: Mapping[str, Mapping[str, Any]],
    strong_ids: StrongIdIndex,
    options: SnapshotOptions,
    primary_label: str,
    extended_label: str,
) -> List[Dict[str, str]]:
    rows = []
    for channel_id, channel in channels.items():
        if not strong_ids.contains(channel_id):
            continue
        status = str(channel.get("subscriber_count_status") or "missing")
        value = channel.get("subscriber_count")
        band = _subscriber_band(channel, options, primary_label, extended_label)
        rows.append(
            {
                "channel_id": channel_id,
                "channel_name": str(channel.get("channel_title") or ""),
                "subscriber_band": band or "out_of_band_or_unknown_subscribers",
                "subscriber_count": str(value) if status == "known" else "UNKNOWN",
                "reason": ",".join(dict.fromkeys(_strong_labels(strong_ids.records_for(channel_id)))),
            }
        )
    rows.sort(key=lambda row: row["channel_id"])
    return rows


def _add_strong_id(
    index: StrongIdIndex,
    channel_id: str,
    *,
    source: str,
    key_type: str,
    action: str,
    evidence_strength: str,
) -> None:
    channel_id = channel_id.strip()
    if not channel_id:
        return
    index.hits.setdefault(_fold(channel_id), []).append(
        {
            "channel_id": channel_id,
            "source": source,
            "key_type": key_type,
            "action": action,
            "evidence_strength": evidence_strength,
        }
    )


def _load_brand_exclusion(path: Optional[Path]) -> BrandExclusion:
    if path is None:
        return BrandExclusion()
    if not path.is_file():
        raise SnapshotUsageError(f"找不到品牌排除文件: {path}")
    return BrandExclusion.load(str(path))


def _load_keywords(path: Optional[Path]) -> List[str]:
    if path is None:
        return []
    if not path.is_file():
        raise SnapshotUsageError(f"找不到主题关键词文件: {path}")
    keywords = []
    for line in path.read_text(encoding="utf-8").splitlines():
        text = line.strip()
        if text and not text.startswith("#"):
            keywords.append(text)
    return keywords


def _load_json(path: Path) -> Any:
    if not path.is_file():
        raise SnapshotUsageError(f"找不到文件: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise SnapshotUsageError(f"JSON 无法解析: {path}") from exc


def _public_count(value: Any, *, present: bool, hidden: bool = False) -> Tuple[str, Optional[int]]:
    if hidden:
        return "hidden", None
    if not present or value is None or isinstance(value, bool):
        return "missing", None
    if isinstance(value, (int, float)):
        number = int(value)
        return ("known", number) if number >= 0 else ("missing", None)
    text = str(value).strip()
    if text == "" or text.casefold() in NON_NUMERIC:
        return "missing", None
    if text.isdigit():
        return "known", int(text)
    return "missing", None


def _metric(item: Mapping[str, Any], key: str, raw_shape: bool) -> Tuple[str, Optional[int]]:
    if raw_shape:
        stats = item.get("statistics") if isinstance(item.get("statistics"), dict) else {}
        return _public_count(stats.get(key), present=key in stats)
    return _public_count(item.get(key), present=key in item)


def _hidden_subscriber_flag(item: Mapping[str, Any], raw_shape: bool) -> bool:
    if raw_shape:
        stats = item.get("statistics") if isinstance(item.get("statistics"), dict) else {}
        flag = stats.get("hiddenSubscriberCount")
    else:
        flag = item.get("hiddenSubscribers", item.get("hiddenSubscriberCount"))
    return flag is True or str(flag).strip().lower() in {"true", "1", "yes"}


def _subscriber_raw(item: Mapping[str, Any], raw_shape: bool) -> Tuple[bool, Any]:
    if raw_shape:
        stats = item.get("statistics") if isinstance(item.get("statistics"), dict) else {}
        return "subscriberCount" in stats, stats.get("subscriberCount")
    return "subscriberCount" in item, item.get("subscriberCount")


def _flat_video_raw(
    item: Mapping[str, Any],
    view_status: str,
    view_value: Optional[int],
    like_status: str,
    like_value: Optional[int],
    comment_status: str,
    comment_value: Optional[int],
    duration_raw: str,
) -> Dict[str, Any]:
    stats: Dict[str, Any] = {}
    if view_status == "known":
        stats["viewCount"] = view_value
    if like_status == "known":
        stats["likeCount"] = like_value
    if comment_status == "known":
        stats["commentCount"] = comment_value
    tags = item.get("tags") or []
    if isinstance(tags, str):
        tags = [part.strip() for part in tags.split(",") if part.strip()]
    return {
        "id": item.get("videoId") or item.get("id") or "",
        "snippet": {
            "channelId": item.get("channelId") or "",
            "channelTitle": item.get("channelTitle") or "",
            "title": item.get("title") or "",
            "description": item.get("description") or "",
            "publishedAt": item.get("publishedAt") or "",
            "liveBroadcastContent": item.get("liveBroadcastContent") or "none",
            "tags": tags,
        },
        "statistics": stats,
        "contentDetails": {
            "duration": duration_raw,
            "caption": item.get("caption") or "false",
        },
    }


def _flat_channel_raw(
    item: Mapping[str, Any],
    status: str,
    value: Optional[int],
    hidden: bool,
) -> Dict[str, Any]:
    stats: Dict[str, Any] = {}
    if hidden:
        stats["hiddenSubscriberCount"] = True
    if status == "known":
        stats["subscriberCount"] = value
    related = item.get("relatedPlaylists")
    if not isinstance(related, dict):
        related = {}
    return {
        "id": item.get("channelId") or item.get("id") or "",
        "snippet": {
            "title": item.get("title") or item.get("channelTitle") or "",
            "description": item.get("description") or "",
            "publishedAt": item.get("publishedAt") or "",
            "country": item.get("country") or "",
        },
        "statistics": stats,
        "brandingSettings": {"channel": {}},
        "contentDetails": {"relatedPlaylists": related},
    }


def _duration_raw(item: Mapping[str, Any], raw_shape: bool) -> str:
    if raw_shape:
        content = item.get("contentDetails") if isinstance(item.get("contentDetails"), dict) else {}
        return str(content.get("duration") or "")
    return str(item.get("duration") or "")


def _video_description(item: Mapping[str, Any], raw_shape: bool) -> str:
    if raw_shape:
        snippet = item.get("snippet") if isinstance(item.get("snippet"), dict) else {}
        return str(snippet.get("description") or "")
    return str(item.get("description") or "")


def _engagement(
    view_status: str,
    view_value: Optional[int],
    like_status: str,
    like_value: Optional[int],
    comment_status: str,
    comment_value: Optional[int],
) -> Optional[float]:
    if view_status != "known" or not view_value or like_status != "known" or comment_status != "known":
        return None
    return round(((like_value or 0) + (comment_value or 0)) / view_value * 100, 2)


def _subscriber_band(
    channel: Mapping[str, Any],
    options: SnapshotOptions,
    primary_label: str,
    extended_label: str,
) -> Optional[str]:
    if channel.get("subscriber_count_status") != "known":
        return None
    value = channel.get("subscriber_count")
    if not isinstance(value, int):
        return None
    if options.min_subscribers <= value <= options.primary_max:
        return primary_label
    if options.primary_max < value <= options.extended_max:
        return extended_label
    return None


def _band_labels(options: SnapshotOptions) -> Tuple[str, str]:
    if options.min_subscribers == 10000 and options.primary_max == 50000 and options.extended_max == 100000:
        return "10K-50K", "50K-100K_extended"
    return (
        f"{options.min_subscribers}-{options.primary_max}",
        f"{options.primary_max + 1}-{options.extended_max}_extended",
    )


def _theme_videos(videos: Sequence[Dict[str, Any]], keywords: Sequence[str]) -> List[Dict[str, Any]]:
    if not keywords:
        return list(videos)
    return [video for video in videos if _keyword_hits(_match_text(video), keywords)]


def _keyword_hits(text: str, keywords: Sequence[str]) -> List[str]:
    folded = html.unescape(text or "").casefold()
    return [keyword for keyword in keywords if keyword.casefold() in folded]


def _match_text(video: Mapping[str, Any]) -> str:
    """Title and description only. The search query must not create a theme hit."""
    return f"{video.get('title') or ''} {video.get('description') or ''}"


def _representative(videos: Sequence[Mapping[str, Any]]) -> Mapping[str, Any]:
    def rank(video: Mapping[str, Any]) -> Tuple[int, int, str]:
        known = video.get("view_count_status") == "known" and isinstance(video.get("view_count_truth"), int)
        views = int(video["view_count_truth"]) if known else -1
        return (0 if known else 1, -views, str(video.get("video_id") or ""))

    return sorted(videos, key=rank)[0]


def _sort_key(row: Mapping[str, Any]) -> Tuple[int, int, int, str]:
    return (
        0 if row.get("_passed_any") else 1,
        1 if row.get("_extended") else 0,
        -int(row.get("_views_sort") or -1),
        str(row.get("channel_id") or ""),
    )


def _review_index(value: Any) -> Dict[str, Any]:
    if not isinstance(value, dict):
        return {"kept": set(), "dropped": {}}
    kept = value.get("kept_channel_ids") or []
    dropped = value.get("dropped") or {}
    return {
        "kept": {str(item) for item in kept} if isinstance(kept, list) else set(),
        "dropped": dropped if isinstance(dropped, dict) else {},
    }


def _review_for(channel_id: str, review: Mapping[str, Any]) -> Tuple[str, str]:
    dropped = review.get("dropped") or {}
    if channel_id in dropped:
        return "dropped", str(dropped.get(channel_id) or "")
    if channel_id in (review.get("kept") or set()):
        return "kept", ""
    return "not_listed", ""


def _creator_country(value: Any) -> str:
    text = str(value or "").strip()
    if text.casefold() in {"", "unknown", "none", "null"}:
        return "UNKNOWN"
    return text


def _hidden_text(channel: Mapping[str, Any]) -> str:
    if channel.get("hidden_subscriber_flag") is True:
        return "true"
    if channel.get("subscriber_count_status") == "known":
        return "false"
    return "UNKNOWN"


def _is_short(video: Mapping[str, Any]) -> bool:
    seconds = video.get("duration_seconds_truth")
    return isinstance(seconds, int) and seconds < 60


def _first_title(videos: Sequence[Mapping[str, Any]]) -> str:
    for video in videos:
        title = str(video.get("channel_title") or "")
        if title:
            return title
    return ""


def _counts(
    *,
    search_rows: Sequence[Mapping[str, Any]],
    videos: Sequence[Mapping[str, Any]],
    channels: Mapping[str, Mapping[str, Any]],
    qualified: Sequence[Mapping[str, Any]],
    new_channels: Sequence[Mapping[str, Any]],
    existing_channels: Sequence[Mapping[str, Any]],
    shortlist: Sequence[Mapping[str, Any]],
    written: Sequence[Mapping[str, Any]],
    strong_ids: StrongIdIndex,
    fidelity: Sequence[Mapping[str, str]],
    brand: BrandExclusion,
    keywords: Sequence[str],
    options: SnapshotOptions,
    primary_label: str,
    extended_label: str,
) -> Dict[str, int]:
    search_video_ids = {row.get("video_id") for row in search_rows if row.get("video_id")}
    search_channel_ids = {row.get("channel_id") for row in search_rows if row.get("channel_id")}
    video_ids = {video.get("video_id") for video in videos}
    brand_excluded_videos = sum(1 for video in videos if video.get("filter_reason") == "品牌官方频道")
    qualified_ids = {id(video) for video in qualified}
    below = sum(
        1
        for video in videos
        if id(video) not in qualified_ids and video.get("filter_reason") != "品牌官方频道"
    )
    theme_video_ids = {
        video.get("channel_id")
        for video in videos
        if keywords and _keyword_hits(_match_text(video), keywords)
    }
    theme_video_count = sum(
        1 for video in videos if keywords and _keyword_hits(_match_text(video), keywords)
    )
    band_primary = 0
    band_extended = 0
    brand_excluded_channels = 0
    for channel_id, channel in channels.items():
        title = str(channel.get("channel_title") or "")
        if brand.is_excluded(channel_id, title):
            brand_excluded_channels += 1
        band = _subscriber_band(channel, options, primary_label, extended_label)
        if band == primary_label:
            band_primary += 1
        elif band == extended_label:
            band_extended += 1
    return {
        "search_rows": len(search_rows),
        "unique_search_videos": len(search_video_ids),
        "unique_search_channels": len(search_channel_ids),
        "unique_videos": len(video_ids),
        "unique_channels": len(channels),
        "videos_with_details": len(videos),
        "channels_with_details": len(channels),
        "search_videos_missing_details": len(search_video_ids - video_ids),
        "search_channels_missing_details": len(search_channel_ids - set(channels)),
        "original_filter_qualified_videos": len(qualified),
        "original_filter_brand_excluded_videos": brand_excluded_videos,
        "original_filter_below_threshold_videos": below,
        "deduped_new_channels": len(new_channels),
        "deduped_existing_strong_id_channels": len(existing_channels),
        "strong_id_index_size": len(strong_ids),
        "snapshot_channels_strong_id_hits": sum(1 for channel_id in channels if strong_ids.contains(channel_id)),
        "brand_excluded_channels": brand_excluded_channels,
        "subscriber_hidden": sum(1 for item in fidelity if item["status"] == "hidden"),
        "subscriber_missing": sum(1 for item in fidelity if item["status"] == "missing"),
        "subscriber_known_zero": sum(
            1
            for channel in channels.values()
            if channel.get("subscriber_count_status") == "known" and channel.get("subscriber_count") == 0
        ),
        "band_primary_channels": band_primary,
        "band_extended_channels": band_extended,
        "theme_filter_applied": 1 if keywords else 0,
        "theme_matched_videos": theme_video_count,
        "theme_matched_channels": len(theme_video_ids),
        "shortlist_before_cap": len(shortlist),
        "shortlist_written": len(written),
        "shortlist_truncated": max(0, len(shortlist) - len(written)),
    }


def _boundaries() -> List[str]:
    return [
        "subscriberCount 隐藏、缺失或非数字写成 UNKNOWN，不写成 0；明确给出的 0 仍是 0。",
        "country=US 和检索 regionCode=US 都不是美国受众。audience_geo_status 保持 UNKNOWN。",
        "公开简介里的 Amazon 普通链接不是平台原生商品标签。amazon_native_tag 保持 UNKNOWN。",
        "邮箱、联系史、退订、退信和品牌身份未经独立来源核验，全部保持 UNKNOWN。",
        "ready_to_send 固定为 no。本命令不生成开发信，也不生成可自动发送的配置。",
        "50K–100K 只标成扩展区间，不和 10K–50K 混写。",
        "频道名或分类词命中品牌措辞，不会自动改写 brand_channel_check。",
        "原视频筛选默认是播放量大于 10000 或互动率大于 3%。未达标的人不会被算成原筛选通过。",
        "人工标签来自调用方提供的复核文件，或 snapshot 里已有的 description_review。它不是新的外部核验。",
        "主题词命中不是内容相关性，也不是推荐理由。",
    ]


def _export_excel(output_dir: Path, videos: Sequence[Dict[str, Any]], shortlist: Sequence[Mapping[str, Any]]) -> str:
    exporter = ExcelExporter(str(output_dir), keyword="snapshot", timestamp="offline")
    export_videos = [_excel_video(video) for video in videos]
    exporter.export_search_videos(export_videos, keyword="")
    exporter.export_influencers([_excel_influencer(row) for row in shortlist])
    return str(exporter.base_dir)


def _excel_video(video: Mapping[str, Any]) -> Dict[str, Any]:
    copied = {
        key: value
        for key, value in video.items()
        if key not in {"description", "search_descriptions"}
    }
    if copied.get("view_count_status") != "known":
        copied["view_count"] = None
        copied["view_count_available"] = False
    if copied.get("like_count_status") != "known":
        copied["like_count"] = None
        copied["like_count_available"] = False
    if copied.get("comment_count_status") != "known":
        copied["comment_count"] = None
        copied["comment_count_available"] = False
    if copied.get("engagement_truth") is None:
        copied["engagement_rate"] = None
    return copied


def _excel_influencer(row: Mapping[str, Any]) -> Dict[str, Any]:
    subscribers: Any
    if row["subscriber_count"] == "UNKNOWN":
        subscribers = "UNKNOWN"
    else:
        subscribers = int(row["subscriber_count"])
    views: Any = row["video_views"]
    return {
        "channel_id": row["channel_id"],
        "channel_title": row["channel_name"],
        "channel_url": row["channel_url"],
        "subscriber_count": subscribers,
        "country": "" if row["creator_country"] == "UNKNOWN" else row["creator_country"],
        "channel_description": "",
        "channel_initial_assessment": row["taxonomy_preliminary"],
        "contact_email": "",
        "email_status": "UNKNOWN",
        "rep_video_url": row["video_url"],
        "rep_video_title": row["video_title"],
        "rep_video_views": None if views == "UNKNOWN" else int(views),
        "rep_video_engagement": None if row["video_engagement_rate"] == "UNKNOWN" else float(row["video_engagement_rate"]),
        "source_keyword": row["search_query"],
        "dev_status": "待核-不可发送",
        "kol_name": "",
    }


def _write_csv(path: Path, rows: Sequence[Mapping[str, Any]]) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=CANDIDATE_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({column: _scrub(row.get(column, "")) for column in CANDIDATE_COLUMNS})


def _write_json(path: Path, payload: Mapping[str, Any]) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _format_command(options: SnapshotOptions) -> str:
    parts = [
        "python main.py snapshot-input",
        "--input",
        str(options.input_path),
        "--output-dir",
        str(options.output_dir),
        "--limit",
        str(options.limit),
    ]
    for path in options.strong_id_paths:
        parts.extend(["--strong-ids", str(path)])
    if options.brand_exclusion_path:
        parts.extend(["--brand-exclusion", str(options.brand_exclusion_path)])
    if options.theme_keywords_path:
        parts.extend(["--theme-keywords-file", str(options.theme_keywords_path)])
    if options.human_review_path:
        parts.extend(["--human-review", str(options.human_review_path)])
    if not options.write_excel:
        parts.append("--no-excel")
    return " ".join(parts)


def _search_preferences(snapshot: Mapping[str, Any]) -> Dict[str, str]:
    prefs = snapshot.get("search_preferences") if isinstance(snapshot.get("search_preferences"), dict) else {}
    return {
        "regionCode": str(prefs.get("regionCode") or ""),
        "relevanceLanguage": str(prefs.get("relevanceLanguage") or ""),
        "note": "search preference only; not creator country and not audience",
    }


def _shape_label(shapes: set) -> str:
    if shapes == {"raw"}:
        return "raw"
    if shapes == {"flat"}:
        return "flat"
    if not shapes:
        return "empty"
    return "mixed"


def _fold(channel_id: str) -> str:
    return channel_id.strip().casefold()


def _is_relative_to(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
    except ValueError:
        return False
    return True


def _scrub(value: Any) -> Any:
    if isinstance(value, str):
        return EMAIL_PATTERN.sub("[redacted-email]", value)
    if isinstance(value, list):
        return [_scrub(item) for item in value]
    if isinstance(value, dict):
        return {key: _scrub(item) for key, item in value.items()}
    return value


