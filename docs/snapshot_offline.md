# 离线 snapshot

`snapshot-input` 读取已经采集好的 YouTube Data API JSON，在本机完成解析、视频筛选、频道查重和导出。它不调用 YouTube，不读取 API Key，不写飞书，不调用 17TRACK，也不生成或发送开发信。

这不是端到端在线自动化通过。在线搜索、翻页、最近视频、飞书授权和物流刷新仍走原来的命令，那些命令这次没有被这个入口替代。

## 输入

支持两种 item 形状，可以混在同一次响应里：

- 原始 API item：视频和频道使用 `snippet` / `statistics` / `contentDetails`，搜索 item 使用 `id.videoId`。
- 已经铺平的 item：视频使用 `videoId`、`viewCount`、`duration`，频道使用 `channelId`、`subscriberCount`、`hiddenSubscribers`。

`subscriberCount` 在隐藏、缺失或不是数字时写成 `UNKNOWN`，不写成 0。接口明确给出的 0 仍保留为 0。

## 强频道 ID

`--strong-ids` 可以重复。CSV 有 `key_type` 和 `key_value` 时，只采用 `channel_id` 与 `channel_id_lower`，并且 `evidence_strength` 必须为空或 `strong`。姓名、handle 和邮箱键不会参与匹配。另一种 CSV 只读取显式的 `channel_id` 列，不读取姓名列。

没有命中只说明该 ID 不在所给文件中，不能当成从未联系。

## 审计和复核短名单

订阅区间内、且没有被强频道 ID 或品牌频道 ID 排除的行写入 `audit.csv`。这不是推荐名单。

`--human-review` 是 CSV，列至少为 `channel_id,decision`，可选 `reason,source`。`decision` 只认 `kept` 和 `dropped`。其他值，包括空值和 unknown，都不会进入复核短名单。

只有 `kept` 写入 `review-shortlist.csv`。`dropped` 即使命中主题词和订阅区间，也留在 audit，队列是 `human_dropped_not_recommended`。没有人工标签的人留在 audit，队列是 `pending_identity_check_not_recommended`。强频道 ID 和品牌频道 ID 优先于人工 kept。频道名本身不能判成官方，也不能因此排除评测作者。

人工 kept 但没有通过原视频阈值时，保留原筛选的 fail，并写上 `human_retained_exception_auto_threshold_not_passed`。这不是把阈值放宽后的通过。

主题词命中不是相关性，也不是推荐。`ready_to_send` 固定为 `no`。身份 UNKNOWN 不是验证通过。

输出目录必须在 git 工作树之外。真实候选、排除表、真实品牌配置和运行结果不要提交。下面的配置文件只含合成 ID。

## 合成样例

```bash
python main.py snapshot-input \
  --input tests/fixtures/snapshot_offline_synthetic.json \
  --strong-ids tests/fixtures/snapshot_offline_strong_ids.csv \
  --brand-exclusion tests/fixtures/snapshot_offline_brands.json \
  --theme-keywords-file tests/fixtures/snapshot_offline_theme.txt \
  --human-review tests/fixtures/snapshot_offline_human_review.csv \
  --output-dir /tmp/snapshot-offline-demo
```

样例输出是 `audit.csv`、`review-shortlist.csv`、`summary.json` 和 `report.md`。样例里的频道 ID、邮箱和链接都是合成的。
