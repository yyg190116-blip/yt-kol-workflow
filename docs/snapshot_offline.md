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

## 短名单

默认保留订阅数 10K–50K，并把 50K–100K 标成 `50K-100K_extended`。隐藏或缺失订阅数不进入这两个区间。`--limit` 是 1 到 10，默认 10；不足不会补人。

`country=US` 和搜索 `regionCode=US` 都不写成美国受众。Amazon 普通链接不写成原生商品标签。邮箱、联系史、退订、退信和品牌身份未核验时保持 `UNKNOWN`。`ready_to_send` 固定为 `no`。

主题词只匹配标题和简介。搜索词本身不会让一条不相关的视频变成主题命中。

输出目录必须在 git 工作树之外。真实候选、排除表和运行结果不要提交。

## 合成样例

```bash
python main.py snapshot-input \
  --input tests/fixtures/snapshot_offline_synthetic.json \
  --strong-ids tests/fixtures/snapshot_offline_strong_ids.csv \
  --brand-exclusion tests/fixtures/snapshot_offline_brands.json \
  --theme-keywords-file tests/fixtures/snapshot_offline_theme.txt \
  --output-dir /tmp/snapshot-offline-demo
```

样例里的频道 ID、邮箱和链接都是合成的。
