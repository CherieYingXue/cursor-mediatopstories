# Russia Media Reports

一个独立的手机友好 Web 应用：从俄罗斯主流媒体的**俄语原文** RSS 抓取当天头条，按
**政治 / 经济 / 社会 / 普京 / 梅金斯基**五大板块自动归类，标题翻译成中文，标题本身
可点击打开原新闻页。打开页面自动更新，另有手动更新按钮，抓取失败会降级到静态兜底。

本项目与原来的 `Daily Top Stories Checker`（媒体清单、Excel 导入、Zoology
卡片、IBO 小测验、24 小时时钟等）**已完全分离**，独立部署为 Render 服务
`russia-media-reports`，从分支 `cursor/russia-reports-main-d482` 部署，
永远不会污染原来那套代码。

## 页面与接口

- `GET /` → 302 跳到 `/russia`（可通过环境变量 `LANDING_PATH` 覆盖）
- `GET /russia` → 静态 HTML 页面，加载后 JS 自动调用 `/api/russia-news`
- `GET /api/russia-news` → 并行抓 8 家俄语原文 RSS，归类，翻译。默认 10 分钟内存缓存。
- `GET /api/russia-news?refresh=1` → 强制绕过缓存重新抓取（手动更新按钮就是这么触发的）
- `GET /healthz` → 供 Render / uptime 用的健康检查

抓取的俄媒（全部为俄语原文站，不使用任何俄媒英文分站）：

- ТАСС · РИА Новости · РТ на русском · Lenta.ru · Российская газета
- Коммерсантъ · Известия · Kremlin.ru

翻译走谷歌无鉴权端点 `translate.googleapis.com/translate_a/single?client=gtx`。
翻译失败时前端会退化为只显示俄语原标题 + 提示"点击标题查看俄语原文"。

## 本地跑

```
pip install -r requirements.txt
python app.py
open http://127.0.0.1:5000/
```

## Render 部署

`render.yaml` 已配置好：

- 服务名：`russia-media-reports`
- 分支：`cursor/russia-reports-main-d482`（不会碰到 `main`）
- 环境变量：`SECRET_KEY`（自动生成）、`LANDING_PATH=/russia`、`RUSSIA_CACHE_TTL=600`
- 健康检查：`/healthz`
- 启动：`gunicorn -w 2 -b 0.0.0.0:$PORT app:app`

在 [dashboard.render.com](https://dashboard.render.com) → New → Blueprint →
选仓库 → 在 branch 下拉里选 `cursor/russia-reports-main-d482` → Apply 即可。

如果已经从旧的 Feature 分支创建过 `russia-media-reports` 服务，只需在 Render 该服务的
Settings → Branch 里把分支改成 `cursor/russia-reports-main-d482`，然后手动 Deploy
一次即可。
