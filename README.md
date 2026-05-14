# Daily Top Stories Checker

This app checks top stories for a list of news domains and stores the latest results.

## What it does

- Optional **media catalog** from `top story checker media list.xlsx` in the app folder: column **A = country**, **B = media name**, **C = URL**; the list and results show country and name; top story still uses Google News RSS by hostname from the URL.
- Fetches each domain's top story via Google News RSS (`site:domain`).
- Saves each run to a SQLite database.
- Exports latest run to `exports/top_stories_*.csv`.
- Saves your custom domain list for reuse.
- Supports domain list import/export and **merging domains from an Excel file** (`.xlsx` / `.xlsm`).
- Runs automatic daily checks on a schedule.
- Shows a mobile-friendly web page.
- Supports "Add to Home Screen" as a basic PWA.

## Run locally

1. Install Python 3.10+.
2. Open terminal in this folder.
3. Install dependencies:
   - `python -m pip install -r requirements.txt`
4. Run:
   - `python app.py`
5. Open:
   - `http://127.0.0.1:5000`

## One-click first run (recommended on Windows)

- Double-click:
  - `start_app_first_run.bat`

What it does:

- Creates `.venv` if needed.
- Installs dependencies.
- Imports `domains_full.txt` into app settings on first run only.
- Starts the app.

## Daily schedule

- Each manual run checks **at most 10** domains you select with checkboxes. That selection is **remembered** for the next page load and for the **daily scheduled** run.
- The scheduled job runs **only** those remembered domains (still present in your saved media list). If you have not run “Check Top Stories Now” yet, the daily job does nothing until you do.
- Default run time is `08:00` (server local time).
- Change it with environment variables:
  - `DAILY_RUN_HOUR` (0-23)
  - `DAILY_RUN_MINUTE` (0-59)

Example:

- `DAILY_RUN_HOUR=6`
- `DAILY_RUN_MINUTE=30`

## Import/export domain list

- In UI:
  - Use **Save Domain List** to save the full list from the textarea (without running a check).
  - Use **Add media from Excel** to upload `.xlsx` or `.xlsm`: the first worksheet is scanned; URLs and domains are merged into the saved list (duplicates skipped).
- API/helper endpoint:
  - `GET /export-domains` writes `exports/domains_latest.txt`

## Build Windows EXE

Run:

- `build_exe.bat`

After build, use:

- `dist/topstories.exe`

## Deploy to GitHub + Render (公网 / 手机访问)

我无法替你登录 GitHub 或 Render；按下面做即可（约 10 分钟）。

### 1. 在 GitHub 新建空仓库

1. 打开 [github.com/new](https://github.com/new)
2. Repository name 自定（例如 `daily-top-stories-checker`）
3. **不要**勾选 “Add a README”
4. 创建后复制仓库地址，例如 `https://github.com/你的用户名/daily-top-stories-checker.git`

### 2. 在本机推送代码

在项目目录 `C:\Users\xhs\Desktop\cursor` 打开 **Git Bash** 或 **PowerShell**，执行（把 URL 换成你的）：

```bash
git remote add origin https://github.com/你的用户名/仓库名.git
git push -u origin main
```

若提示登录：用 GitHub 的 **Personal Access Token** 作为密码（不要用账号密码）。  
创建 Token：[github.com/settings/tokens](https://github.com/settings/tokens) → Generate new token → 勾选 `repo`。

### 3. 在 Render 部署

1. 打开 [render.com](https://render.com) 并登录（可用 GitHub 登录）
2. **New** → **Blueprint**（或 **Web Service** 并连接同一仓库）
3. 选择刚推送的仓库；若使用 Blueprint，会读取本仓库的 `render.yaml`
4. 部署完成后会得到 `https://xxx.onrender.com`
5. 手机浏览器打开该地址 → 菜单里 **添加到主屏幕**（PWA）

### 4. 环境变量（可选）

在 Render 面板为该服务设置：

- `DAILY_RUN_HOUR`（默认 `8`）
- `DAILY_RUN_MINUTE`（默认 `0`）

### 说明（Render 免费实例）

- 免费 Web 服务在无访问时可能休眠，首次打开会慢几秒。
- 磁盘非持久化时，SQLite 数据在重装/休眠后可能丢失；域名列表可重新用页面或 `domains_full.txt` 导入。

本仓库已包含 `render.yaml` 与 `Procfile`，按上面推送后即可被 Render 识别。

## Important note about phone install

- A Windows `.exe` does **not** install on a phone browser.
- For phone use, run this as a hosted web app (PWA) and open that URL on your phone.
- Then tap browser menu -> **Add to Home Screen**.
