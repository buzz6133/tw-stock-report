# TW Stock Report (GitHub Pages)

這個專案會產生台股投資報告並發佈到 GitHub Pages。

## 使用

1. 更新持股

```
python3 /Users/buzz6/Documents/New project/stock_report.py input
```

2. 產生報表並更新 Pages 內容

```
/Users/buzz6/Documents/New project/publish.sh
```

## Pages 設定

GitHub Repo 設定裡：Settings → Pages → Build and deployment → Source：`Deploy from a branch`，Branch 選 `main`，Folder 選 `docs`。

發佈網址會是：
```
https://<your_github_username>.github.io/<repo>
```
