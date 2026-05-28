# 采购审批系统

一个可直接部署到 Debian VPS 的轻量采购申请与审批系统，使用 Flask、SQLite 和服务端图片压缩。

## 功能

- 用户提交采购申请：申请日期、采购内容、采购图片、发票。
- 图片会在服务端压缩为 JPEG，并限制到最大 500KB。
- 发票支持图片或 PDF；图片发票同样压缩到 500KB，PDF 最大 10MB。
- 审批人可将申请标记为“已同意”或“未同意”，并填写审批备注。
- `user1`、`user2` 和审批人都可以查看全部申请及状态：待审批、已同意、未同意。
- SQLite 数据和上传文件存放在 `instance/`，不会提交到 Git。
- 可用 `DATABASE_PATH` 自定义 SQLite 文件路径，`UPLOAD_FOLDER` 自定义附件目录。

## 默认账号

首次运行会自动创建 3 个账号。生产环境请通过 `.env` 立即改密码。

| 用户名 | 默认密码 | 角色 |
| --- | --- | --- |
| `approver` | `approver123` | 审批人 |
| `user1` | `user1123` | 申请用户 |
| `user2` | `user2123` | 申请用户 |

## 本地运行

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
flask --app app run
```

浏览器打开 `http://127.0.0.1:5000`。

主流程冒烟测试：

```bash
.venv/bin/python scripts/smoke_test.py
```

## Docker 部署

```bash
cp .env.example .env
docker compose up -d --build
```

默认监听 `http://服务器IP:8000`。数据保存在 Docker volume `procurement_data`。

## Debian VPS 部署

```bash
sudo apt update
sudo apt install -y git python3-venv python3-pip nginx
sudo git clone https://github.com/你的用户名/你的仓库名.git /opt/procurement-approval
cd /opt/procurement-approval
sudo python3 -m venv .venv
sudo .venv/bin/pip install -r requirements.txt
sudo cp .env.example .env
sudo nano .env
sudo mkdir -p instance/uploads
sudo chown -R www-data:www-data /opt/procurement-approval
sudo cp deploy/procurement-approval.service /etc/systemd/system/
sudo systemctl daemon-reload
sudo systemctl enable --now procurement-approval
```

可用 Nginx 反向代理到 `127.0.0.1:8000`，再用 Certbot 配置 HTTPS。

## 同步到 GitHub

如果已经创建好 GitHub 仓库：

```bash
git remote add origin git@github.com:你的用户名/你的仓库名.git
git push -u origin main
```

如果安装并登录了 GitHub CLI：

```bash
gh repo create 你的仓库名 --private --source=. --remote=origin --push
```

之后在任何 Debian VPS 上执行 `git clone` 即可部署。
