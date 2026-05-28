# 采购审批系统

一个可直接部署到 Debian VPS 的轻量采购申请与审批系统，使用 Flask、SQLite 和服务端图片压缩。

## 功能

- 用户提交采购申请：申请日期、采购内容、多行采购明细、采购图片、发票。
- 采购明细支持单价、数量和行总价，页面会自动计算申请总价。
- 图片会在服务端压缩为 JPEG，并限制到最大 500KB。
- 发票支持图片或 PDF；图片发票同样压缩到 500KB，PDF 最大 10MB。
- `jose` 和 `admin` 可将申请标记为“已同意”或“未同意”，并填写审批备注。
- `admin` 可新增用户、修改现有用户密码、按日期/状态筛选并导出 Excel。
- UI 支持中文和西班牙语，登录页可选择语言。
- 所有登录用户都可以查看全部申请及状态：待审批、已同意、未同意。
- SQLite 数据和上传文件存放在 `instance/`，不会提交到 Git。
- 可用 `DATABASE_PATH` 自定义 SQLite 文件路径，`UPLOAD_FOLDER` 自定义附件目录。

## 默认账号

首次运行会自动创建 3 个账号。账号密码写在应用配置里；上线后可用 `admin` 登录到用户管理里修改密码。

| 用户名 | 默认密码 | 角色 |
| --- | --- | --- |
| `jose` | `jose1234` | 审批人 |
| `admin` | `admin1234` | 管理员、可审批、可导出 Excel |
| `carlos` | `carlos1234` | 申请人 |

`SECRET_KEY` 仍建议在 `.env` 中设置为随机值，用于保护登录会话。

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
docker network create web
docker compose up -d --build
```

容器使用哥伦比亚时区 `America/Bogota`。Compose 不直接暴露宿主机端口，只加入外部桥接网络 `web`，用于给 Nginx Proxy Manager 反向代理。数据保存在 Docker volume `procurement_data`。

已部署过旧版本时，拉取代码后直接重建即可，数据库会自动迁移并新增 `jose`、`admin`、`carlos`：

```bash
git pull
docker network create web 2>/dev/null || true
docker compose up -d --build
```

## Nginx Proxy Manager

域名 `buy.dlxy.top` 需要先在 DNS 中添加 A 记录，指向 VPS 公网 IP。

Nginx Proxy Manager 容器也必须加入同一个 Docker 网络：

```bash
docker network connect web nginx-proxy-manager容器名
```

在 Nginx Proxy Manager 中新增 Proxy Host：

- Domain Names: `buy.dlxy.top`
- Scheme: `http`
- Forward Hostname / IP: `procurement-approval`
- Forward Port: `8000`
- 勾选 Block Common Exploits
- SSL 页签选择 Request a new SSL Certificate
- 勾选 Force SSL 和 HTTP/2 Support

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
