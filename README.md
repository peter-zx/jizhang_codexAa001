# 分销记账工具

面向“总舵主 + 分销商”的月度对账系统。总舵主只有一个，可以查看全部分销商数据；分销商只能查看和维护自己的人员数据。系统支持注册登录、人员录入、Excel 上传导入、当月统计、唯一文件名导出。

## 环境隔离

项目内置三层隔离：

- Python 依赖隔离：使用项目根目录 `.venv/`，不依赖系统 Python 包。
- 配置隔离：`env/development.env`、`env/production.env.example` 分开维护，通过 `scripts/use-env.ps1` 激活到 `.env`。
- 数据隔离：开发环境默认使用 `data/dev.db`，生产环境默认使用 `data/prod.db`；上传、导出、日志都在项目目录下。

激活开发环境：

```powershell
.\scripts\use-env.ps1 -Environment development
```

生产部署时复制并修改：

```powershell
Copy-Item env\production.env.example env\production.env
.\scripts\use-env.ps1 -Environment production
```

## 已落实的工程原则

- 不写死本地绝对路径：所有路径从项目根目录和 `.env` 配置解析。
- 开发/生产环境分离：通过 `env/*.env` 和 `.env` 激活机制隔离。
- 用户文件只通过上传接口进入：`/imports` 接收 `.xlsx` 后保存到 `data/uploads`。
- 输出文件名唯一：上传和导出统一使用 UUID 文件名，避免并发覆盖。
- 分层清晰：接口层 `app/main.py`，业务层 `app/services`，存储层 `app/services/storage.py`，模型层 `app/models.py`。
- I/O 异常处理和日志：导入、导出、上传、数据库写入均有异常处理和 `logs/app.log`。
- 数据流清晰：上传/录入 -> 校验 -> 存储 -> 月度统计 -> 导出。
- 多用户隔离：分销商按 `owner_id` 隔离；总舵主角色可聚合查看。
- 异步扩展点：Excel 导入导出服务已独立，后续可替换为 Celery/RQ/后台任务。
- 服务器部署结构：提供 `Dockerfile`、`docker-compose.yml`、`.env.example`。

## 目录结构

```text
.
├─ app/
├─ data/
│  ├─ uploads/
│  └─ exports/
├─ env/
│  ├─ development.env
│  └─ production.env.example
├─ logs/
├─ scripts/
│  ├─ use-env.ps1
│  └─ run-dev.ps1
├─ tests/
├─ requirements.txt
├─ Dockerfile
└─ docker-compose.yml
```

## 本地运行

```powershell
cd C:\Users\admin\Desktop\0501
.\scripts\use-env.ps1 -Environment development
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\scripts\run-dev.ps1
```

打开：

```text
http://127.0.0.1:8000
```

首次启动会自动创建唯一总舵主账号。分销商通过注册页面创建，角色固定为分销商，不能自行注册为总舵主。

## Docker 部署

```powershell
Copy-Item env\production.env.example .env
# 修改 .env 后：
docker compose up -d --build
```

生产建议：

- `APP_ENV=production`
- `SECRET_KEY` 使用强随机值
- `GRANDMASTER_PASSWORD` 使用强密码
- SQLite 可用于轻量单机；多人并发更高时改成 PostgreSQL，并设置 `DATABASE_URL`
- `data/` 和 `logs/` 做持久化备份
- 放到 Nginx/Caddy 后面时启用 HTTPS

## Excel 字段

当前兼容样表表头：

```text
序号, 姓名, 在职地区, 安置时间, 工资卡, 服务费, 发薪类型, 应发, 应返, 结算所属期, 在职离职, 备注, 渠道, 残疾类型1, 残疾等级1, 残疾类型2, 残疾等级2, 入职时间, 年龄, 性别
```

`结算所属期` 会标准化为 `YYYY-MM`，例如 `2026 年 1 月` 会保存为 `2026-01`。

## 后续可扩展项

- 把导入/导出接入 Celery 或 RQ，适配大文件后台处理。
- 增加人员编辑、删除、停用审计。删除属于敏感操作，建议做软删除并保留日志。
- 增加总舵主审批分销商注册。
- 增加月度锁账，防止对账完成后被修改。
- 增加 PostgreSQL、对象存储、审计日志和权限细粒度策略。
