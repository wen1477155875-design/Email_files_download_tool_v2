# 邮件附件下载工具 - 公司电脑部署说明

## 文件清单

| 文件 | 作用 |
| --- | --- |
| `安装依赖.bat` | **第一次先运行这个**：自动创建 .venv、装依赖、跑自检 |
| `启动界面.bat` | 日常入口：打开图形界面（不弹黑窗口） |
| `ui.py` | 图形界面（tkinter，Python 自带） |
| `main.py` | 命令行入口（check / run / task / state） |
| `pipeline.py` | 主流程：扫描 → 过滤 → 去重 → 下载 → 解压 |
| `outlook_client.py` | Outlook COM 封装（必须用经典桌面版 Outlook） |
| `scheduler.py` | Windows 计划任务注册（普通用户即可，免管理员） |
| `extractor.py` | ZIP 解压（防 zip slip / zip bomb / 中文名乱码） |
| `state_store.py` | 去重状态（下载过的不再重复下载） |
| `config_loader.py` / `utils.py` | 配置解析 / 工具函数 |
| `smoke_test.py` | 离线自检（假数据，不连 Outlook） |
| `config.yaml` | 唯一需要改的配置文件（也可全在界面里改） |
| `requirements.txt` | 第三方依赖清单（pywin32、PyYAML） |

## 部署步骤

1. **装 Python**（如果公司电脑没有）：python.org 下载 3.10+，
   安装时勾选 **tcl/tk**（界面需要）和 **Add Python to PATH**。
2. **确认经典版 Outlook 已配置邮箱账号**：
   必须是桌面版 Outlook（Office 自带的那个），不能是 Win11 自带的「Outlook (new)」。
3. **双击 `安装依赖.bat`**：脚本会自动检测电脑上"能实际运行"的 Python（PATH 里的 py/python、常见安装路径逐个实测，自动跳过被公司策略拦截的），然后把依赖用 `pip --user` 装到你的用户目录——**不需要管理员权限，也不创建 .venv**。
   脚本会生成 `python_env.txt` 记录找到的 Python 路径，`启动界面.bat` 靠它启动界面。
4. **双击 `启动界面.bat`**，在界面里填写：
   - 邮箱（公司邮箱显示名或地址）、文件夹（留空 = 默认收件箱）
   - 关注发件人（每行一个）
   - 下载路径、解压路径（建议放 D 盘自己的目录）
   - 定时方式（每天固定时间 / 每隔 N 分钟）→ 点「启动定时」
5. 先勾 **dry-run（只预览不下载）** 点「立即运行」验证一遍，
   确认无误后去掉勾选正式跑。

## 注意事项

- **计划任务只在"你登录 Windows"时触发**（COM 方案的限制），注销状态不会跑。
- 界面关掉后定时仍然生效（任务归 Windows 任务计划程序管，可在任务计划程序里看到
  `EmailAttachmentDownloader`）。
- 依赖装在用户目录（`%APPDATA%\Python`），卸载命令：`python -m pip uninstall -y pywin32 pyyaml`。
- 若机器上有旧的 `.venv` 文件夹可以删掉，新版脚本不再使用它。
- `data/state.json` 是去重记录：删掉它会导致历史邮件重新下载一遍。
- 如果公司网络限制 pip，把 `安装依赖.bat` 里的清华镜像地址换成公司内网源，
  或在有网的机器下载 wheel 包拷过去离线安装。

## 常见问题

| 现象 | 处理 |
| --- | --- |
| `No module named tkinter` | Python 安装时没勾 tcl/tk，重装勾上 |
| 连不上 Outlook / Folders 报错 | 经典版 Outlook 没配账号，先打开 Outlook 登录一次 |
| 邮箱名找不到 | 界面"邮箱"留空 = 用 Outlook 默认邮箱 |
| 下载 0 但"已存在 N" | 之前下载过被去重；文件若被移走会自动补下（新版行为） |
| 定时没触发 | 任务计划程序里查 `EmailAttachmentDownloader` 上次运行结果 |
| 提示找不到 pywin32 / 安装依赖报错 | 公司网络连不上 pip 源。脚本会自动轮询默认源/清华/阿里/腾讯镜像，全部失败时改用包内 `wheels/` 目录离线安装（覆盖 Python 3.10~3.13，64 位）。仍失败请把完整报错和 [1/5] 显示的 Python 版本发回来 |
| `This program is blocked by group policy` | Python 启动被公司策略拦截。脚本会自动改试其他 Python；若全部被拦，需 IT 放行 python.exe 或提供公司认可的 Python。可先在 cmd 里跑 `python --version` 确认哪个能用 |
