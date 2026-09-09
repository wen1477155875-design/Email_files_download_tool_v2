"""Windows 计划任务注册。

首选任务计划程序的 COM 接口（Schedule.Service），失败时回退到 schtasks 命令行。

为什么优先用 COM：
  1. 能设置 WorkingDirectory —— schtasks 的 /TR 不支持工作目录，路径解析极易出错。
  2. 参数不需要二次转义（中文路径、含空格路径都安全）。
  3. 能精确控制"错过后补跑""实例并发""执行超时"等行为。
  4. 不依赖 schtasks.exe 是否可用（某些受管控环境会禁用该命令）。

权限说明：
  Outlook COM 必须在已登录的会话中运行，所以统一使用
  LogonType=3（交互式令牌，仅当用户登录时运行），不保存密码、不需要管理员权限。
"""

from __future__ import annotations

import locale
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from config_loader import Config

# Task Scheduler 常量
TASK_TRIGGER_TIME = 1
TASK_TRIGGER_DAILY = 2
TASK_ACTION_EXEC = 0
TASK_CREATE_OR_UPDATE = 6
TASK_LOGON_INTERACTIVE_TOKEN = 3
TASK_INSTANCES_IGNORE_NEW = 2

TASK_XML_TEMPLATE = """<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Author>EmailAttachmentDownloader</Author>
    <Description>{description}</Description>
  </RegistrationInfo>
  <Triggers>
{trigger}  </Triggers>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <StartWhenAvailable>{start_when_available}</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>{execution_time_limit}</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>{command}</Command>
      <Arguments>{arguments}</Arguments>
      <WorkingDirectory>{working_directory}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"""

CALENDAR_TRIGGER = """    <CalendarTrigger>
      <StartBoundary>{start}</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByDay>
        <DaysInterval>1</DaysInterval>
      </ScheduleByDay>
    </CalendarTrigger>
"""

TIME_TRIGGER = """    <TimeTrigger>
      <StartBoundary>{start}</StartBoundary>
      <Enabled>true</Enabled>
      <Repetition>
        <Interval>PT{minutes}M</Interval>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
    </TimeTrigger>
"""


class SchedulerError(RuntimeError):
    pass


def _decode(data: bytes) -> str:
    for encoding in ("utf-8", "gbk", locale.getpreferredencoding(False) or "utf-8"):
        try:
            return data.decode(encoding)
        except Exception:
            continue
    return data.decode("utf-8", "replace")


def _run(cmd: List[str]) -> subprocess.CompletedProcess:
    """执行外部命令。schtasks 可能被安全策略禁用，这里统一转成友好错误。"""
    try:
        return subprocess.run(cmd, capture_output=True)
    except (PermissionError, FileNotFoundError) as exc:
        raise SchedulerError(
            f"无法执行 {cmd[0]}（{exc}）。"
            "该命令可能被安全策略禁用，请改用 COM 方式或联系管理员。"
        ) from exc


def _parse_time(value: str) -> str:
    text = (value or "09:00").strip()
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).strftime("%H:%M:%S")
        except ValueError:
            continue
    raise SchedulerError(f"schedule.time 格式应为 HH:MM，当前是 {value!r}")


def _com_available() -> bool:
    try:
        import win32com.client  # noqa: F401
        return True
    except Exception:
        return False


def _service():
    import win32com.client
    service = win32com.client.Dispatch("Schedule.Service")
    service.Connect()
    return service


# --------------------------------------------------------------------------
# COM 方式
# --------------------------------------------------------------------------

def _build_definition(service, cfg: Config, python_exe: Path, script: Path):
    definition = service.NewTask(0)

    definition.RegistrationInfo.Author = "EmailAttachmentDownloader"
    definition.RegistrationInfo.Description = "定时下载 Outlook 指定发件人的邮件附件"

    if cfg.schedule.mode == "daily":
        trigger = definition.Triggers.Create(TASK_TRIGGER_DAILY)
        trigger.StartBoundary = f"{datetime.now():%Y-%m-%d}T{_parse_time(cfg.schedule.time)}"
        trigger.DaysInterval = 1
    else:
        minutes = max(1, int(cfg.schedule.interval_minutes or 30))
        trigger = definition.Triggers.Create(TASK_TRIGGER_TIME)
        trigger.StartBoundary = f"{datetime.now():%Y-%m-%dT%H:%M:%S}"
        trigger.Repetition.Interval = f"PT{minutes}M"
        trigger.Repetition.Duration = "P1D"      # 必须大于 Interval，否则注册失败
    trigger.Enabled = True

    action = definition.Actions.Create(TASK_ACTION_EXEC)
    action.Path = str(python_exe)
    action.Arguments = f'"{script}" run'
    action.WorkingDirectory = str(cfg.base_dir)

    settings = definition.Settings
    settings.Enabled = True
    settings.Hidden = False
    settings.StartWhenAvailable = bool(cfg.schedule.start_when_available)
    # 注意：COM 属性名是 MultipleInstances，XML 里才叫 MultipleInstancesPolicy
    settings.MultipleInstances = TASK_INSTANCES_IGNORE_NEW
    settings.DisallowStartIfOnBatteries = False
    settings.StopIfGoingOnBatteries = False
    settings.ExecutionTimeLimit = cfg.schedule.execution_time_limit or "PT30M"
    settings.AllowDemandStart = True

    return definition


def _install_com(cfg: Config, python_exe: Path, script: Path) -> None:
    service = _service()
    folder = service.GetFolder("\\")
    definition = _build_definition(service, cfg, python_exe, script)
    folder.RegisterTaskDefinition(
        cfg.schedule.task_name,
        definition,
        TASK_CREATE_OR_UPDATE,
        "",                                # 当前用户
        "",                                # 不保存密码
        TASK_LOGON_INTERACTIVE_TOKEN,      # 仅当用户登录时运行
    )


def _uninstall_com(task_name: str) -> bool:
    service = _service()
    service.GetFolder("\\").DeleteTask(task_name, 0)
    return True


def _status_com(task_name: str) -> Optional[str]:
    service = _service()
    try:
        task = service.GetFolder("\\").GetTask(task_name)
    except Exception:
        return None
    result = int(task.LastTaskResult)
    meaning = {
        0: "成功",
        1: "调用被函数拒绝",
        0x41300: "任务已就绪",
        0x41301: "任务正在运行",
        0x41302: "任务已禁用",
        0x41303: "任务尚未运行",
        0x41306: "任务已终止",
    }.get(result, f"0x{result:X}")
    return (
        f"任务名称  : {task.Name}\n"
        f"启用状态  : {task.Enabled}\n"
        f"当前状态  : {task.State} (3=就绪/触发中, 4=已禁用)\n"
        f"下次运行  : {task.NextRunTime}\n"
        f"上次运行  : {task.LastRunTime}\n"
        f"上次结果  : {result} ({meaning})"
    )


def _run_now_com(task_name: str) -> bool:
    service = _service()
    service.GetFolder("\\").GetTask(task_name).Run("")
    return True


# --------------------------------------------------------------------------
# 对外接口
# --------------------------------------------------------------------------

def build_task_xml(cfg: Config, python_exe: Path, script: Path) -> str:
    """生成任务 XML（仅 schtasks 回退路径使用）。"""
    if cfg.schedule.mode == "daily":
        start = f"{datetime.now():%Y-%m-%d}T{_parse_time(cfg.schedule.time)}"
        trigger = CALENDAR_TRIGGER.format(start=start)
    else:
        minutes = max(1, int(cfg.schedule.interval_minutes or 30))
        start = f"{datetime.now():%Y-%m-%dT%H:%M:%S}"
        trigger = TIME_TRIGGER.format(start=start, minutes=minutes)

    return TASK_XML_TEMPLATE.format(
        description="定时下载 Outlook 指定发件人的邮件附件",
        trigger=trigger,
        start_when_available="true" if cfg.schedule.start_when_available else "false",
        execution_time_limit=cfg.schedule.execution_time_limit,
        command=str(python_exe),
        arguments=f'"{script}" run',
        working_directory=str(cfg.base_dir),
    )


def install_task(cfg: Config, python_exe: Optional[Path] = None, script: Optional[Path] = None, logger=None) -> None:
    python_exe = Path(python_exe) if python_exe else Path(_current_python())
    script = Path(script) if script else Path(__file__).resolve().parent / "main.py"
    task_name = cfg.schedule.task_name

    if _com_available():
        try:
            _install_com(cfg, python_exe, script)
            _log(logger, f"计划任务已注册：{task_name}（COM 方式）")
            _log(logger, f"  触发方式：{cfg.schedule.mode}"
                         + (f" {cfg.schedule.time}" if cfg.schedule.mode == "daily"
                            else f" 每 {cfg.schedule.interval_minutes} 分钟"))
            _log(logger, f"  解释器  ：{python_exe}")
            _log(logger, f"  脚本    ：{script} run")
            _log(logger, f"  工作目录：{cfg.base_dir}")
            _log(logger, "  登录类型：仅当用户登录时运行（Outlook COM 的硬性要求）")
            return
        except Exception as exc:
            _log(logger, f"COM 方式注册失败，回退 schtasks：{exc}")

    # ---- 回退：schtasks + XML ----
    xml_text = build_task_xml(cfg, python_exe, script)
    xml_path = Path(tempfile.gettempdir()) / f"{task_name.replace(' ', '_')}.xml"
    xml_path.write_text(xml_text, encoding="utf-16")

    result = _run(["schtasks", "/Create", "/TN", task_name, "/XML", str(xml_path), "/F"])
    if result.returncode == 0:
        _log(logger, f"计划任务已注册：{task_name}（schtasks 方式）")
        _log(logger, f"  解释器：{python_exe}")
        _log(logger, f"  工作目录：{cfg.base_dir}")
        return

    _log(logger, f"XML 方式注册失败，回退命令行方式。输出：{_decode(result.stdout).strip()}")
    task_run = f'"{python_exe}" "{script}" run'
    if cfg.schedule.mode == "daily":
        cmd = ["schtasks", "/Create", "/TN", task_name, "/TR", task_run,
               "/SC", "DAILY", "/ST", _parse_time(cfg.schedule.time)[:5], "/F"]
    else:
        cmd = ["schtasks", "/Create", "/TN", task_name, "/TR", task_run,
               "/SC", "MINUTE", "/MO", str(max(1, int(cfg.schedule.interval_minutes or 30))), "/F"]
    fallback = _run(cmd)
    if fallback.returncode != 0:
        raise SchedulerError(
            "注册计划任务失败：\n"
            f"{_decode(fallback.stdout).strip()}\n{_decode(fallback.stderr).strip()}"
        )
    _log(logger, f"计划任务已注册（命令行方式）：{task_name}")


def uninstall_task(task_name: str, logger=None) -> bool:
    if _com_available():
        try:
            _uninstall_com(task_name)
            _log(logger, f"计划任务已删除：{task_name}")
            return True
        except Exception as exc:
            _log(logger, f"COM 删除失败（可能不存在），尝试 schtasks：{exc}")

    result = _run(["schtasks", "/Delete", "/TN", task_name, "/F"])
    if result.returncode == 0:
        _log(logger, f"计划任务已删除：{task_name}")
        return True
    message = _decode(result.stdout).strip() or _decode(result.stderr).strip()
    _log(logger, f"删除计划任务失败（可能不存在）：{message}")
    return False


def query_status(task_name: str) -> Optional[str]:
    """查询任务状态，返回可展示的文本；任务不存在返回 None。供 UI 使用。"""
    if _com_available():
        info = _status_com(task_name)
        if info:
            return info

    result = _run(["schtasks", "/Query", "/TN", task_name, "/V", "/FO", "LIST"])
    if result.returncode != 0:
        return None
    return _decode(result.stdout).strip()


def task_status(task_name: str, logger=None) -> bool:
    if _com_available():
        info = _status_com(task_name)
        if info:
            _log(logger, info)
            return True

    result = _run(["schtasks", "/Query", "/TN", task_name, "/V", "/FO", "LIST"])
    if result.returncode != 0:
        _log(logger, f"计划任务不存在：{task_name}")
        return False
    _log(logger, _decode(result.stdout).strip())
    return True


def run_task_now(task_name: str, logger=None) -> bool:
    if _com_available():
        try:
            _run_now_com(task_name)
            _log(logger, f"已手动触发：{task_name}")
            return True
        except Exception as exc:
            _log(logger, f"COM 触发失败，尝试 schtasks：{exc}")

    result = _run(["schtasks", "/Run", "/TN", task_name])
    if result.returncode != 0:
        _log(logger, f"手动触发失败：{_decode(result.stdout).strip()}")
        return False
    _log(logger, f"已手动触发：{task_name}")
    return True


def _current_python() -> str:
    import sys
    return sys.executable


def _log(logger, message: str) -> None:
    if logger:
        logger.info(message)
    else:
        print(message)
