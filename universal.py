import win32com.client
import pythoncom


class Universal:
    # 按优先级排列：Microsoft Word 优先，然后 WPS Writer
    _WORD_PROG_IDS = [
        "Word.Application",       # Microsoft Word
        "WPS.Application",         # WPS Writer (新版)
        "wps.Application",         # WPS Writer (小写变体)
        "Kwps.Application",        # WPS Writer (旧版 Kingsoft)
        "KWPS.Application",        # WPS Writer (旧版变体)
    ]

    def __init__(self):
        self.active_word_app = None
        self._disposed = False
        self.last_error = ""

    @property
    def active_word_app(self):
        return self._active_word_app

    @active_word_app.setter
    def active_word_app(self, value):
        self._active_word_app = value

    def get_active_word_app(self, is_create_word_app=False):
        """获取正在运行的 Word 或 WPS 文字处理应用实例。

        先尝试连接已运行的实例（按 _WORD_PROG_IDS 顺序），
        如果需要且未找到运行实例，则尝试创建新实例。

        连不上时，每个 ProgID 的失败原因会汇总进 self.last_error ——
        界面要靠它告诉用户真正的原因，而不是笼统地报一句
        「请先打开一个文档」。程序运行在云同步目录里时同步客户端会把
        进程关进沙箱，沙箱内看不到桌面会话里的 Word/WPS，
        self.last_error 就是唯一能区分这种情况的证据。
        """
        self.last_error = ""
        reasons = []
        try:
            pythoncom.CoInitialize()

            # 先尝试获取已运行的实例
            for prog_id in self._WORD_PROG_IDS:
                try:
                    self._active_word_app = win32com.client.GetActiveObject(prog_id)
                    if self._active_word_app is not None:
                        print(f"已连接到: {prog_id}")
                        return self._active_word_app
                except Exception as ex:
                    reasons.append(f"{prog_id}: {ex}")
                    continue

            # 没有运行实例，尝试创建新实例
            if is_create_word_app:
                for prog_id in self._WORD_PROG_IDS:
                    try:
                        self._active_word_app = win32com.client.Dispatch(prog_id)
                        self._active_word_app.Visible = True
                        print(f"已创建新实例: {prog_id}")
                        return self._active_word_app
                    except Exception as ex:
                        reasons.append(f"新建 {prog_id}: {ex}")
                        continue

            print("未找到任何可用的 Word/WPS 文字处理应用实例")
            self.last_error = "；".join(reasons)
            return None
        except Exception as ex:
            self.last_error = str(ex)
            print(f"获取Word/WPS实例失败: {ex}")
            return None


def _integrity_rid():
    """本进程令牌的完整性级别 RID（0x0=Untrusted、0x1000=Low、0x2000=Medium…）。

    查询失败返回 None。
    """
    import ctypes
    from ctypes import wintypes
    adv, k32 = ctypes.windll.advapi32, ctypes.windll.kernel32
    k32.GetCurrentProcess.restype = wintypes.HANDLE
    adv.OpenProcessToken.argtypes = [wintypes.HANDLE, wintypes.DWORD,
                                     ctypes.POINTER(wintypes.HANDLE)]
    adv.OpenProcessToken.restype = wintypes.BOOL
    adv.GetTokenInformation.argtypes = [wintypes.HANDLE, ctypes.c_int,
                                        ctypes.c_void_p, wintypes.DWORD,
                                        ctypes.POINTER(wintypes.DWORD)]
    adv.GetTokenInformation.restype = wintypes.BOOL
    adv.GetSidSubAuthorityCount.argtypes = [ctypes.c_void_p]
    adv.GetSidSubAuthorityCount.restype = ctypes.c_void_p
    adv.GetSidSubAuthority.argtypes = [ctypes.c_void_p, wintypes.DWORD]
    adv.GetSidSubAuthority.restype = ctypes.c_void_p

    handle = wintypes.HANDLE()
    if not adv.OpenProcessToken(k32.GetCurrentProcess(), 0x0008,
                                ctypes.byref(handle)):
        return None
    try:
        need = wintypes.DWORD(0)
        adv.GetTokenInformation(handle, 25, None, 0, ctypes.byref(need))  # 25=TokenIntegrityLevel
        buf = ctypes.create_string_buffer(need.value)
        if not adv.GetTokenInformation(handle, 25, buf, need.value,
                                       ctypes.byref(need)):
            return None
        sid = ctypes.cast(buf, ctypes.POINTER(ctypes.c_void_p))[0]
        count = ctypes.cast(adv.GetSidSubAuthorityCount(sid),
                            ctypes.POINTER(ctypes.c_ubyte))[0]
        last = adv.GetSidSubAuthority(sid, count - 1)
        return ctypes.cast(last, ctypes.POINTER(wintypes.DWORD))[0]
    except Exception:
        return None
    finally:
        k32.CloseHandle(handle)


def is_restricted_run():
    """当前进程是不是被云同步客户端降权运行了？

    实测：从云同步盘（如 360 安全云盘同步版）目录里直接启动的 exe，
    进程令牌的完整性级别会被降成 Low —— 低完整性级别的进程受系统隔离限制：
    看不到中完整性进程（Word/WPS）的运行对象表，也写不了同步目录以外的文件。
    这个降级会传给它的所有子进程，进程内没有任何痕迹，程序自己救不回来，
    只能提示用户改用同目录的「启动程序.cmd」。

    查完整性级别是首选判据；查不到时退回到「能不能写系统临时目录」。
    """
    import os
    rid = _integrity_rid()
    if rid is not None and rid <= 0x1000:      # Untrusted(0) / Low(0x1000)
        return True

    # 兜底：往系统临时目录写个文件。
    # 注意不能用 tempfile.gettempdir()——它在 %TEMP% 写不进去时会
    # 回退到当前目录，于是永远"成功"，反而漏判。
    temp = os.environ.get("TEMP") or os.environ.get("TMP")
    if not temp:
        return False
    probe = os.path.join(temp, "DocFormatTool_write_probe.tmp")
    try:
        with open(probe, "w", encoding="utf-8") as fh:
            fh.write("x")
        os.remove(probe)
        return False
    except Exception:
        return True


def restricted_run_message():
    """受限运行时给用户的完整解决办法。"""
    return (
        "程序现在是从云同步盘的目录里直接运行的。同步客户端会把同步目录里\n"
        "启动的程序降成「低完整性级别」运行，这种进程受 Windows 隔离限制：\n"
        "\n"
        "    · 看不到你已经打开的 Word / WPS，所以「选取当前文档」必然失败；\n"
        "    · 也写不了同步目录以外的文件。\n"
        "\n"
        "这个降级会传给程序启动的所有子进程，程序自己解不开。请这样做：\n"
        "\n"
        "1. 关掉本窗口；\n"
        "2. 到程序所在的文件夹里，双击「启动程序.cmd」（不是 exe）；\n"
        "   它会把程序复制到系统目录后再启动，功能完全正常。\n"
        "\n"
        "以后建快捷方式（包括在别的电脑上），请让快捷方式指向「启动程序.cmd」，\n"
        "不要指向 DocFormatTool.exe。"
    )


def _sandbox_hint():
    return (
        "如果程序是从云同步盘（如 360 安全云盘同步版）的目录里直接双击 exe 运行的，"
        "请改成双击同一个目录下的「启动程序.cmd」。同步客户端会把同步目录里的 exe "
        "放进受限沙箱里运行，沙箱内的程序看不到已经打开的 Word/WPS，也写不了文件。"
    )


def describe_no_word_app(reason=""):
    """连接不到 Word/WPS 时给最终用户看的说明（含失败原因与沙箱提示）。"""
    restricted = is_restricted_run()
    lines = []
    if restricted:
        # 受限运行时先说解决办法，别让用户去点「选取当前文档」白费功夫
        lines += [restricted_run_message(), "", "———", ""]
    lines.append("没有找到正在运行的 Word / WPS。")
    if reason:
        lines += ["", "失败原因：", reason[:400]]
    lines += [
        "",
        "请确认：",
        "1. 这台电脑已安装 Microsoft Word 或 WPS 文字，并且已经打开要处理的文档；",
    ]
    if not restricted:
        lines.append("2. " + _sandbox_hint())
    return "\n".join(lines)
