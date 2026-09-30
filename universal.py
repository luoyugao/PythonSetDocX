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


def _sandbox_hint():
    return (
        "如果程序是从云同步盘（如 360 安全云盘同步版）的目录里直接双击 exe 运行的，"
        "请改成双击同一个目录下的「启动程序.cmd」。同步客户端会把同步目录里的 exe "
        "放进受限沙箱里运行，沙箱内的程序看不到已经打开的 Word/WPS，也写不了文件。"
    )


def describe_no_word_app(reason=""):
    """连接不到 Word/WPS 时给最终用户看的说明（含失败原因与沙箱提示）。"""
    lines = ["没有找到正在运行的 Word / WPS。"]
    if reason:
        lines += ["", "失败原因：", reason[:400]]
    lines += [
        "",
        "请确认：",
        "1. 这台电脑已安装 Microsoft Word 或 WPS 文字，并且已经打开要处理的文档；",
        "2. " + _sandbox_hint(),
    ]
    return "\n".join(lines)
