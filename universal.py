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
        """
        try:
            pythoncom.CoInitialize()

            # 先尝试获取已运行的实例
            for prog_id in self._WORD_PROG_IDS:
                try:
                    self._active_word_app = win32com.client.GetActiveObject(prog_id)
                    if self._active_word_app is not None:
                        print(f"已连接到: {prog_id}")
                        return self._active_word_app
                except Exception:
                    continue

            # 没有运行实例，尝试创建新实例
            if is_create_word_app:
                for prog_id in self._WORD_PROG_IDS:
                    try:
                        self._active_word_app = win32com.client.Dispatch(prog_id)
                        self._active_word_app.Visible = True
                        print(f"已创建新实例: {prog_id}")
                        return self._active_word_app
                    except Exception:
                        continue

            print("未找到任何可用的 Word/WPS 文字处理应用实例")
            return None
        except Exception as ex:
            print(f"获取Word/WPS实例失败: {ex}")
            return None