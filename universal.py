import win32com.client
import pythoncom


class Universal:
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
        try:
            pythoncom.CoInitialize()
            self._active_word_app = win32com.client.GetActiveObject("Word.Application")
            if self._active_word_app is None and is_create_word_app:
                self._active_word_app = win32com.client.Dispatch("Word.Application")
                self._active_word_app.Visible = True
            return self._active_word_app
        except Exception as ex:
            print(f"获取Word实例失败: {ex}")
            return None