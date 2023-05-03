from XLSXCreator import XLSXCreator
from SapoWindow import SapoWindow
import ttkbootstrap as ttk

# sistema absolutamente podereoso de organização

app = ttk.Window(
    title="SAPO",
    themename="yeti",
    #https://ttkbootstrap.readthedocs.io/en/latest/themes/
    minsize=(1000,300)
)
SapoWindow(app)
app.mainloop()

