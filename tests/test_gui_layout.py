import inspect
import tkinter as tk

from gui_origin import OriginFrame
from gui_shared import create_page
from gui_tensile import TensileFrame
from main import MainApp, StartupSplash, __version__, get_bootloader_splash


def test_v316_spec_uses_pyinstaller_bootloader_splash():
    spec_text = open("育材堂报告助手V3.16.spec", encoding="utf-8").read()

    assert "Splash(" in spec_text
    assert "assets/startup_splash.png" in spec_text
    assert "splash," in spec_text
    assert "splash.binaries," in spec_text


def test_main_uses_a_loading_screen_before_enabling_drag_and_drop():
    main_source = inspect.getsource(__import__("main").main)

    assert "root = tk.Tk()" in main_source
    assert "root = TkinterDnD.Tk()" not in main_source
    assert "root.withdraw()" in main_source
    assert "splash = StartupSplash(root)" in main_source
    assert "root.TkdndVersion = TkinterDnD._require(root)" in main_source
    assert main_source.index("root.withdraw()") < main_source.index("splash = StartupSplash(root)")
    assert main_source.index("splash = StartupSplash(root)") < main_source.index("TkinterDnD._require(root)")


def _all_label_texts(widget):
    texts = []
    for child in widget.winfo_children():
        if isinstance(child, tk.Label):
            texts.append(child.cget("text"))
        texts.extend(_all_label_texts(child))
    return texts


def test_gui_layout_startup_and_removed_helper_text_are_configured():
    root = tk.Tk()
    root.withdraw()
    try:
        splash = StartupSplash(root)
        root.update_idletasks()

        assert splash.window.winfo_manager() == "place"
        assert splash.spinner.cget("text")
        splash.destroy()

        assert get_bootloader_splash() is None

        host = tk.Frame(root)
        host.pack(fill="both", expand=True)
        page = create_page(host, scrollable=True)
        root.update_idletasks()

        assert hasattr(page, "_scrollable_page")
        assert page is page._scrollable_page.scrollable_frame
        assert page._scrollable_page.canvas.winfo_manager() == "pack"

        app = MainApp(root)
        root.geometry("860x700")
        root.update_idletasks()

        subtitles = [
            widget
            for widget in root.winfo_children()[0].winfo_children()[0].winfo_children()[0].winfo_children()
            if isinstance(widget, tk.Label)
            and widget.cget("text") == f"材料试验报告处理与Origin绘图工具 V{__version__}"
        ]

        assert subtitles
        assert subtitles[0].cget("anchor") == "w"
        assert subtitles[0].cget("justify") == "left"
        assert root.tk.getint(subtitles[0].cget("wraplength")) > 0

        tensile = TensileFrame(root)
        origin = OriginFrame(root)
        root.update_idletasks()

        all_text = "\n".join(_all_label_texts(tensile) + _all_label_texts(origin))

        removed_text = [
            "选择报告中展示的延伸率组合。",
            "延伸率字段",
            "设置 Origin 模板、每图曲线数、XY 列与输出尺寸。",
            "导入 Word 或 Excel 原始数据，生成报告或绘制 Origin 曲线。",
            "设置 Origin 模板、图片尺寸和 PPT 输出方式。",
        ]
        for text in removed_text:
            assert text not in all_text
    finally:
        root.destroy()
