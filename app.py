from __future__ import annotations

import webbrowser

import customtkinter as ctk

try:
    from tkinterdnd2 import TkinterDnD
    TKDND_AVAILABLE = True
except Exception:
    TkinterDnD = None
    TKDND_AVAILABLE = False
from PIL import Image

from tabs.common import resource_path
from tabs.rename_tab import RenameTab
from tabs.compress_tab import CompressTab
from tabs.metadata_tab import MetadataTab
from tabs.citation_tab import CitationTab
from tabs.image_tab import ImageTab
from tabs.help_tab import HelpTab

APP_TITLE = "PaperKit"
APP_VERSION = "v1.2.0"
APP_SUBTITLE = "Academic PDF workflow toolkit"
PROGRAMMER_NAME = "Farid Gazani"
GITHUB_PAGE = "https://github.com/faridinho/PaperKit"
APP_ICON_PATH = "assets/paperkit_icon.ico"
APP_LOGO_PATH = "assets/paperkit_logo.png"

# Modern dark palette
COLORS = {
    "app_bg": "#0F172A",        # slate-950
    "sidebar": "#111827",       # gray-900
    "sidebar_soft": "#1E293B",  # slate-800
    "content_bg": "#0B1120",    # deep navy
    "card": "#0F1B33",
    "card_light": "#162640",
    "primary": "#2563EB",       # blue-600
    "primary_hover": "#1D4ED8", # blue-700
    "accent": "#38BDF8",        # sky-400
    "text": "#F8FAFC",          # slate-50
    "muted": "#94A3B8",         # slate-400
    "border": "#334155",        # slate-700
    "danger": "#DC2626",        # red-600
    "danger_hover": "#B91C1C",
}


class PaperKitApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        if TKDND_AVAILABLE:
            try:
                self.TkdndVersion = TkinterDnD._require(self)
            except Exception as exc:
                print(f"Drag-and-drop could not be initialized: {exc}")

        self.title(f"{APP_TITLE} {APP_VERSION}")
        self.geometry("1240x860")
        self.minsize(1060, 740)
        self.configure(fg_color=COLORS["app_bg"])

        try:
            self.iconbitmap(resource_path(APP_ICON_PATH))
        except Exception as exc:
            print(f"Icon could not be loaded: {exc}")

        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        self.logo_image = None
        self.pages: dict[str, ctk.CTkFrame] = {}
        self.nav_buttons: dict[str, ctk.CTkButton] = {}
        self.current_page = "Rename"
        self.appearance_var = ctk.StringVar(value="Dark")

        self._build_ui()
        self.show_page("Rename")

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main_area()

    def _build_sidebar(self) -> None:
        sidebar = ctk.CTkFrame(self, width=250, corner_radius=0, fg_color=COLORS["sidebar"])
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_rowconfigure(9, weight=1)
        sidebar.grid_columnconfigure(0, weight=1)
        sidebar.grid_propagate(False)

        brand = ctk.CTkFrame(sidebar, fg_color="transparent")
        brand.grid(row=0, column=0, padx=18, pady=(22, 18), sticky="ew")
        brand.grid_columnconfigure(1, weight=1)

        try:
            image = Image.open(resource_path(APP_LOGO_PATH))
            self.logo_image = ctk.CTkImage(light_image=image, dark_image=image, size=(85, 85))
            ctk.CTkLabel(brand, image=self.logo_image, text="").grid(row=0, column=0, rowspan=2, padx=(0, 12), sticky="w")
        except Exception as exc:
            print(f"Logo could not be loaded: {exc}")

        ctk.CTkLabel(
            brand,
            text=APP_TITLE,
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=COLORS["text"],
            anchor="w",
        ).grid(row=0, column=1, sticky="w")

        ctk.CTkLabel(
            brand,
            text=APP_VERSION,
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=COLORS["accent"],
            anchor="w",
        ).grid(row=1, column=1, sticky="w")

        ctk.CTkLabel(
            sidebar,
            text="TOOLS",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=COLORS["muted"],
            anchor="w",
        ).grid(row=1, column=0, padx=24, pady=(8, 6), sticky="ew")

        nav_items = [
            ("Rename", "Rename"),
            ("Compress", "Compress"),
            ("Metadata", "Metadata"),
            ("Citations", "Citations"),
            ("Image Extraction", "Image Extraction"),
        ]

        for idx, (key, label) in enumerate(nav_items, start=2):
            self.nav_buttons[key] = self._make_nav_button(sidebar, label, key)
            self.nav_buttons[key].grid(row=idx, column=0, padx=14, pady=5, sticky="ew")

        ctk.CTkLabel(
            sidebar,
            text="SUPPORT",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=COLORS["muted"],
            anchor="w",
        ).grid(row=7, column=0, padx=24, pady=(20, 6), sticky="ew")

        self.nav_buttons["Help"] = self._make_nav_button(sidebar, "Help", "Help")
        self.nav_buttons["Help"].grid(row=8, column=0, padx=14, pady=5, sticky="new")

        footer = ctk.CTkFrame(sidebar, fg_color="transparent")
        footer.grid(row=10, column=0, padx=18, pady=(10, 18), sticky="sew")
        footer.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            footer,
            text=APP_SUBTITLE,
            font=ctk.CTkFont(size=12),
            text_color=COLORS["muted"],
            wraplength=200,
            justify="left",
            anchor="w",
        ).grid(row=0, column=0, sticky="ew", pady=(0, 10))

        ctk.CTkLabel(
            footer,
            text="Appearance",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=COLORS["muted"],
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", pady=(0, 6))

        ctk.CTkOptionMenu(
            footer,
            values=["Dark", "Light", "System"],
            variable=self.appearance_var,
            height=34,
            fg_color=COLORS["sidebar_soft"],
            button_color=COLORS["primary"],
            button_hover_color=COLORS["primary_hover"],
            command=self.change_appearance,
        ).grid(row=2, column=0, sticky="ew", pady=(0, 10))

        ctk.CTkButton(
            footer,
            text="Open GitHub",
            height=36,
            corner_radius=10,
            fg_color=COLORS["sidebar_soft"],
            hover_color=COLORS["primary"],
            command=lambda: webbrowser.open(GITHUB_PAGE),
        ).grid(row=3, column=0, sticky="ew")


    def change_appearance(self, value: str) -> None:
        ctk.set_appearance_mode(value)

    def _make_nav_button(self, master: ctk.CTkFrame, text: str, page_name: str) -> ctk.CTkButton:
        return ctk.CTkButton(
            master,
            text=text,
            height=44,
            corner_radius=12,
            anchor="w",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="transparent",
            hover_color=COLORS["sidebar_soft"],
            text_color=COLORS["text"],
            command=lambda: self.show_page(page_name),
        )

    def _build_main_area(self) -> None:
        main = ctk.CTkFrame(self, fg_color=COLORS["content_bg"], corner_radius=0)
        main.grid(row=0, column=1, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        self.page_header = ctk.CTkFrame(main, fg_color="transparent")
        self.page_header.grid(row=0, column=0, padx=26, pady=(22, 8), sticky="ew")
        self.page_header.grid_columnconfigure(0, weight=1)

        self.page_title = ctk.CTkLabel(
            self.page_header,
            text="",
            font=ctk.CTkFont(size=30, weight="bold"),
            text_color=COLORS["text"],
            anchor="w",
        )
        self.page_title.grid(row=0, column=0, sticky="ew")

        self.page_subtitle = ctk.CTkLabel(
            self.page_header,
            text="",
            font=ctk.CTkFont(size=14),
            text_color=COLORS["muted"],
            anchor="w",
        )
        self.page_subtitle.grid(row=1, column=0, sticky="ew", pady=(3, 0))

        self.content_frame = ctk.CTkFrame(main, fg_color="transparent")
        self.content_frame.grid(row=1, column=0, padx=20, pady=(4, 20), sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)

        self.pages["Rename"] = RenameTab(self.content_frame)
        self.pages["Compress"] = CompressTab(self.content_frame)
        self.pages["Metadata"] = MetadataTab(self.content_frame)
        self.pages["Citations"] = CitationTab(self.content_frame)
        self.pages["Image Extraction"] = ImageTab(self.content_frame)
        self.pages["Help"] = HelpTab(
            self.content_frame,
            app_title=APP_TITLE,
            app_version=APP_VERSION,
            github_page=GITHUB_PAGE,
        )

        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")
            page.grid_remove()

    def show_page(self, page_name: str) -> None:
        if page_name not in self.pages:
            return

        for name, page in self.pages.items():
            if name == page_name:
                page.grid()
            else:
                page.grid_remove()

        self.current_page = page_name
        self._update_nav_styles()
        self._update_header(page_name)

    def _update_nav_styles(self) -> None:
        for name, button in self.nav_buttons.items():
            if name == self.current_page:
                button.configure(
                    fg_color=COLORS["primary"],
                    hover_color=COLORS["primary_hover"],
                    text_color="white",
                )
            else:
                button.configure(
                    fg_color="transparent",
                    hover_color=COLORS["sidebar_soft"],
                    text_color=COLORS["text"],
                )

    def _update_header(self, page_name: str) -> None:
        subtitles = {
            "Rename": "Preview, detect duplicates, and safely rename academic PDFs.",
            "Compress": "Reduce PDF file size with Ghostscript profiles and cancellation support.",
            "Metadata": "Choose metadata fields and export a clean Excel spreadsheet.",
            "Citations": "Automatically export RIS and BibTeX citation files from available PDF metadata.",
            "Image Extraction": "Preview embedded PDF images and extract only the images you select.",
            "Help": "Detailed setup notes and a practical guide to each PaperKit function.",
        }
        titles = {
            "Rename": "Rename PDFs",
            "Compress": "Compress PDFs",
            "Metadata": "Metadata export",
            "Citations": "Citation export",
            "Image Extraction": "Image extraction",
            "Help": "Help",
        }
        self.page_title.configure(text=titles.get(page_name, page_name))
        self.page_subtitle.configure(text=subtitles.get(page_name, ""))


if __name__ == "__main__":
    app = PaperKitApp()
    app.mainloop()
