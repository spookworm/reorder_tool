import main

# ============================================================
# GUI APPLICATION
# ============================================================

class RankerApp:
    def __init__(
        self,
        root: tk.Tk,
    ):
        self.root = root

        self.root.title(
            f"{APP_NAME} {APP_VERSION}"
        )

        self.root.geometry(
            f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}"
        )

        self.root.minsize(
            1050,
            720,
        )

        self.engine = None
        self.source_file = None
        self.state_store = None
        self.current_pair = None

        self.target_var = tk.IntVar(
            value=DEFAULT_TARGET_COMPARISONS
        )

        self.mode_var = tk.StringVar(
            value="BALANCED"
        )

        self.auto_stop_var = tk.BooleanVar(
            value=True
        )

        self.status_var = tk.StringVar(
            value=(
                "Open your Goodreads Excel export "
                "to begin."
            )
        )

        self.progress_var = tk.DoubleVar(
            value=0.0
        )

        self.stats_var = tk.StringVar(
            value=""
        )

        self.left_title_var = tk.StringVar()
        self.left_author_var = tk.StringVar()
        self.left_meta_var = tk.StringVar()

        self.right_title_var = tk.StringVar()
        self.right_author_var = tk.StringVar()
        self.right_meta_var = tk.StringVar()

        self.build_styles()
        self.build_menu()
        self.build_ui()
        self.bind_keys()

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self.on_close,
        )

    # ========================================================
    # STYLES
    # ========================================================

    def build_styles(self):
        self.style = ttk.Style(self.root)

        # "clam" is much easier to theme consistently than Vista.
        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        self.colors = {
            "bg": "#0D0F14",
            "surface": "#151820",
            "surface_2": "#1B1F2A",
            "surface_3": "#222735",
            "border": "#2A3040",
            "text": "#F4F5F7",
            "muted": "#8E96A8",
            "accent": "#8B5CF6",
            "accent_hover": "#9D72FF",
            "accent_dark": "#6D42D8",
            "green": "#34D399",
            "red": "#FB7185",
            "yellow": "#FBBF24",
            "blue": "#60A5FA",
        }

        c = self.colors

        self.root.configure(bg=c["bg"])

        # --------------------------------------------------------
        # General
        # --------------------------------------------------------

        self.style.configure(
            ".",
            background=c["bg"],
            foreground=c["text"],
            font=("Segoe UI", 10),
        )

        self.style.configure(
            "TFrame",
            background=c["bg"],
        )

        self.style.configure(
            "Surface.TFrame",
            background=c["surface"],
        )

        self.style.configure(
            "Surface2.TFrame",
            background=c["surface_2"],
        )

        self.style.configure(
            "TLabel",
            background=c["bg"],
            foreground=c["text"],
        )

        self.style.configure(
            "Muted.TLabel",
            background=c["bg"],
            foreground=c["muted"],
        )

        # --------------------------------------------------------
        # Header
        # --------------------------------------------------------

        self.style.configure(
            "Title.TLabel",
            background=c["bg"],
            foreground=c["text"],
            font=("Segoe UI", 24, "bold"),
        )

        self.style.configure(
            "Subtitle.TLabel",
            background=c["bg"],
            foreground=c["muted"],
            font=("Segoe UI", 10),
        )

        self.style.configure(
            "Section.TLabel",
            background=c["bg"],
            foreground=c["muted"],
            font=("Segoe UI", 9, "bold"),
        )

        # --------------------------------------------------------
        # Book cards
        # --------------------------------------------------------

        self.style.configure(
            "BookTitle.TLabel",
            background=c["surface"],
            foreground=c["text"],
            font=("Segoe UI", 19, "bold"),
        )

        self.style.configure(
            "BookAuthor.TLabel",
            background=c["surface"],
            foreground="#B7BECC",
            font=("Segoe UI", 11),
        )

        self.style.configure(
            "BookMeta.TLabel",
            background=c["surface"],
            foreground=c["muted"],
            font=("Segoe UI", 9),
        )

        # --------------------------------------------------------
        # Buttons
        # --------------------------------------------------------

        self.style.configure(
            "Choice.TButton",
            background=c["surface_2"],
            foreground=c["text"],
            borderwidth=0,
            relief="flat",
            font=("Segoe UI", 12, "bold"),
            padding=(18, 14),
        )

        self.style.map(
            "Choice.TButton",
            background=[
                ("active", c["accent"]),
                ("pressed", c["accent_dark"]),
                ("disabled", c["surface"]),
            ],
            foreground=[
                ("disabled", "#555B6B"),
            ],
        )

        self.style.configure(
            "Primary.TButton",
            background=c["accent"],
            foreground="#FFFFFF",
            borderwidth=0,
            relief="flat",
            font=("Segoe UI", 10, "bold"),
            padding=(14, 9),
        )

        self.style.map(
            "Primary.TButton",
            background=[
                ("active", c["accent_hover"]),
                ("pressed", c["accent_dark"]),
            ],
        )

        self.style.configure(
            "Secondary.TButton",
            background=c["surface_2"],
            foreground="#D8DCE5",
            borderwidth=1,
            bordercolor=c["border"],
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padding=(12, 8),
        )

        self.style.map(
            "Secondary.TButton",
            background=[
                ("active", c["surface_3"]),
            ],
            foreground=[
                ("active", "#FFFFFF"),
            ],
        )

        self.style.configure(
            "Preset.TButton",
            background=c["surface"],
            foreground="#AEB6C6",
            borderwidth=1,
            bordercolor=c["border"],
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padding=(10, 6),
        )

        self.style.map(
            "Preset.TButton",
            background=[
                ("active", c["surface_3"]),
            ],
            foreground=[
                ("active", "#FFFFFF"),
            ],
            bordercolor=[
                ("active", c["accent"]),
            ],
        )

        # --------------------------------------------------------
        # Spinbox / checkbox
        # --------------------------------------------------------

        self.style.configure(
            "TSpinbox",
            fieldbackground=c["surface"],
            background=c["surface"],
            foreground=c["text"],
            bordercolor=c["border"],
            arrowcolor=c["muted"],
        )

        self.style.configure(
            "TCheckbutton",
            background=c["bg"],
            foreground=c["muted"],
            font=("Segoe UI", 9),
        )

        self.style.map(
            "TCheckbutton",
            background=[
                ("active", c["bg"]),
            ],
            foreground=[
                ("active", c["text"]),
            ],
        )

        # --------------------------------------------------------
        # Progress
        # --------------------------------------------------------

        self.style.configure(
            "Dark.Horizontal.TProgressbar",
            troughcolor=c["surface"],
            background=c["accent"],
            borderwidth=0,
            thickness=7,
        )

    # ========================================================
    # MENU
    # ========================================================

    def build_menu(self):
        menu = tk.Menu(self.root)

        file_menu = tk.Menu(
            menu,
            tearoff=False,
        )

        file_menu.add_command(
            label="Open Goodreads Excel…",
            command=self.open_file,
            accelerator="Ctrl+O",
        )

        file_menu.add_command(
            label="Export Ranking",
            command=self.export,
            accelerator="Ctrl+E",
        )

        file_menu.add_separator()

        file_menu.add_command(
            label="Exit",
            command=self.on_close,
        )

        menu.add_cascade(
            label="File",
            menu=file_menu,
        )

        ranking_menu = tk.Menu(
            menu,
            tearoff=False,
        )

        ranking_menu.add_command(
            label="Undo Last Choice",
            command=self.undo,
            accelerator="U",
        )

        ranking_menu.add_command(
            label="Finish Now",
            command=self.finish_now,
        )

        ranking_menu.add_command(
            label="View Ranking",
            command=self.show_ranking,
        )

        ranking_menu.add_command(
            label="View Diagnostics",
            command=self.show_diagnostics,
        )

        menu.add_cascade(
            label="Ranking",
            menu=ranking_menu,
        )

        help_menu = tk.Menu(
            menu,
            tearoff=False,
        )

        help_menu.add_command(
            label="Keyboard Shortcuts",
            command=self.show_shortcuts,
        )

        help_menu.add_command(
            label="About",
            command=self.show_about,
        )

        menu.add_cascade(
            label="Help",
            menu=help_menu,
        )

        self.root.config(
            menu=menu
        )

    # ========================================================
    # MAIN UI
    # ========================================================

    def build_ui(self):
        c = self.colors

        outer = tk.Frame(
            self.root,
            bg=c["bg"],
        )

        outer.pack(
            fill="both",
            expand=True,
            padx=26,
            pady=22,
        )

        # ========================================================
        # HEADER
        # ========================================================

        header = tk.Frame(
            outer,
            bg=c["bg"],
        )

        header.pack(
            fill="x",
            pady=(0, 18),
        )

        title_area = tk.Frame(
            header,
            bg=c["bg"],
        )

        title_area.pack(
            side="left",
        )

        tk.Label(
            title_area,
            text="GOODREADS",
            bg=c["bg"],
            fg=c["accent"],
            font=("Segoe UI", 9, "bold"),
        ).pack(
            anchor="w",
        )

        ttk.Label(
            title_area,
            text="Your To-Read Ranker",
            style="Title.TLabel",
        ).pack(
            anchor="w",
        )

        ttk.Label(
            title_area,
            text="Build your personal reading order, one decision at a time.",
            style="Subtitle.TLabel",
        ).pack(
            anchor="w",
            pady=(2, 0),
        )

        controls = tk.Frame(
            header,
            bg=c["bg"],
        )

        controls.pack(
            side="right",
            anchor="n",
        )

        ttk.Label(
            controls,
            text="EVIDENCE",
            style="Section.TLabel",
        ).grid(
            row=0,
            column=0,
            padx=(0, 7),
        )

        self.spinbox = ttk.Spinbox(
            controls,
            from_=MIN_TARGET_COMPARISONS,
            to=MAX_TARGET_COMPARISONS,
            textvariable=self.target_var,
            width=5,
        )

        self.spinbox.grid(
            row=0,
            column=1,
        )

        self.spinbox.bind(
            "<Return>",
            self.target_changed,
        )

        ttk.Checkbutton(
            controls,
            text="Auto-stop",
            variable=self.auto_stop_var,
        ).grid(
            row=0,
            column=2,
            padx=(14, 12),
        )

        ttk.Button(
            controls,
            text="+  Open Excel",
            command=self.open_file,
            style="Primary.TButton",
        ).grid(
            row=0,
            column=3,
        )

        # ========================================================
        # STRATEGY BAR
        # ========================================================

        strategy = tk.Frame(
            outer,
            bg=c["surface"],
            highlightbackground=c["border"],
            highlightthickness=1,
        )

        strategy.pack(
            fill="x",
            pady=(0, 14),
        )

        inner = tk.Frame(
            strategy,
            bg=c["surface"],
        )

        inner.pack(
            fill="x",
            padx=12,
            pady=10,
        )

        tk.Label(
            inner,
            text="STRATEGY",
            bg=c["surface"],
            fg=c["muted"],
            font=("Segoe UI", 8, "bold"),
        ).pack(
            side="left",
            padx=(2, 10),
        )

        presets = [
            ("⚡ Quick", "QUICK"),
            ("Balanced", "BALANCED"),
            ("Accurate", "ACCURATE"),
            ("🏆 Top 25", "TOP_25_FOCUS"),
            ("Maximum", "MAX_ACCURACY"),
        ]

        for label, mode in presets:
            ttk.Button(
                inner,
                text=label,
                style="Preset.TButton",
                command=lambda m=mode: self.set_mode(m),
            ).pack(
                side="left",
                padx=2,
            )

        tk.Label(
            inner,
            textvariable=self.mode_var,
            bg=c["surface"],
            fg=c["accent"],
            font=("Segoe UI", 8, "bold"),
        ).pack(
            side="left",
            padx=(12, 0),
        )

        # ========================================================
        # STATUS / PROGRESS
        # ========================================================

        status = tk.Frame(
            outer,
            bg=c["bg"],
        )

        status.pack(
            fill="x",
            pady=(0, 16),
        )

        status_top = tk.Frame(
            status,
            bg=c["bg"],
        )

        status_top.pack(
            fill="x",
        )

        ttk.Label(
            status_top,
            textvariable=self.status_var,
            style="Subtitle.TLabel",
        ).pack(
            side="left",
        )

        ttk.Label(
            status_top,
            textvariable=self.stats_var,
            style="Subtitle.TLabel",
        ).pack(
            side="right",
        )

        ttk.Progressbar(
            status,
            variable=self.progress_var,
            maximum=100,
            style="Dark.Horizontal.TProgressbar",
        ).pack(
            fill="x",
            pady=(9, 0),
        )

        # ========================================================
        # BOOK COMPARISON
        # ========================================================

        choices = tk.Frame(
            outer,
            bg=c["bg"],
        )

        choices.pack(
            fill="both",
            expand=True,
        )

        self.left_card = self.make_book_card(
            choices,
            "left",
        )

        self.left_card.pack(
            side="left",
            fill="both",
            expand=True,
        )

        # Center VS area
        middle = tk.Frame(
            choices,
            bg=c["bg"],
            width=90,
        )

        middle.pack(
            side="left",
            fill="y",
        )

        middle.pack_propagate(False)

        tk.Frame(
            middle,
            bg=c["border"],
            width=1,
        ).pack(
            side="left",
            fill="y",
            pady=35,
        )

        vs_area = tk.Frame(
            middle,
            bg=c["bg"],
        )

        vs_area.place(
            relx=0.5,
            rely=0.5,
            anchor="center",
        )

        tk.Label(
            vs_area,
            text="VS",
            bg=c["bg"],
            fg=c["muted"],
            font=("Segoe UI", 11, "bold"),
        ).pack()

        tk.Label(
            vs_area,
            text="choose",
            bg=c["bg"],
            fg="#555C6D",
            font=("Segoe UI", 8),
        ).pack(
            pady=(2, 0),
        )

        self.right_card = self.make_book_card(
            choices,
            "right",
        )

        self.right_card.pack(
            side="left",
            fill="both",
            expand=True,
        )

        # ========================================================
        # CHOICE BUTTONS
        # ========================================================

        buttons = tk.Frame(
            outer,
            bg=c["bg"],
        )

        buttons.pack(
            fill="x",
            pady=(15, 0),
        )

        self.left_button = ttk.Button(
            buttons,
            text="←  LEFT     [1]",
            command=lambda: self.choose("left"),
            style="Choice.TButton",
        )

        self.left_button.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 5),
        )

        self.tie_button = ttk.Button(
            buttons,
            text="≈  EQUAL     [3]",
            command=lambda: self.choose("tie"),
            style="Choice.TButton",
        )

        self.tie_button.pack(
            side="left",
            fill="x",
            expand=True,
            padx=5,
        )

        self.right_button = ttk.Button(
            buttons,
            text="RIGHT     [2]  →",
            command=lambda: self.choose("right"),
            style="Choice.TButton",
        )

        self.right_button.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(5, 0),
        )

        # ========================================================
        # FOOTER
        # ========================================================

        footer = tk.Frame(
            outer,
            bg=c["bg"],
        )

        footer.pack(
            fill="x",
            pady=(12, 0),
        )

        ttk.Button(
            footer,
            text="Undo  U",
            command=self.undo,
            style="Secondary.TButton",
        ).pack(
            side="left",
        )

        ttk.Button(
            footer,
            text="Skip  4",
            command=self.skip,
            style="Secondary.TButton",
        ).pack(
            side="left",
            padx=6,
        )

        ttk.Button(
            footer,
            text="Finish",
            command=self.finish_now,
            style="Secondary.TButton",
        ).pack(
            side="left",
        )

        ttk.Button(
            footer,
            text="Ranking",
            command=self.show_ranking,
            style="Secondary.TButton",
        ).pack(
            side="right",
        )

        ttk.Button(
            footer,
            text="Export",
            command=self.export,
            style="Secondary.TButton",
        ).pack(
            side="right",
            padx=6,
        )

        tk.Label(
            footer,
            text="1 / 2 choose   ·   3 equal   ·   4 skip   ·   U undo",
            bg=c["bg"],
            fg="#555C6D",
            font=("Segoe UI", 8),
        ).pack(
            side="right",
            padx=12,
        )

    # ========================================================
    # BOOK CARD
    # ========================================================
    def make_book_card(self, parent, side):
        c = self.colors

        frame = tk.Frame(
            parent,
            bg=c["surface"],
            highlightbackground=c["border"],
            highlightthickness=1,
            bd=0,
        )

        content = ttk.Frame(
            frame,
            style="Surface.TFrame",
            padding=28,
        )

        content.pack(
            fill="both",
            expand=True,
        )

        if side == "left":
            title_var = self.left_title_var
            author_var = self.left_author_var
            meta_var = self.left_meta_var
        else:
            title_var = self.right_title_var
            author_var = self.right_author_var
            meta_var = self.right_meta_var

        # Small label at top of card.
        side_label = "OPTION A" if side == "left" else "OPTION B"

        ttk.Label(
            content,
            text=side_label,
            style="Section.TLabel",
        ).pack(
            pady=(4, 22),
        )

        title = ttk.Label(
            content,
            textvariable=title_var,
            style="BookTitle.TLabel",
            wraplength=450,
            justify="center",
        )

        title.pack(
            pady=(0, 10),
        )

        ttk.Label(
            content,
            textvariable=author_var,
            style="BookAuthor.TLabel",
            wraplength=450,
            justify="center",
        ).pack(
            pady=(0, 14),
        )

        ttk.Label(
            content,
            textvariable=meta_var,
            style="BookMeta.TLabel",
            wraplength=450,
            justify="center",
        ).pack(
            pady=(0, 20),
        )

        description = tk.Text(
            content,
            height=10,
            width=45,
            wrap="word",
            relief="flat",
            borderwidth=0,
            bg=c["surface"],
            fg="#B9BFCC",
            insertbackground=c["text"],
            selectbackground=c["accent_dark"],
            selectforeground="#FFFFFF",
            font=("Segoe UI", 10),
            padx=4,
            pady=4,
            highlightthickness=0,
        )

        description.pack(
            fill="both",
            expand=True,
        )

        description.configure(
            state="disabled",
        )

        if side == "left":
            self.left_description = description
        else:
            self.right_description = description

        return frame


    # ========================================================
    # KEYBOARD
    # ========================================================

    def bind_keys(self):
        self.root.bind(
            "<KeyPress-1>",
            lambda event: self.choose(
                "left"
            ),
        )

        self.root.bind(
            "<KeyPress-2>",
            lambda event: self.choose(
                "right"
            ),
        )

        self.root.bind(
            "<KeyPress-3>",
            lambda event: self.choose(
                "tie"
            ),
        )

        self.root.bind(
            "<KeyPress-4>",
            lambda event: self.skip(),
        )

        self.root.bind(
            "<Left>",
            lambda event: self.choose(
                "left"
            ),
        )

        self.root.bind(
            "<Right>",
            lambda event: self.choose(
                "right"
            ),
        )

        self.root.bind(
            "<KeyPress-t>",
            lambda event: self.choose(
                "tie"
            ),
        )

        self.root.bind(
            "<KeyPress-T>",
            lambda event: self.choose(
                "tie"
            ),
        )

        self.root.bind(
            "<KeyPress-s>",
            lambda event: self.skip(),
        )

        self.root.bind(
            "<KeyPress-S>",
            lambda event: self.skip(),
        )

        self.root.bind(
            "<KeyPress-u>",
            lambda event: self.undo(),
        )

        self.root.bind(
            "<KeyPress-U>",
            lambda event: self.undo(),
        )

        self.root.bind(
            "<Control-o>",
            lambda event: self.open_file(),
        )

        self.root.bind(
            "<Control-e>",
            lambda event: self.export(),
        )

    # ========================================================
    # FILE
    # ========================================================

    def open_file(self):
        path = filedialog.askopenfilename(
            title=(
                "Open Goodreads Excel export"
            ),
            filetypes=[
                (
                    "Excel files",
                    "*.xlsx",
                ),
                (
                    "All files",
                    "*.*",
                ),
            ],
        )

        if not path:
            return

        self.load_file(
            Path(path)
        )

    def load_file(
        self,
        path: Path,
    ):
        try:
            _, books = load_goodreads(
                path
            )

        except Exception as exc:
            messagebox.showerror(
                "Could not open Goodreads file",
                str(exc),
            )
            return

        if len(books) < 2:
            messagebox.showwarning(
                "Not enough books",
                (
                    'I found fewer than two books where '
                    '"Exclusive Shelf" is "to-read".'
                ),
            )
            return

        self.source_file = path

        self.state_store = (
            StateStore(path)
        )

        existing = (
            self.state_store.load(
                books
            )
        )

        if existing is not None:
            resume = messagebox.askyesno(
                "Resume previous ranking?",
                (
                    f"I found a saved ranking for:\n\n"
                    f"{path.name}\n\n"
                    f"{len(existing.comparisons)} "
                    f"comparisons have already been made.\n\n"
                    f"Resume it?"
                ),
            )

            if resume:
                self.engine = existing

            else:
                self.engine = (
                    RankingEngine(
                        books,
                        self.target_var.get(),
                    )
                )

        else:
            self.engine = (
                RankingEngine(
                    books,
                    self.target_var.get(),
                )
            )

        self.target_var.set(
            self.engine.target_comparisons
        )
        self.mode_var.set(getattr(self.engine, "mode", "BALANCED"))

        self.current_pair = None

        self.status_var.set(
            (
                f"{len(books)} books on your "
                f"to-read shelf · {path.name}"
            )
        )

        self.refresh()

    # ========================================================
    # SPEED PRESETS / STRATEGIES
    # ========================================================

    def set_mode(self, mode: str):
        if self.engine is None:
            self.mode_var.set(mode)
            return
        mode = str(mode).upper()
        if mode not in PRESETS:
            mode = "BALANCED"
        self.mode_var.set(mode)
        current_target = self.engine.target_comparisons
        self.engine.mode = mode
        self.engine.config = make_config(mode, current_target)
        self.engine.target_comparisons = self.engine.config.target_comparisons
        self.engine._analysis_cache = None
        self.engine._ranking_cache = None
        self.engine.current_mode_changed = True
        self.current_pair = None
        self.save_state()
        self.refresh()

    def set_target(
        self,
        value: int,
    ):
        self.target_var.set(
            value
        )

        self.target_changed()

    def target_changed(
        self,
        event=None,
    ):
        if self.engine is None:
            return

        try:
            value = int(
                self.target_var.get()
            )

        except ValueError:
            value = (
                DEFAULT_TARGET_COMPARISONS
            )

        value = max(
            MIN_TARGET_COMPARISONS,
            min(
                MAX_TARGET_COMPARISONS,
                value,
            ),
        )

        self.target_var.set(
            value
        )

        self.engine.target_comparisons = value
        self.engine.config = make_config(self.engine.mode, value)
        self.engine._analysis_cache = None
        self.engine._ranking_cache = None

        self.save_state()

        self.current_pair = None

        self.refresh()

    # ========================================================
    # REFRESH
    # ========================================================

    def refresh(self):
        if self.engine is None:
            return

        # Ensure model is current before selecting the pair.
        self.engine.ensure_model()

        if (
            self.auto_stop_var.get()
            and self.engine.is_finished()
        ):
            self.show_finished()

            return

        if (
            self.current_pair is None
            or not self.current_pair_valid()
        ):
            self.current_pair = (
                self.engine.choose_pair()
            )

        if self.current_pair is None:
            self.show_finished()

            return

        self.enable_choices()

        left_id, right_id = (
            self.current_pair
        )

        book_map = (
            self.engine.book_map()
        )

        left = book_map[
            left_id
        ]

        right = book_map[
            right_id
        ]

        self.show_book(
            left,
            self.engine.ratings[
                left.id
            ],
            self.left_title_var,
            self.left_author_var,
            self.left_meta_var,
            self.left_description,
        )

        self.show_book(
            right,
            self.engine.ratings[
                right.id
            ],
            self.right_title_var,
            self.right_author_var,
            self.right_meta_var,
            self.right_description,
        )

        progress = (
            self.engine.progress()
            * 100.0
        )

        self.progress_var.set(
            progress
        )

        comparisons = len(
            self.engine.comparisons
        )

        average = (
            self.engine.average_comparisons()
        )

        coverage = (
            self.engine.coverage()
            * 100.0
        )

        stability = (
            self.engine.stability
            * 100.0
        )

        top_stability = (
            self.engine.top_stability
            * 100.0
        )

        diagnostics = self.engine.diagnostics()
        self.stats_var.set(
            (
                f"{comparisons} decisions · {average:.1f} avg/book · "
                f"coverage {coverage:.0f}% · Top-25 confidence {diagnostics['top25_confidence'] * 100:.0f}% · "
                f"Top-25 stability {diagnostics['top25_stability'] * 100:.0f}% · "
                f"~{diagnostics['estimated_additional']} more"
            )
        )

        self.status_var.set(
            (
                f"{self.engine.phase_label()} · "
                "Which book would you rather read?"
            )
        )

    def show_finished(self):
        self.disable_choices()

        self.left_title_var.set(
            "🏆 Ranking ready"
        )

        self.left_author_var.set("")
        self.left_meta_var.set("")

        self.right_title_var.set(
            "You can export now"
        )

        self.right_author_var.set("")
        self.right_meta_var.set("")

        self.set_description(
            self.left_description,
            (
                "The adaptive ranking engine has "
                "determined that additional comparisons "
                "are unlikely to materially improve the "
                "current ordering."
            ),
        )

        self.set_description(
            self.right_description,
            (
                "You can continue refining by turning "
                "off 'Stop when stable', or export the "
                "current ranking to Excel."
            ),
        )

        self.progress_var.set(
            max(
                95.0,
                self.engine.progress()
                * 100.0,
            )
        )

        diagnostics = self.engine.diagnostics()
        self.stats_var.set(
            (
                f"{len(self.engine.comparisons)} decisions · "
                f"Top-10 {diagnostics['top10_confidence'] * 100:.0f}% · "
                f"Top-25 {diagnostics['top25_confidence'] * 100:.0f}% · "
                f"boundary unresolved {diagnostics['unresolved_boundary']} · "
                f"{diagnostics['phase']}"
            )
        )

        self.status_var.set(
            "Ranking complete. Export when ready."
        )

    def current_pair_valid(self):
        if (
            self.engine is None
            or self.current_pair is None
        ):
            return False

        left, right = (
            self.current_pair
        )

        if left not in self.engine.ratings:
            return False

        if right not in self.engine.ratings:
            return False

        if left == right:
            return False

        return (
            self.engine.pair_key(
                left,
                right,
            )
            not in self.engine.played
        )

    # ========================================================
    # DISPLAY BOOK
    # ========================================================

    def show_book(
        self,
        book,
        rating,
        title_var,
        author_var,
        meta_var,
        description_widget,
    ):
        title_var.set(
            book.title
        )

        author_var.set(
            book.author
            if book.author
            else "Unknown author"
        )

        meta = []

        if book.pages:
            meta.append(
                f"{book.pages} pages"
            )

        if book.year:
            meta.append(
                f"Published {book.year}"
            )

        stats = self.engine.book_stats(book.id) if self.engine is not None else None
        if stats is not None:
            meta.append(f"Current #{stats['rank']}")
            meta.append(f"Top 25 {stats['top25_probability'] * 100:.0f}%")
            meta.append(f"Likely {stats['rank_low']}–{stats['rank_high']}")

        meta.append(f"Rating {rating.rating:.0f}")
        meta.append(f"±{rating.rd:.0f}")
        meta.append(f"{rating.comparisons} comparisons")

        if book.my_rating:
            meta.append(
                f"My Goodreads: {book.my_rating}"
            )

        meta_var.set(
            "  ·  ".join(meta)
        )

        description = (
            book.description
            if book.description
            else "No description available."
        )

        self.set_description(
            description_widget,
            description,
        )

    def set_description(self, widget, text):
        widget.configure(
            state="normal",
            bg=self.colors["surface"],
            fg="#B9BFCC",
            insertbackground=self.colors["text"],
            selectbackground=self.colors["accent_dark"],
            selectforeground="#FFFFFF",
        )

        widget.delete(
            "1.0",
            "end",
        )

        widget.insert(
            "1.0",
            truncate(text, 1800),
        )

        widget.configure(
            state="disabled",
        )

    # ========================================================
    # CHOICES
    # ========================================================

    def choose(
        self,
        result: str,
    ):
        if (
            self.engine is None
            or self.current_pair is None
        ):
            return

        left_id, right_id = (
            self.current_pair
        )

        try:
            self.engine.apply_match(
                left_id,
                right_id,
                result,
            )

            self.save_state()

        except Exception as exc:
            messagebox.showerror(
                "Could not record choice",
                str(exc),
            )

            return

        self.current_pair = None

        self.refresh()

    def skip(self):
        if self.engine is None or self.current_pair is None:
            return
        left, right = self.current_pair
        self.engine.skip_pair(left, right)
        self.current_pair = None
        self.refresh()

    # ========================================================
    # FINISH
    # ========================================================

    def finish_now(self):
        if self.engine is None:
            return

        self.engine.fit_bradley_terry()
        self.engine.calculate_stability()

        self.save_state()

        self.current_pair = None

        self.show_finished()

    # ========================================================
    # SAVE
    # ========================================================

    def save_state(self):
        if (
            self.engine is None
            or self.state_store is None
        ):
            return

        try:
            self.state_store.save(
                self.engine
            )

        except Exception as exc:
            print(
                "Warning: could not save state:",
                exc,
            )

    # ========================================================
    # UNDO
    # ========================================================

    def undo(self):
        if self.engine is None:
            return

        if self.engine.undo():
            self.save_state()

            self.current_pair = None

            self.refresh()

    # ========================================================
    # BUTTON STATES
    # ========================================================

    def enable_choices(self):
        self.left_button.configure(
            state="normal"
        )

        self.tie_button.configure(
            state="normal"
        )

        self.right_button.configure(
            state="normal"
        )

    def disable_choices(self):
        self.left_button.configure(
            state="disabled"
        )

        self.tie_button.configure(
            state="disabled"
        )

        self.right_button.configure(
            state="disabled"
        )

    # ========================================================
    # RANKING WINDOW
    # ========================================================

    def show_ranking(self):
        if self.engine is None:
            messagebox.showinfo("No ranking", "Open a Goodreads file first.")
            return
        self.engine.ensure_model()
        window = tk.Toplevel(self.root)
        window.configure(bg=self.colors["bg"])
        window.title("Current Ranking · Top-K uncertainty")
        window.geometry("1450x760")
        frame = tk.Frame(
            window,
            bg=self.colors["bg"],
        )
        frame.pack(
            fill="both",
            expand=True,
            padx=18,
            pady=18,
        )
        frame.pack(fill="both", expand=True)
        d = self.engine.diagnostics()
        ttk.Label(frame, text="Current Ranking", style="Title.TLabel").pack(anchor="w")
        ttk.Label(frame, text=(
            f"{self.engine.mode} · Top-25 confidence {d['top25_confidence']*100:.0f}% · "
            f"Top-25 stability {d['top25_stability']*100:.0f}% · "
            f"{d['unresolved_boundary']} unresolved boundary items · "
            f"~{d['estimated_additional']} additional decisions estimated"
        ), foreground="#666666").pack(anchor="w", pady=(0, 8))
        columns = ("rank","title","author","rating","rd","expected","interval","top10","top25","comparisons","record","status")
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        headings = {
            "rank":"Rank","title":"Title","author":"Author","rating":"Model","rd":"RD",
            "expected":"Likely rank","interval":"10–90% rank","top10":"Top 10 %","top25":"Top 25 %",
            "comparisons":"Decisions","record":"W/L/T","status":"Status"
        }
        widths = {"rank":55,"title":320,"author":200,"rating":80,"rd":65,"expected":85,"interval":105,"top10":85,"top25":85,"comparisons":75,"record":75,"status":150}
        for c in columns:
            tree.heading(c, text=headings[c]); tree.column(c, width=widths[c], anchor="w" if c in {"title","author","status"} else "center")
        scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side="left", fill="both", expand=True); scroll.pack(side="right", fill="y")
        for rank, item in enumerate(self.engine.ranking(), start=1):
            book = item["book"]; rating = item["rating"]; st = self.engine.book_stats(book.id)
            status = "TOP 10" if rank <= 10 else ("TOP 25" if rank <= 25 else ("BORDERLINE" if st["top25_probability"] >= .20 else "Long tail"))
            if rating.rd >= 220 or st["rank_high"] - st["rank_low"] >= self.engine.config.boundary_width:
                status += " · uncertain"
            record = f"{rating.wins}/{rating.losses}/{rating.ties}"
            tree.insert("", "end", values=(rank, book.title, book.author, f"{item['final_rating']:.0f}", f"±{rating.rd:.0f}",
                                              st["expected_rank"], f"{st['rank_low']}–{st['rank_high']}",
                                              f"{st['top10_probability']*100:.0f}%", f"{st['top25_probability']*100:.0f}%",
                                              rating.comparisons, record, status))

    # ========================================================
    # DIAGNOSTICS WINDOW
    # ========================================================

    def show_diagnostics(self):
        if self.engine is None:
            messagebox.showinfo("No ranking", "Open a Goodreads file first.")
            return
        self.engine.ensure_model()
        d = self.engine.diagnostics()
        window = tk.Toplevel(self.root)
        window.configure(bg=self.colors["bg"])
        window.title("Top-K Ranking Diagnostics")
        window.geometry("780x720")
        frame = ttk.Frame(window, padding=20);
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Ranking Diagnostics", style="Title.TLabel").pack(anchor="w", pady=(0, 12))
        rows = [
            ("Books", d["books"]), ("Human decisions", d["comparisons"]),
            ("Average decisions/book", f"{d['average_comparisons']:.2f}"),
            ("Evidence coverage", f"{d['coverage']*100:.1f}%"),
            ("Mode", d["mode"]), ("Phase", d["phase"]),
            ("Top-10 membership confidence", f"{d['top10_confidence']*100:.1f}%"),
            ("Top-25 membership confidence", f"{d['top25_confidence']*100:.1f}%"),
            ("Top-10 stability", f"{d['top10_stability']*100:.1f}%"),
            ("Top-25 stability", f"{d['top25_stability']*100:.1f}%"),
            ("Unresolved Top-25 boundary", d["unresolved_boundary"]),
            ("Estimated additional decisions", d["estimated_additional"]),
            ("Simulation count", self.engine.analysis().get("simulations", 0)),
            ("Seed", d["seed"]),
        ]
        for label, value in rows:
            row=ttk.Frame(frame); row.pack(fill="x", pady=3)
            ttk.Label(row, text=label, width=34).pack(side="left")
            ttk.Label(row, text=str(value)).pack(side="left")
        ttk.Separator(frame).pack(fill="x", pady=14)
        ttk.Label(frame, text="How this version allocates evidence", font=("Segoe UI",13,"bold")).pack(anchor="w", pady=(0,8))
        explanation=(
            "The system does not try to perfectly rank the long tail. It first establishes broad evidence, "
            "then builds a bounded elite pool and concentrates comparisons on the Top 10, Top 25 boundary, "
            "high-uncertainty challengers, and pairs whose outcome can materially change the important ordering.\n\n"
            "Top-K probabilities and rank intervals are Monte-Carlo estimates from the current rating and uncertainty model. "
            "They are deliberately labelled estimates, not guarantees. Goodreads ratings and metadata are never used to decide which book you prefer.\n\n"
            "A ranking is considered complete only when Top-25 membership, Top-10 membership, simulation stability, and the boundary are sufficiently resolved."
        )
        ttk.Label(frame, text=explanation, justify="left", wraplength=720).pack(anchor="w")

    # ========================================================
    # EXPORT
    # ========================================================

    def export(self):
        if (
            self.engine is None
            or self.source_file is None
        ):
            messagebox.showinfo(
                "Nothing to export",
                "Open a Goodreads file first.",
            )

            return

        try:
            # Always make sure final model is current.
            self.engine.fit_bradley_terry()
            self.engine.calculate_stability()

            output_path = export_results(
                self.source_file,
                self.engine,
            )

            self.save_state()

        except PermissionError:
            messagebox.showerror(
                "Could not save",
                (
                    "Windows could not save the ranked "
                    "workbook.\n\n"
                    "If the output file is open in Excel, "
                    "close it and try again."
                ),
            )

            return

        except Exception as exc:
            messagebox.showerror(
                "Export failed",
                str(exc),
            )

            return

        open_file = messagebox.askyesno(
            "Ranking exported",
            (
                "Ranking exported successfully.\n\n"
                f"{output_path}\n\n"
                "Open it now?"
            ),
        )

        if open_file:
            self.open_external(
                output_path
            )

    @staticmethod
    def open_external(
        path: Path,
    ):
        try:
            if sys.platform.startswith(
                "win"
            ):
                os.startfile(
                    str(path)
                )

            elif sys.platform == "darwin":
                subprocess.Popen(
                    [
                        "open",
                        str(path),
                    ]
                )

            else:
                subprocess.Popen(
                    [
                        "xdg-open",
                        str(path),
                    ]
                )

        except Exception:
            pass

    # ========================================================
    # HELP
    # ========================================================

    def show_shortcuts(self):
        messagebox.showinfo(
            "Keyboard shortcuts",
            (
                "1   Choose left book\n"
                "2   Choose right book\n"
                "3   Tie / equal\n"
                "4   Skip pair\n"
                "←   Choose left book\n"
                "→   Choose right book\n"
                "T   Tie / equal\n"
                "S   Skip pair\n"
                "U   Undo\n"
                "Ctrl+O   Open Goodreads Excel\n"
                "Ctrl+E   Export ranking"
            ),
        )

    def show_about(self):
        messagebox.showinfo(
            "About",
            (
                f"{APP_NAME}\n"
                f"Version {APP_VERSION}\n\n"
                "A large-library adaptive pairwise "
                "ranking engine for Goodreads shelves.\n\n"
                "ENGINE:\n"
                "• Glicko-2 live uncertainty tracking\n"
                "• Active-learning pair selection\n"
                "• Swiss-style initial exploration\n"
                "• Regularized Bradley-Terry global ranking\n"
                "• Adaptive early stopping\n"
                "• Ranking stability analysis\n"
                "• Top-ranking refinement\n"
                "• Goodreads My Rating as a weak prior\n\n"
                "PERFORMANCE:\n"
                "• Cached comparison indexes\n"
                "• No repeated pair scans\n"
                "• Fast undo snapshots\n"
                "• Sparse comparison fitting\n"
                "• Efficient Excel export lookup\n\n"
                "Only rows where \"Exclusive Shelf\" "
                "equals \"to-read\" are included.\n\n"
                "Your original Goodreads workbook is "
                "never modified."
            ),
        )

    # ========================================================
    # CLOSE
    # ========================================================

    def on_close(self):
        self.save_state()
        self.root.destroy()