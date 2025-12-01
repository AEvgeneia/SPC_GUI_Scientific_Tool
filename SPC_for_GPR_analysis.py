import os
import io
import sys
import csv
import platform
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkfont
from ttkthemes import ThemedTk
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt
import pandas as pd
from dataframe_for_GPR_analysis import DataframeForAnalysis 


# ---------------------- Font defaults per OS ---------------------- #
system_os = platform.system()
if system_os == "Windows":
    FONT_FAMILY = "Segoe UI"
    FONT_SIZE = 9
    TEXT_FONT = ("Consolas", 9)
elif system_os == "Darwin":
    FONT_FAMILY = "San Francisco"
    FONT_SIZE = 12
    TEXT_FONT = ("Menlo", 12)
else:
    FONT_FAMILY = "Ubuntu"
    FONT_SIZE = 11
    TEXT_FONT = ("DejaVu Sans Mono", 11)


# ============================ OOP App ============================ #
class SPCApp:
    """
    Main controller: holds shared state and creates tabs.
    In Step 1 we implement only the Import tab.
    """
    def __init__(self, root: tk.Tk):
        self.root = root
        self.df_soc: pd.DataFrame | None = None
        self.file_path: str | None = None

        self._configure_root()
        self._create_notebook()
        self._create_tabs()

    # ---------- root & theme ---------- #
    def _configure_root(self):
        self.root.title("GPR Statistical Process Control Analysis")
        # DPI awareness (Windows)
        if sys.platform == "win32":
            try:
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                pass

        # Window size
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = min(1000, screen_width - 570)
        window_height = min(800, screen_height - 120)
        self.root.geometry(f"{window_width}x{window_height}+50+30")

        # Fonts
        default_font = tkfont.nametofont("TkDefaultFont")
        default_font.configure(family=FONT_FAMILY, size=FONT_SIZE)
        self.root.option_add("*Font", default_font)
        self.default_font = default_font

        # Closing
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _create_notebook(self):
        self.tab_control = ttk.Notebook(self.root)
        self.tab_control.pack(expand=True, fill="both")

    def _create_tabs(self):
        self.import_tab = ImportTab(self.tab_control, app=self)
        self.tab_control.add(self.import_tab.frame, text="Import File")

        self.analysis_tab = AnalysisTab(self.tab_control, app=self)
        self.tab_control.add(self.analysis_tab.frame, text="Analyze", state="disabled")

        self.spc_tab = SPCTab(self.tab_control, app=self)
        self.tab_control.add(self.spc_tab.frame, text="Statistical Process Control", state="disabled")

        self.tab_control.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    # ---------- Commands used by tabs ---------- #
    def load_file(self):
        """File dialog + DataFrame_soc load. Updates ImportTab UI."""
        path = filedialog.askopenfilename(
            filetypes=[("Data Files", "*.xlsx *.xls *.csv"),
                    ("Excel Files", "*.xlsx *.xls"),
                    ("CSV Files", "*.csv")]
        )
        if not path:
            return

        # Always reset UI before loading a new file
        self.clear_previous_data()

        try:
            # ================= LOAD ================= #
            self.df_soc = DataframeForAnalysis.from_file(path)
            self.file_path = path

            # --- Import tab ---
            self.import_tab.on_file_loaded(os.path.basename(path))
            # Ensure head/tail views show first 5 rows after load
            if self.df_soc is not None:
                self.import_tab.show_head()
                self.import_tab.show_tail()

            # --- ANALYZE tab ---
            idx_analysis = self.tab_control.index(self.analysis_tab.frame)
            self.tab_control.tab(idx_analysis, state="normal")
            self.analysis_tab.enable()  # activates summary + histogram UI

            # Always show Statistics after loading a file
            try:
                self.analysis_tab.sub.select(0)      # Make Statistics sub-tab active
                self.analysis_tab.show_statistics()  # Populate Stats area
            except Exception as e:
                print("[DEBUG] Failed to initialize statistics:", e)

            if hasattr(self.analysis_tab, "_enable_all_tabs"):
                self.analysis_tab._enable_all_tabs()

            # --- SPC tab ---
            idx_spc = self.tab_control.index(self.spc_tab.frame)
            self.tab_control.tab(idx_spc, state="normal")
            self.spc_tab.update_checkboxes(os.path.basename(path))
            
            # Automatically select Shewhart sub-tab after each successful load
            try:
                self.spc_tab.sub_tabs.select(0)  # 0 = Shewhart
            except Exception as e:
                print("[DEBUG] Failed to reset SPC sub-tab:", e)

            # --- Optional warning handling ---
            if hasattr(self.df_soc, "load_warnings") and self.df_soc.load_warnings:
                messagebox.showwarning(
                    "Missing Columns",
                    "Column(s) with problem:\n\n" + "\n".join(self.df_soc.load_warnings)
                )
            # --- Force head/tail spinboxes & views back to 5 rows --- #
            try:
                self.import_tab.head_row_count.set(5)
                self.import_tab.tail_row_count.set(5)
                if hasattr(self.import_tab, "show_head"):
                    self.import_tab.show_head()
                if hasattr(self.import_tab, "show_tail"):
                    self.import_tab.show_tail()
            except Exception as e:
                print("[DEBUG] Failed to refresh head/tail after reload:", e)

        # ================= ERROR HANDLING ================= #
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")

            # --- Import tab ---
            self.import_tab.on_load_failed()

            # --- Disable Analyze tab ---
            try:
                self.analysis_tab.file_label.config(text="❌ Failed to load file", foreground="red")
                self.analysis_tab._disable_all_tabs()
                idx_analysis = self.tab_control.index(self.analysis_tab.frame)
                self.tab_control.tab(idx_analysis, state="disabled")
            except Exception:
                pass

            # --- Disable SPC tab ---
            try:
                self.spc_tab.on_load_failed()
                for method in self.spc_tab.methods:
                    frame = self.spc_tab.checkbox_frames.get(method)
                    if frame and hasattr(frame, "inner_frame"):
                        for w in frame.inner_frame.winfo_children():
                            w.destroy()
                        frame.vars_dict.clear()
                    container = self.spc_tab.plot_containers.get(method)
                    if container:
                        for w in container.winfo_children():
                            w.destroy()
                idx_spc = self.tab_control.index(self.spc_tab.frame)
                self.tab_control.tab(idx_spc, state="disabled")
            except Exception:
                pass

            # --- Reset stored data ---
            self.df_soc = None
            self.file_path = None


    def clear_previous_data(self):
        """Clear old text, plots, and checkboxes before loading a new file."""
        # --- Import Tab ---
        try:
            self.import_tab.file_label.config(text="", fg="green")
            for widget in [self.import_tab.summary_output,
                        self.import_tab.info_output,
                        self.import_tab.head_output,
                        self.import_tab.tail_output]:
                widget.config(state='normal')
                widget.delete('1.0', tk.END)
                widget.config(state='disabled')
           
            # --- Force Head/Tail spinboxes to show "5" visibly and logically --- #
            try:
                def _reset_spinbox(spinbox, var):
                    """Force visible and internal reset to 5."""
                    var.set(5)
                    spinbox.config(state="normal")
                    spinbox.delete(0, "end")
                    spinbox.insert(0, "5")
                    spinbox.update_idletasks()

                if hasattr(self.import_tab, "head_output") and hasattr(self.import_tab.head_output, "_spinbox"):
                    _reset_spinbox(self.import_tab.head_output._spinbox, self.import_tab.head_output._spinvar)
                if hasattr(self.import_tab, "tail_output") and hasattr(self.import_tab.tail_output, "_spinbox"):
                    _reset_spinbox(self.import_tab.tail_output._spinbox, self.import_tab.tail_output._spinvar)

                self.import_tab.head_row_count.set(5)
                self.import_tab.tail_row_count.set(5)

                for i in range(4):
                    self.import_tab.sub.tab(i, state="disabled")

            except Exception as e:
                print("[DEBUG] Failed to reset spinboxes:", e)

        except Exception:
            pass

        # --- Analysis Tab ---
        try:
            self.analysis_tab.file_label.config(text="", foreground="green")

            # Clear histograms + Anderson checkboxes
            for frame in [self.analysis_tab.checkbox_frame,
                        self.analysis_tab.checkbox_frame_anderson]:
                if hasattr(frame, "inner_frame"):
                    for w in frame.inner_frame.winfo_children():
                        w.destroy()
                    frame.vars_dict.clear()

            # Clear histogram plot area
            for w in self.analysis_tab.canvas_container.winfo_children():
                w.destroy()

            # Clear Anderson tree
            for item in self.analysis_tab.tree.get_children():
                self.analysis_tab.tree.delete(item)
        except Exception:
            pass

        # --- SPC Tab ---
        try:
            # close any open popups (summary/outliers)
            for w in list(self.spc_tab.open_windows):
                try:
                    if w.winfo_exists(): w.destroy()
                except Exception:
                    pass
            self.spc_tab.open_windows.clear()
            self.spc_tab.window_counters = {"summary": 0}

            for method in self.spc_tab.methods:
                # clear filename label
                key = method + "_file_label"
                if key in self.spc_tab.method_tabs:
                    self.spc_tab.method_tabs[key].config(text="", foreground="green")

                # clear checkboxes
                frame = self.spc_tab.checkbox_frames.get(method)
                if frame and hasattr(frame, "inner_frame"):
                    for w in frame.inner_frame.winfo_children():
                        w.destroy()
                    frame.vars_dict.clear()

                # clear plots
                container = self.spc_tab.plot_containers.get(method)
                if container:
                    for w in container.winfo_children():
                        w.destroy()
        except Exception:
            pass


    def _on_close(self):
        if messagebox.askokcancel("Quit", "Do you really want to exit the application?"):
            plt.close('all')
            self.root.destroy()
            self.root.quit()
            sys.exit(0)

    def _on_tab_changed(self, event):
        """Automatically refresh Analysis tab statistics when selected."""
        selected_tab = event.widget.select()
        tab_text = event.widget.tab(selected_tab, "text")

        if tab_text == "Analyze" and self.df_soc is not None:
            try:
                # Enable all its sub-tabs if not already enabled
                self.analysis_tab._enable_all_tabs()

                # Only refresh if user is CURRENTLY on Statistics sub-tab
                current_sub = self.analysis_tab.sub.index("current")
                if current_sub == 0:  # 0 = Statistics
                    self.analysis_tab.show_statistics()

            except Exception as e:
                print(f"[DEBUG] Failed to refresh statistics: {e}")



# ============================ Import Tab ============================ #
class ImportTab:
    """
    Handles: load file, and the four sub-tabs (Summary/Info/Head/Tail).
    """
    def __init__(self, parent_notebook: ttk.Notebook, app: SPCApp):
        self.app = app
        self.frame = ttk.Frame(parent_notebook)

        self._build_header()
        self._build_subtabs()

    # ---------- layout ---------- #
    def _build_header(self):
        ttk.Label(
            self.frame,
            text="Upload a file to inspect its contents. Note: the details shown here correspond to the original DataFrame."
        ).pack(pady=10)

        self.warning_label = ttk.Label(
            self.frame,
            text="⚠️ Use the recommended Excel layout to ensure proper functionality.",
            foreground="red",
            font=(FONT_FAMILY, max(FONT_SIZE-1, 8), "italic"),
            wraplength=800,
            justify="center"
        )
        self.warning_label.pack(pady=(0, 5))

        ttk.Button(self.frame, text="📁 Load File", command=self.app.load_file).pack(padx=10, pady=10)

        self.file_label = ttk.Label(self.frame, text="", foreground="green",
                                   font=(FONT_FAMILY, max(FONT_SIZE, 8), "bold"))
        self.file_label.pack(pady=5)

    def _build_subtabs(self):
        self.sub = ttk.Notebook(self.frame)
        self.sub.pack(expand=1, fill="both", padx=10, pady=10)

        self.summary_tab = ttk.Frame(self.sub)
        self.info_tab = ttk.Frame(self.sub)
        self.head_tab = ttk.Frame(self.sub)
        self.tail_tab = ttk.Frame(self.sub)

        self.sub.add(self.summary_tab, text="Summary")
        self.sub.add(self.info_tab, text="Info")
        self.sub.add(self.head_tab, text="Head")
        self.sub.add(self.tail_tab, text="Tail")

        # Disable until file is loaded
        for i in range(4):
            self.sub.tab(i, state="disabled")

        # Text areas
        self.summary_output = self._make_text_area(self.summary_tab)
        self.info_output = self._make_text_area(self.info_tab)

        # Head/Tail with spinboxes
        self.head_row_count = tk.IntVar(value=5)
        self.head_output = self._make_text_area(
            self.head_tab, spin_var=self.head_row_count, spin_cmd=self.show_head
        )

        self._enable_horizontal_scroll(self.head_output)

        self.tail_row_count = tk.IntVar(value=5)
        self.tail_output = self._make_text_area(
            self.tail_tab, spin_var=self.tail_row_count, spin_cmd=self.show_tail
        )

        self._enable_horizontal_scroll(self.tail_output)

    # ---------- composed widgets ---------- #
    def _make_text_area(self, parent, placeholder="", spin_var=None, spin_cmd=None):
        outer = tk.Frame(parent)
        outer.pack(expand=True, fill="both", padx=10, pady=10)

        text_frame = tk.Frame(outer)
        text_frame.pack(expand=True, fill="both")

        x_scroll = tk.Scrollbar(text_frame, orient='horizontal')
        y_scroll = tk.Scrollbar(text_frame, orient='vertical')

        text_widget = tk.Text(
            text_frame,
            wrap="none",
            font=TEXT_FONT,
            xscrollcommand=x_scroll.set,
            yscrollcommand=y_scroll.set,
            bd=0,
            highlightthickness=0,
            relief="flat"
        )
        text_widget.grid(row=0, column=0, sticky="nsew")
        x_scroll.grid(row=1, column=0, sticky="ew")
        x_scroll.config(command=text_widget.xview)
        y_scroll.grid(row=0, column=1, sticky="ns")

        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        if placeholder:
            text_widget.insert(tk.END, placeholder)
            text_widget.config(state='disabled')

        if spin_var is not None and spin_cmd is not None:
            ctrl = tk.Frame(outer)
            ctrl.pack(pady=5)
            tk.Label(ctrl, text="Rows to show:").pack(side="left")
            spin = tk.Spinbox(ctrl, from_=1, to=100, textvariable=spin_var, width=5, command=spin_cmd)
            spin.pack(side="left", padx=5)
            # Keep handles on the text widget for later resets
            text_widget._spinbox = spin
            text_widget._spinvar = spin_var
        
        text_widget.update_idletasks()
        return text_widget
    
    def _enable_horizontal_scroll(self, text_widget: tk.Text):
        """Enable horizontal scrolling with the mouse wheel anywhere over the text area."""
        import platform

        def _on_mousewheel(event):
            sysname = platform.system()
            if sysname == "Darwin":               # macOS
                step = -5 if event.delta > 0 else 5
                text_widget.xview_scroll(step, "units")
            else:                                 # Windows / Linux
                step = int(-event.delta / 120)
                text_widget.xview_scroll(step * 7, "units")

        # Linux touchpad events
        def _on_btn4(_): text_widget.xview_scroll(-7, "units")
        def _on_btn5(_): text_widget.xview_scroll(7, "units")

        def _bind_all(_):
            text_widget.bind_all("<MouseWheel>", _on_mousewheel)
            text_widget.bind_all("<Button-4>", _on_btn4)
            text_widget.bind_all("<Button-5>", _on_btn5)

        def _unbind_all(_):
            text_widget.unbind_all("<MouseWheel>")
            text_widget.unbind_all("<Button-4>")
            text_widget.unbind_all("<Button-5>")

        # Activate bindings while the pointer is over the widget
        text_widget.bind("<Enter>", _bind_all)
        text_widget.bind("<Leave>", _unbind_all)


    # ---------- public entry points ---------- #
    def on_file_loaded(self, filename: str):
        self.warning_label.pack_forget()
        self.file_label.config(text=f"Loaded: {filename}", foreground="green")

        # Enable tabs
        for i in range(4):
            self.sub.tab(i, state="normal")

        # Populate all views
        self.show_summary()
        self.show_info()
        self.show_head()
        self.show_tail()

        # Focus Summary
        self.sub.select(0)

    def on_load_failed(self):
        self.file_label.config(text="❌ Failed to load file", foreground="red")
        for i in range(4):
            self.sub.tab(i, state="disabled")
        for widget in [self.summary_output, self.info_output, self.head_output, self.tail_output]:
            widget.config(state='normal')
            widget.delete('1.0', tk.END)
            widget.config(state='disabled')

    # ---------- renderers ---------- #
    def show_summary(self):
        out = self.summary_output
        out.config(state='normal')
        out.delete('1.0', tk.END)
        if self.app.df_soc is not None:
            data = self.app.df_soc.get_summary_data()
            site_list = data.get('Site of Cancer', [])
            site_str = ", ".join(site_list) if isinstance(site_list, list) else str(site_list)
            out.insert(tk.END, f"📌 Site(s) of Cancer: {site_str}\n\n")
            out.insert(tk.END, f"📐 Shape: {data.get('Shape','')}\n\n")
            out.insert(tk.END, f"📆 Date Range: {data.get('Date Range','')}\n\n")
            if "sorted data" in data:
                out.insert(tk.END, f"⏳ QA Date Sorting: {data['sorted data']}\n\n")
            out.insert(tk.END, "🧾 Columns:\n")
            for col in data.get('Columns', []):
                out.insert(tk.END, f"   • {col}\n")
            out.insert(tk.END, "\n🧾 Data for the SPC Analysis:\n")
            for col in data.get('Data for analysis', []):
                out.insert(tk.END, f"   • {col}\n")
        out.config(state='disabled')

    def show_info(self):
        out = self.info_output
        out.config(state='normal')
        out.delete('1.0', tk.END)
        if self.app.df_soc is not None:
            buffer = io.StringIO()
            self.app.df_soc.info(buf=buffer)
            out.insert(tk.END, buffer.getvalue())
        out.config(state='disabled')

    def show_head(self):
        out = self.head_output
        out.config(state='normal')
        out.delete('1.0', tk.END)
        if self.app.df_soc is not None:
            n = self.head_row_count.get()
            out.insert(tk.END, self.app.df_soc.head(n).to_string())
        out.config(state='disabled')

    def show_tail(self):
        out = self.tail_output
        out.config(state='normal')
        out.delete('1.0', tk.END)
        if self.app.df_soc is not None:
            n = self.tail_row_count.get()
            out.insert(tk.END, self.app.df_soc.tail(n).to_string())
        out.config(state='disabled')

# ============================ Analysis Tab ============================ #
class AnalysisTab:
    """
    Handles Tab 2: summary statistics, histograms, and normality tests.
    """
    def __init__(self, parent_notebook: ttk.Notebook, app: SPCApp):
        self.app = app
        self.frame = ttk.Frame(parent_notebook)

        self._build_layout()
        self._disable_all_tabs()

    # ---------- layout ---------- #
    def _build_layout(self):
        ttk.Label(self.frame, text="Summary statistics and distribution check for the initial and SPC-processed DataFrames.").pack(pady=10)
        self.file_label = ttk.Label(self.frame, text="", foreground="green")
        self.file_label.pack(pady=5)

        # Create sub-tabs
        self.sub = ttk.Notebook(self.frame)
        self.sub.pack(expand=1, fill="both", padx=10, pady=10)

        self.stats_tab = ttk.Frame(self.sub)
        self.hists_tab = ttk.Frame(self.sub)
        self.anderson_tab = ttk.Frame(self.sub)

        self.sub.add(self.stats_tab, text="Statistics")
        self.sub.add(self.hists_tab, text="Histograms")
        self.sub.add(self.anderson_tab, text="Normality Test")

        # === Statistics output === #
        self.stats_output = self._make_text_area(self.stats_tab)
        self._enable_horizontal_scroll(self.stats_output)
        self.sub.bind("<<NotebookTabChanged>>", self._on_subtab_changed)

        # === Histograms === #
        self._build_hist_tab()

        # === Anderson–Darling === #
        self._build_anderson_tab()

    def _on_subtab_changed(self, event):
        try:
            current_tab = self.sub.tab(self.sub.select(), "text")
            if current_tab == "Statistics" and self.app.df_soc is not None:
                self.show_statistics()
        except Exception:
            pass

    def _build_hist_tab(self):
        frame = ttk.Frame(self.hists_tab)
        frame.pack(pady=10)

        ttk.Label(frame, text="Select QA metrics to plot:").pack()

        # ✅ Use our new checkbox grid instead of a listbox
        self.checkbox_frame = self._create_checkbox_grid(frame)

        ttk.Button(frame, text="📊 Plot Selected Histograms",
                command=self.show_histograms).pack(pady=5)


        # Scrollable plot area
        self.plot_canvas = tk.Canvas(self.hists_tab, bg="white", highlightthickness=0)
        self.scrollbar = tk.Scrollbar(self.hists_tab, orient="vertical",
                                      command=self.plot_canvas.yview)
        self.plot_canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        self.plot_canvas.pack(side="left", expand=True, fill="both")

        self.canvas_container = ttk.Frame(self.plot_canvas)
        self.plot_canvas.create_window((0, 0), window=self.canvas_container, anchor="nw")
        self.canvas_container.bind("<Configure>",
                                   lambda e: self.plot_canvas.configure(scrollregion=self.plot_canvas.bbox("all")))

        # 🆕 Enable scroll anywhere over histograms
        self._bind_mousewheel(self.plot_canvas, self.canvas_container)

    def _build_anderson_tab(self):
        frame = ttk.Frame(self.anderson_tab)
        frame.pack(pady=10)

        ttk.Label(frame, text="Select QA metrics to test:").pack()

        # ✅ Use our new checkbox grid here too
        self.checkbox_frame_anderson = self._create_checkbox_grid(frame)

        ttk.Button(frame, text="Run Anderson–Darling Test",
                command=self.run_anderson_test).pack(pady=5)


        # Treeview
        style = ttk.Style()
        style.configure("Treeview", font=(FONT_FAMILY, FONT_SIZE))
        style.configure("Treeview.Heading", font=(FONT_FAMILY, FONT_SIZE, "bold"))

        self.tree = ttk.Treeview(
            self.anderson_tab,
            columns=("GPR", "Statistic", "Critical", "Normality"),
            show="headings"
        )
        self.tree.heading("GPR", text="GPR")
        self.tree.heading("Statistic", text="A² Statistic")
        self.tree.heading("Critical", text="Critical Value (5%)")
        self.tree.heading("Normality", text="Normality")
        self.tree.pack(expand=True, fill="both", padx=10, pady=10)

        ttk.Button(self.anderson_tab, text="💾 Save as CSV",
                   command=self.save_anderson_to_csv).pack(pady=5)

    # ---------- small widget helpers ---------- #

    def _enable_horizontal_scroll(self, text_widget: tk.Text):
        """Enable horizontal scrolling with the mouse wheel anywhere over the text area."""
        import platform

        def _on_mousewheel(event):
            sysname = platform.system()
            if sysname == "Darwin":               # macOS
                step = -5 if event.delta > 0 else 5
                text_widget.xview_scroll(step, "units")
            else:                                 # Windows / Linux
                step = int(-event.delta / 120)*7
                text_widget.xview_scroll(step, "units")

        # Linux touchpad events
        def _on_btn4(_): text_widget.xview_scroll(-7, "units")
        def _on_btn5(_): text_widget.xview_scroll(7, "units")

        def _bind_all(_):
            text_widget.bind_all("<MouseWheel>", _on_mousewheel)
            text_widget.bind_all("<Button-4>", _on_btn4)
            text_widget.bind_all("<Button-5>", _on_btn5)

        def _unbind_all(_):
            text_widget.unbind_all("<MouseWheel>")
            text_widget.unbind_all("<Button-4>")
            text_widget.unbind_all("<Button-5>")

        # Activate bindings while the pointer is over the widget
        text_widget.bind("<Enter>", _bind_all)
        text_widget.bind("<Leave>", _unbind_all)


    def _create_checkbox_grid(self, parent, columns=3):
        """
        Create a scrollable grid of checkboxes (up to 5 per row)
        with centered 'Select All / None' buttons underneath.
        """
        # Outer container
        outer_frame = ttk.Frame(parent)
        outer_frame.pack(pady=(5, 0), fill="x", expand=True)

        # === Scrollable canvas area === #
        canvas_frame = ttk.Frame(outer_frame)
        canvas_frame.pack(fill="both", expand=True)

        canvas = tk.Canvas(canvas_frame, height=120, width=625, highlightthickness=0, borderwidth=0)
        canvas.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Frame inside the canvas
        inner_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=inner_frame, anchor="nw")

        # Update scroll region dynamically
        inner_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        vars_dict = {}

        # === Buttons centered UNDER the grid === #
        button_frame = ttk.Frame(outer_frame)
        button_frame.pack(pady=(8, 4))
        ttk.Button(button_frame, text="Select All",
                command=lambda: self._select_all(vars_dict, True)).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Select None",
                command=lambda: self._select_all(vars_dict, False)).pack(side="left", padx=10)

        # Store references
        outer_frame.inner_frame = inner_frame
        outer_frame.vars_dict = vars_dict
        outer_frame.columns = columns

        return outer_frame

    
    def _select_all(self, vars_dict, state: bool):
        for var in vars_dict.values():
            var.set(state)


    def _make_text_area(self, parent):
        outer = tk.Frame(parent)
        outer.pack(expand=True, fill="both", padx=10, pady=10)

        # Frame for text and scrollbars
        text_frame = tk.Frame(outer)
        text_frame.pack(expand=True, fill="both")

        # Scrollbars
        x_scroll = tk.Scrollbar(text_frame, orient='horizontal')
        y_scroll = tk.Scrollbar(text_frame, orient='vertical')

        # Text widget
        txt = tk.Text(
            text_frame,
            wrap="none",
            font=TEXT_FONT,
            xscrollcommand=x_scroll.set,
            yscrollcommand=y_scroll.set,
            bd=0,
            highlightthickness=0,
            relief="flat"
        )

        # Layout
        txt.grid(row=0, column=0, sticky="nsew")
        x_scroll.grid(row=1, column=0, sticky="ew")
        y_scroll.grid(row=0, column=1, sticky="ns")

        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        x_scroll.config(command=txt.xview)
        y_scroll.config(command=txt.yview)

        return txt

    # ---------- public control ---------- #
    def enable(self):
        idx = self.app.tab_control.index(self.frame)
        self.app.tab_control.tab(idx, state="normal")
        self.file_label.config(text=f"Loaded: {os.path.basename(self.app.file_path)}", foreground="green")
        self._enable_all_tabs()
        self.update_listboxes()
        self.show_statistics()

    # ---------- tab state ---------- #
    def _enable_all_tabs(self):
        for i in range(3):
            self.sub.tab(i, state="normal")

    def _disable_all_tabs(self):
        for i in range(3):
            self.sub.tab(i, state="disabled")

    # ---------- logic ---------- #
    def update_listboxes(self):
        """Populate checkbox grids with numeric analysis columns."""
        
        if self.app.df_soc is None:
            return

        # Clear previous checkboxes
        for frame in [self.checkbox_frame, self.checkbox_frame_anderson]:
            for widget in frame.inner_frame.winfo_children():
                widget.destroy()
            frame.vars_dict.clear()

        numeric_cols = self.app.df_soc[
            self.app.df_soc.columns.intersection(self.app.df_soc.data_for_analysis)
        ].select_dtypes(include='number').columns

        rows_per_column = 4  # ✅ you can change this to 4 or 6 later if needed

        for idx, col in enumerate(numeric_cols):
            colnum, row = divmod(idx, rows_per_column)

            # Histogram tab checkboxes
            var1 = tk.BooleanVar(value=False)
            cb1 = ttk.Checkbutton(self.checkbox_frame.inner_frame, text=col, variable=var1)
            cb1.grid(row=row, column=colnum, sticky="w", padx=14, pady=4)
            self.checkbox_frame.vars_dict[col] = var1

            # Anderson–Darling tab checkboxes
            var2 = tk.BooleanVar(value=False)
            cb2 = ttk.Checkbutton(self.checkbox_frame_anderson.inner_frame, text=col, variable=var2)
            cb2.grid(row=row, column=colnum, sticky="w", padx=14, pady=4)
            self.checkbox_frame_anderson.vars_dict[col] = var2



    def show_statistics(self):
        out = self.stats_output
        out.config(state='normal')
        out.delete('1.0', tk.END)
        if self.app.df_soc is not None:
            out.insert(tk.END, self.app.df_soc.get_statistics())
        out.config(state='disabled')

    def show_histograms(self):
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

        for w in self.canvas_container.winfo_children():
            w.destroy()

        df = self.app.df_soc
        if df is None:
            return

        selected = [col for col, var in self.checkbox_frame.vars_dict.items() if var.get()]
        if not selected:
            messagebox.showwarning("No Selection", "Please select at least one QA metric.")
            return

        try:
            system_os = platform.system()
            figs = df.plot_histograms_gui(selected_columns=selected, return_fig=True)
            for fig in figs:
                canvas = FigureCanvasTkAgg(fig, master=self.canvas_container)
                canvas.draw()
                canvas.get_tk_widget().pack(expand=True, fill="both")

            self.plot_canvas.yview_moveto(0)

            ttk.Button(self.canvas_container, text="💾 Save All to PDF",
                       command=lambda: self.save_histograms_to_pdf(selected)).pack(pady=10)
        except Exception as e:
            messagebox.showerror("Error", f"Could not display histograms:\n{e}")

    def save_histograms_to_pdf(self, selected):
        df = self.app.df_soc
        if df is None:
            return
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                 filetypes=[("PDF files", "*.pdf")])
        if not file_path:
            return
        try:
            with PdfPages(file_path) as pdf:
                df.plot_histograms_gui(selected_columns=selected, pdf=pdf)
            messagebox.showinfo("Success", f"Saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def run_anderson_test(self):
        df = self.app.df_soc
        if df is None:
            return
        selected = [col for col, var in self.checkbox_frame_anderson.vars_dict.items() if var.get()]

        if not selected:
            messagebox.showwarning("No Selection", "Select at least one QA metric.")
            return
        try:
            for i in self.tree.get_children():
                self.tree.delete(i)
            results = df.run_anderson_test(selected)
            for row in results:
                self.tree.insert("", "end", values=(
                    row["GPR"], f"{row['Statistic']:.3f}",
                    f"{row['Critical Value (5%)']:.3f}", row["Normality"]
                ))
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def save_anderson_to_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV files", "*.csv")])
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["GPR", "A² Statistic", "Critical Value (5%)", "Normality"])
                for child in self.tree.get_children():
                    writer.writerow(self.tree.item(child)["values"])
            messagebox.showinfo("Saved", f"Results saved to {path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _bind_mousewheel(self, canvas: tk.Canvas, area: tk.Widget):
        """Enable scrolling with mouse wheel anywhere over the histogram area."""
        import platform

        def _on_mousewheel(event):
            sysname = platform.system()
            if getattr(event, "state", 0) & 0x0001:  # Shift for horizontal scroll
                if sysname == "Darwin":
                    step = -2 if event.delta > 0 else 2
                    canvas.xview_scroll(step, "units")
                else:
                    step = int(-event.delta / 120)
                    canvas.xview_scroll(step * 3, "units")
                return

            if sysname == "Darwin":  # macOS
                step = -2 if event.delta > 0 else 2
                canvas.yview_scroll(step, "units")
            else:  # Windows/Linux
                step = int(-event.delta / 120)
                canvas.yview_scroll(step * 3, "units")

        def _on_btn4(_): canvas.yview_scroll(-3, "units")
        def _on_btn5(_): canvas.yview_scroll(3, "units")

        def _bind_all(_):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_btn4)
            canvas.bind_all("<Button-5>", _on_btn5)

        def _unbind_all(_):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        area.bind("<Enter>", _bind_all)
        area.bind("<Leave>", _unbind_all)

# ============================ SPC Tab ============================ #
class SPCTab:
    """
    Handles Tab 3: Statistical Process Control (SPC) Analysis.
    Includes Shewhart, WSD, SC, and SWV methods.
    """
    def __init__(self, parent_notebook: ttk.Notebook, app):
        self.app = app
        self.frame = ttk.Frame(parent_notebook)
        self.open_windows = []

        self.summary_window_counters = {
            "shewhart": 1,
            "wsd": 1,
            "sc": 1,
            "swv": 1
        }

        # Create sub-tabs for each SPC method
        self.sub_tabs = ttk.Notebook(self.frame)
        self.sub_tabs.pack(expand=1, fill="both", padx=10, pady=10)

        # Keep references for each method tab
        self.methods = ["shewhart", "wsd", "sc", "swv"]
        self.method_names = {
            "shewhart": "Shewhart",
            "wsd": "Weighted Standard Deviation",
            "sc": "Skewness Correction",
            "swv": "Scaled Weighted Variance"
        }
        self.method_tabs = {}
        self.checkbox_frames = {}
        self.plot_containers = {}

        for method in self.methods:
            self._build_spc_tab(method)

        # Bind method change
        self.active_method = tk.StringVar(value="shewhart")
        self.sub_tabs.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # ONE unified elimination log
        self.elimination_log = []

        # Shared counter for the whole SPC session
        self.elimination_round = 1

    def _build_spc_tab(self, method: str):
        """Create a single SPC method tab with checkbox grid and plot area."""
        tab = ttk.Frame(self.sub_tabs)
        pretty = self.method_names[method]
        self.sub_tabs.add(tab, text=pretty)
        self.method_tabs[method] = tab

        ttk.Label(tab, text=f"Select QA metrics for {method.upper()} SPC analysis:").pack(pady=10)

        file_label = ttk.Label(tab, text="", foreground="green", font=(FONT_FAMILY, FONT_SIZE))
        file_label.pack(pady=(0, 6))
        self.method_tabs[method + "_file_label"] = file_label

        # Checkbox grid
        checkbox_frame = self._create_checkbox_grid(tab)
        self.checkbox_frames[method] = checkbox_frame

        ttk.Button(
            tab,
            text=f"▶ Run {method.upper()} SPC",
            command=lambda m=method: self.run_spc_analysis(m)
        ).pack(pady=(4, 6))
        
        # === Scrollable plot area === #
        plot_area = ttk.Frame(tab)
        plot_area.pack(expand=True, fill="both")

        plot_canvas = tk.Canvas(plot_area, bg="white", highlightthickness=0, borderwidth=0)
        scrollbar = tk.Scrollbar(plot_area, orient="vertical", command=plot_canvas.yview)
        plot_canvas.configure(yscrollcommand=scrollbar.set)
        plot_canvas.pack(side="left", expand=True, fill="both")
        scrollbar.pack(side="right", fill="y")

        container = tk.Frame(plot_canvas, bg="white")
        plot_canvas.create_window((0, 0), window=container, anchor="nw")
        container.bind("<Configure>", lambda e: plot_canvas.configure(scrollregion=plot_canvas.bbox("all")))

        self.plot_containers[method] = container

        self._bind_mousewheel(plot_canvas, container)

        exit_frame = ttk.Frame(tab)
        exit_frame.pack(pady=10)
        ttk.Button(
            exit_frame,
            text="❌ Exit SPC Session",
            command=self.exit_spc_session
        ).pack()

    def _bind_mousewheel(self, canvas: tk.Canvas, area: tk.Widget):
        """Enable scrolling with the mouse wheel anywhere over the plot area.

        We bind on <Enter>/<Leave> of the plot 'area' (the frame that holds figures).
        While the pointer is inside, route all wheel events to the canvas.
        """

        import platform

        def _on_mousewheel(event):
            sysname = platform.system()
            if sysname == "Darwin":           # macOS: delta is already small steps
                step = -2 if event.delta > 0 else 2
                canvas.yview_scroll(step, "units")
            else:                              # Windows/Linux: multiples of 120
                step = int(-event.delta / 120)
                canvas.yview_scroll(step * 3, "units")


        # Linux touchpads / X11 button events
        def _on_btn4(_): canvas.yview_scroll(-3, "units")
        def _on_btn5(_): canvas.yview_scroll( 3, "units")

        def _bind_all(_):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_btn4)
            canvas.bind_all("<Button-5>", _on_btn5)

        def _unbind_all(_):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        # Enter/leave the whole scrollable area (including figures)
        area.bind("<Enter>", _bind_all)
        area.bind("<Leave>", _unbind_all)

    def _create_checkbox_grid(self, parent, columns=5):
        outer_frame = ttk.Frame(parent)
        outer_frame.pack(pady=(5, 0), fill="x", expand=False)

        canvas_frame = ttk.Frame(outer_frame)
        canvas_frame.pack(fill="none", expand=False)

        canvas = tk.Canvas(canvas_frame, height=120, width=625, highlightthickness=0, borderwidth=0)
        canvas.pack(side="left", fill="both", expand=False)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=scrollbar.set)

        inner_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=inner_frame, anchor="n")
        inner_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        vars_dict = {}
        button_frame = ttk.Frame(outer_frame)
        button_frame.pack(pady=(8, 4))
        ttk.Button(button_frame, text="Select All", command=lambda: self._select_all(vars_dict, True)).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Select None", command=lambda: self._select_all(vars_dict, False)).pack(side="left", padx=10)

        outer_frame.inner_frame = inner_frame
        outer_frame.vars_dict = vars_dict
        outer_frame.columns = columns
        return outer_frame

    def _select_all(self, vars_dict, state: bool):
        for var in vars_dict.values():
            var.set(state)

    def update_checkboxes(self, filename=None):
        """Populate all SPC method checkbox grids with available numeric columns."""
        if self.app.df_soc is None:
            return
        
        if filename:
            self.update_spc_file_labels(filename)

        numeric_cols = self.app.df_soc[
            self.app.df_soc.columns.intersection(self.app.df_soc.data_for_analysis)
        ].select_dtypes(include='number').columns

        rows_per_column = 4
        for method in self.methods:
            frame = self.checkbox_frames[method]
            for w in frame.inner_frame.winfo_children():
                w.destroy()
            frame.vars_dict.clear()

            for idx, col in enumerate(numeric_cols):
                colnum, row = divmod(idx, rows_per_column)
                var = tk.BooleanVar(value=False)
                cb = ttk.Checkbutton(frame.inner_frame, text=col, variable=var)
                cb.grid(row=row, column=colnum, sticky="w", padx=14, pady=4)
                frame.vars_dict[col] = var

    def update_spc_file_labels(self, filename):
        """Update 'Loaded: filename' label in all SPC method tabs."""
        for method in self.methods:
            key = method + "_file_label"
            if key in self.method_tabs:
                self.method_tabs[key].config(text=f"Loaded: {filename}", foreground="green")

    def on_load_failed(self):
        """Display a red 'Failed to load file' message on all SPC method tabs."""
        for method in self.methods:
            key = method + "_file_label"
            if key in self.method_tabs:
                self.method_tabs[key].config(text="❌ Failed to load file", foreground="red")


    def run_spc_analysis(self, method):
        for win in list(self.open_windows):
            try:
                if win.winfo_exists() and "Select Outliers" in win.title():
                    win.destroy()
                    self.open_windows.remove(win)
            except:
                pass
        df = self.app.df_soc
        if df is None:
            messagebox.showwarning("No data", "Please load a dataset first.")
            return

        selected = [col for col, var in self.checkbox_frames[method].vars_dict.items() if var.get()]
        if not selected:
            messagebox.showwarning("No Selection", f"Please select at least one QA metric for {method.upper()}.")
            return

        # Clear previous plots
        container = self.plot_containers[method]
        for w in container.winfo_children():
            w.destroy()

        # Call the right DataFrame_soc method
        method_func = {
            "shewhart": df.get_shewhart_x_chart_figs,
            "wsd": df.get_wsd_x_chart_figs,
            "sc": df.get_sc_x_chart_figs,
            "swv": df.get_swv_x_chart_figs
        }.get(method)

        if not method_func:
            messagebox.showerror("Error", f"Unknown SPC method: {method}")
            return

        try:
            figs, outlier_dict, results_list = method_func(selected_columns=selected)
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

            for fig in figs:
                fig.patch.set_facecolor("white")
                fig.patch.set_alpha(1.0)

                frame = tk.Frame(container, bg="white", highlightthickness=0, bd=0)
                frame.pack(expand=True, fill="both", padx=10, pady=5)
                canvas = FigureCanvasTkAgg(fig, master=frame)
                canvas.draw()
                canvas.get_tk_widget().pack(expand=True, fill="both")
            
            plot_canvas = container.master
            if isinstance(plot_canvas, tk.Canvas):
                plot_canvas.yview_moveto(0)
        
            stats_window = self.show_spc_stats(results_list, method)

            # open outlier selection
            if outlier_dict and any(outlier_dict[k] for k in outlier_dict):
                self.open_outliers_window(outlier_dict, method)
            else:
                messagebox.showinfo("No Outliers", f"No outliers detected in {method.upper()} SPC for the selected QA metrics.")
                # Restore summary window visibility exactly after messagebox closes
                stats_window.after_idle(lambda: (stats_window.lift(), stats_window.focus_force()))

        except Exception as e:
            messagebox.showerror("Error", f"Failed to run {method.upper()} SPC:\n{e}")
        
        ttk.Button(container, text="💾 Save All Plots to PDF",
           command=lambda: self.save_spc_plots_to_pdf(figs, method)).pack(pady=8)
        
    def save_spc_plots_to_pdf(self, figs, method):
        path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                            filetypes=[("PDF files", "*.pdf")],
                                            title=f"Save {method.upper()} SPC Plots")
        if not path:
            return
        try:
            with PdfPages(path) as pdf:
                for fig in figs:
                    pdf.savefig(fig, bbox_inches="tight")
            messagebox.showinfo("Saved", f"All {method.upper()} SPC plots saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _on_tab_changed(self, event):
        selected_tab = event.widget.select()
        tab_text = event.widget.tab(selected_tab, "text").strip().lower()
        mapping = {
            "shewhart": "shewhart",
            "weighted standard deviation": "wsd",
            "skewness correction": "sc",
            "scaled weighted variance": "swv"
        }
        self.active_method.set(mapping.get(tab_text, "shewhart"))

    def show_spc_stats(self, results_list, method):
        """Display SPC statistics summary in a popup window."""
        stats_window = tk.Toplevel(self.frame)
        self.open_windows.append(stats_window)
        count = self.summary_window_counters.get(method, 1)

        stats_window.title(f"{method.upper()} SPC Summary #{count}")
        #stats_window.geometry("+100+100")
        count = self.summary_window_counters.get(method, 1)
        x_offset = 50 + (count - 1) * 20
        y_offset = 80 + (count - 1) * 20
        stats_window.geometry(f"+{x_offset}+{y_offset}")
        self.summary_window_counters[method] = count + 1

        style = ttk.Style()
        style.configure("Treeview", font=(FONT_FAMILY, FONT_SIZE))
        style.configure("Treeview.Heading", font=(FONT_FAMILY, FONT_SIZE, "bold"))

        columns = ["GPR", "Mean", "Count", "LCL", "UCL", "LSL", "USL", "Outliers"]
        tree = ttk.Treeview(stats_window, columns=columns, show="headings")
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="center", width=100)
        tree.column("GPR", width=180)

        for row in results_list:
            tree.insert("", "end", values=(
                row["GPR Column"],
                row["Mean (X̄)"],
                row["Counts"],
                row["LCL"],
                row["UCL"],
                row["LSL"] if row["LSL"] else "-",
                row["USL"] if row["USL"] else "-",
                "Yes" if row["Out-of-Control IDs"] else "No"
            ))

        tree.pack(expand=True, fill="both", padx=10, pady=10)

        ttk.Button(stats_window, text="💾 Save as CSV",
                command=lambda: self.save_summary_to_csv(tree, method)).pack(pady=6)
        return stats_window
    
    def save_summary_to_csv(self, tree, method):
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV files", "*.csv")],
                                            title=f"Save {method.upper()} SPC Summary")
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow([c for c in tree["columns"]])
            for child in tree.get_children():
                writer.writerow(tree.item(child)["values"])
        messagebox.showinfo("Saved", f"Summary saved to {path}")

    def open_outliers_window(self, outlier_dict, method):
        """Dialog to select and eliminate outliers for the given SPC method."""
        win = tk.Toplevel(self.frame)
        self.open_windows.append(win)
        win.title(f"Select Outliers for {method.upper()} SPC")
        win.geometry("+900+600")
        win.lift()
        win.focus_force()

        tk.Label(win, text="Select QA metric:").grid(row=0, column=0, sticky="w")
        
        # Keep only metrics that actually have outliers
        valid_metrics = [k for k, v in outlier_dict.items() if v]

        criterion_combo = ttk.Combobox(
            win,
            values=valid_metrics,
            state="readonly",
            width=30
        )
        criterion_combo.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(win, text="Select IDs to eliminate:").grid(row=1, column=0, sticky="nw")
        id_listbox = tk.Listbox(win, selectmode="multiple", height=10, exportselection=False)
        id_listbox.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")

        def on_select_metric(event):
            metric = criterion_combo.get()
            id_listbox.delete(0, tk.END)
            if metric:
                for _id in outlier_dict[metric]:
                    id_listbox.insert(tk.END, _id)

        criterion_combo.bind("<<ComboboxSelected>>", on_select_metric)

        def eliminate():
            metric = criterion_combo.get()
            selected = [id_listbox.get(i) for i in id_listbox.curselection()]
            if not metric or not selected:
                messagebox.showwarning("Missing Selection", "Select a metric and at least one ID.")
                return
            round_num = self.elimination_round
            log = self.app.df_soc.elimination_recalculate_gui(
                method=method,
                confidence_level="99.73%",
                selected_criterion=metric,
                selected_ids=selected,
                round_num=round_num
            )

            self.elimination_round += 1

            # Attach method name to each entry + append to unified log
            for entry in log:
                entry_list = list(entry)  # entry is [round, criterion, id, value]
                entry_list.insert(1, method)  # => [round, method, criterion, id, value]
                self.elimination_log.append(entry_list)
            messagebox.showinfo("Success", "Outliers eliminated and data recalculated.")
            win.destroy()
            # rerun analysis with updated data
            self.run_spc_analysis(method)

            container = self.plot_containers[method]
            plot_canvas = container.master
            if isinstance(plot_canvas, tk.Canvas):
                plot_canvas.yview_moveto(0)
        
        ttk.Button(win, text="Eliminate Selected IDs", command=eliminate).grid(row=2, column=1, pady=10, sticky="e")
    
    def exit_spc_session(self):
        """Reset SPC analysis state and restore original data."""
        if self.app.file_path is None or self.app.df_soc is None:
            return

        if messagebox.askyesno("Reset SPC", "Do you want to reset and reload the original data?"):
            self.app.df_soc = DataframeForAnalysis.from_file(self.app.file_path)
            for method in self.methods:
                container = self.plot_containers[method]
                for w in container.winfo_children():
                    w.destroy()
                for var in self.checkbox_frames[method].vars_dict.values():
                    var.set(False)

            if self.elimination_log:
                save = messagebox.askyesno(
                    "Save Eliminations?",
                    "Do you want to save the elimination log?"
                )

                if save:
                    df = pd.DataFrame(
                        self.elimination_log,
                        columns=["#Elimination", "Method", "Criterion", "ID", "Eliminated Value"]
                    )

                    save_path = filedialog.asksaveasfilename(
                        defaultextension=".csv",
                        filetypes=[("CSV files", "*.csv")],
                        title="Save Elimination Log"
                    )

                    if save_path:
                        df.to_csv(save_path, index=False, encoding="utf-8-sig")
                        messagebox.showinfo("Saved", f"Elimination log saved to:\n{save_path}")

                self.elimination_log = []
                self.elimination_round = 1

            for w in list(self.open_windows):
                try:
                    if w.winfo_exists():
                        w.destroy()
                except Exception:
                    pass
            self.open_windows.clear()

            self.update_spc_file_labels(os.path.basename(self.app.file_path))
            self.app.spc_tab.update_checkboxes(os.path.basename(self.app.file_path))
            plt.close('all')
            self.window_counters = {"summary": 0}
            messagebox.showinfo("SPC Reset", "Original dataset reloaded.")


# ============================ main ============================ #
if __name__ == "__main__":
    root = ThemedTk(theme="arc")
    app = SPCApp(root)
    root.mainloop()
