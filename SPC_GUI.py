import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from ttkthemes import ThemedTk
import tkinter.font as font
from dataframe_SOC import DataFrame_soc  # custom class
import io
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt
import tkinter.font as tkfont
import csv
import pandas as pd
import traceback
import sys
import platform
from tkinter import font

system_os = platform.system()
if system_os == "Windows":
    font_family = "Segoe UI"
    font_size = 9
    text_font = ("Consolas", 9)
elif system_os == "Darwin":
    font_family = "San Francisco"  # macOS system font
    font_size = 12
    text_font = ("Menlo", 12)
else:
    font_family = "Ubuntu"
    font_size = 11
    text_font = ("DejaVu Sans Mono", 11)

# ==================== Globals ==================== #
df_soc = None
hist_canvas = None
file_path = None
summary_window_counters = {
    "shewhart": 1,
    "swv": 1,
    "wsd": 1,
    "sc": 1
}

elimination_logs = {}

elimination_round_counters = {}

outlier_window_ref = None


# =================================================== #
# ==================== Functions ==================== #
# =================================================== #

def load_file():
    global df_soc  # When I assign to df_soc in this function, I want to change the global variable — not create a new local one.
    # filedialog.askopenfilename() returns the full path of the file
    global file_path
    file_path = filedialog.askopenfilename(
        filetypes=[("Data Files", "*.xlsx *.xls *.csv"), ("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv")])
    if file_path:  # Only continue if the user actually selected a file.
        try:
            df_soc = DataFrame_soc.from_file(file_path)
            warning_label.pack_forget()

            filename = os.path.basename(file_path)
            file_label.config(text=f"Loaded: {filename}", fg="green")
            tab_control.tab(1, state="normal")
            update_hist_listbox()
            file_label_2.config(text=f"Loaded: {filename}", fg="green")
            tab_control.tab(2, state="normal")
            file_label_3.config(text=f"Loaded: {filename}", fg="green")

            for i in range(4):
                sub_tabs_tab1.tab(i, state="normal")

            show_summary()
            show_info()
            show_head()
            show_tail()
            show_statistics()

            sub_tabs_tab1.select(0)

            if hasattr(df_soc, "load_warnings") and df_soc.load_warnings:
                messagebox.showwarning(
                    "Missing Columns",
                    "Column(s) with problem:\n\n" + "\n".join(df_soc.load_warnings)
                )

        except Exception as e:  # catch all the valueError from my class
            # with open("error_log.txt", "a", encoding="utf-8") as f:
            # f.write("Error during file load:\n")
            # f.write(traceback.format_exc())
            # f.write("\n\n")
            messagebox.showerror("Error", f"Failed to load file:\n{e}")
            file_label.config(text="❌ Failed to load file", fg="red")
            df_soc = None

            for i in range(4):
                sub_tabs_tab1.tab(i, state="disabled")
            tab_control.tab(1, state="disabled")
            tab_control.tab(2, state="disabled")
            for widget in [summary_output, info_output, head_output, tail_output]:
                widget.config(state='normal')
                widget.delete('1.0', tk.END)
                widget.config(state='disabled')


def create_text_area_with_scrollbars_and_spinbox(
        parent_tab,
        placeholder="",
        label_text="Rows to show:",
        spinbox_var=None,
        spinbox_command=None,
        font=text_font
):
    # Outer frame
    outer_frame = tk.Frame(parent_tab)
    outer_frame.pack(expand=True, fill="both", padx=10, pady=10)

    # Frame for text + scrollbars
    text_frame = tk.Frame(outer_frame)
    text_frame.pack(expand=True, fill="both")

    # Scrollbars
    x_scroll = tk.Scrollbar(text_frame, orient='horizontal')
    y_scroll = tk.Scrollbar(text_frame, orient='vertical')

    # Text widget
    text_widget = tk.Text(
        text_frame,
        wrap="none",
        font=font,
        xscrollcommand=x_scroll.set,
        yscrollcommand=y_scroll.set,
        bd=0,
        highlightthickness=0,
        relief="flat"
    )

    # Layout
    text_widget.grid(row=0, column=0, sticky="nsew")
    x_scroll.grid(row=1, column=0, sticky="ew")
    x_scroll.config(command=text_widget.xview)
    y_scroll.grid(row=0, column=1, sticky="ns")

    text_frame.rowconfigure(0, weight=1)
    text_frame.columnconfigure(0, weight=1)

    # Placeholder text
    if placeholder:
        text_widget.insert(tk.END, placeholder)
        text_widget.config(state='disabled')

    # Optional bottom controls (label + spinbox)
    if spinbox_var is not None and spinbox_command is not None:
        control_frame = tk.Frame(outer_frame)
        control_frame.pack(pady=5)

        tk.Label(control_frame, text=label_text).pack(side="left")
        tk.Spinbox(
            control_frame,
            from_=1,
            to=100,
            textvariable=spinbox_var,
            width=5,
            command=spinbox_command
        ).pack(side="left", padx=5)

    text_widget.update_idletasks()
    return text_widget


def show_summary():
    summary_output.config(state='normal')
    summary_output.delete('1.0', tk.END)
    if df_soc is not None:
        data = df_soc.get_summary_data()
        site_list = data['Site of Cancer']
        if isinstance(site_list, list):
            site_str = ", ".join(site_list)
        else:
            site_str = str(site_list)
        summary_output.insert(tk.END, f"📌 Site(s) of Cancer: {site_str}\n\n")
        summary_output.insert(tk.END, f"📐 Shape: {data['Shape']}\n\n")
        summary_output.insert(tk.END, f"📆 Date Range: {data['Date Range']}\n\n")
        if "sorted data" in data:
            summary_output.insert(tk.END, f"⏳ QA Date Sorting: {data['sorted data']}\n\n")
        summary_output.insert(tk.END, "🧾 Columns:\n")
        for col in data['Columns']:
            summary_output.insert(tk.END, f"   • {col}\n")
        summary_output.insert(tk.END, "\n🧾 Data for the SPC Analysis:\n")
        for col in data['Data for analysis']:
            summary_output.insert(tk.END, f"   • {col}\n")
    summary_output.config(state='disabled')


def show_info():
    info_output.config(state='normal')
    info_output.delete('1.0', tk.END)
    if df_soc is not None:
        buffer = io.StringIO()
        df_soc.info(buf=buffer)
        info_output.insert(tk.END, buffer.getvalue())
    info_output.config(state='disabled')


def show_head():
    head_output.config(state='normal')
    head_output.delete('1.0', tk.END)
    if df_soc is not None:
        n = head_row_count.get()
        head_output.insert(tk.END, df_soc.head(n).to_string())
    head_output.config(state='disabled')


def show_tail():
    tail_output.config(state='normal')
    tail_output.delete('1.0', tk.END)
    if df_soc is not None:
        n = tail_row_count.get()
        tail_output.insert(tk.END, df_soc.tail(n).to_string())
    tail_output.config(state='disabled')


def show_statistics():
    stats_output.config(state='normal')
    stats_output.delete('1.0', tk.END)
    if df_soc is not None:
        stats_output.insert(tk.END, df_soc.get_statistics())
    stats_output.config(state='disabled')


def show_histograms():
    global hist_canvas
    for widget in hist_canvas_container.winfo_children():
        widget.destroy()

    if df_soc is not None:
        try:
            selected_indices = listbox.curselection()
            selected_columns = [listbox.get(i) for i in selected_indices]
            if not selected_columns:
                messagebox.showwarning("No Selection", "Please select at least one column.")
                return

            figs = df_soc.plot_histograms_gui(
                selected_columns=selected_columns,
                return_fig=True,
            )

            if not figs:
                messagebox.showerror("Error", "No histograms were generated.")
                return

            for fig in figs:
                canvas = FigureCanvasTkAgg(fig, master=hist_canvas_container)
                canvas.draw()
                canvas.get_tk_widget().pack(expand=True, fill='both')

            # 🆕 Add save button
            ttk.Button(hist_canvas_container, text="💾 Save All Histograms to One PDF",
                       command=lambda: save_all_histograms_to_pdf(selected_columns)).pack(pady=10)

        except Exception as e:
            messagebox.showerror("Error", f"Could not display histograms:\n{e}")


def save_all_histograms_to_pdf(selected_columns):
    if df_soc is None or not selected_columns:
        messagebox.showwarning("No Data", "Please load data and select columns.")
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        title="Save All Histograms as PDF"
    )

    if not file_path:
        return

    try:
        with PdfPages(file_path) as pdf:
            df_soc.plot_histograms_gui(selected_columns=selected_columns, pdf=pdf)
        messagebox.showinfo("Success", f"Histograms saved to:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save histograms:\n{e}")


def update_hist_listbox():
    if df_soc is not None:
        listbox.delete(0, tk.END)
        listbox_anderson.delete(0, tk.END)
        shewhart_refs["listbox"].delete(0, tk.END)
        wsd_refs["listbox"].delete(0, tk.END)
        sc_refs["listbox"].delete(0, tk.END)
        swv_refs["listbox"].delete(0, tk.END)
        numeric_cols = df_soc[df_soc.columns.intersection(df_soc.data_for_analysis)].select_dtypes(
            include='number').columns
        for col in numeric_cols:
            listbox.insert(tk.END, col)
            listbox_anderson.insert(tk.END, col)
            shewhart_refs["listbox"].insert(tk.END, col)
            wsd_refs["listbox"].insert(tk.END, col)
            sc_refs["listbox"].insert(tk.END, col)
            swv_refs["listbox"].insert(tk.END, col)


def run_anderson_test():
    if df_soc is not None:
        selected_indices = listbox_anderson.curselection()
        selected_columns = [listbox_anderson.get(i) for i in selected_indices]
        if not selected_columns:
            messagebox.showwarning("No Selection", "Please select at least one column.")
            return
        try:
            # Clear previous treeview results
            for item in anderson_tree.get_children():
                anderson_tree.delete(item)

            # Example expected result structure:
            # [
            #     {"Column": "Dose", "Statistic": 0.61, "Critical Value": 0.56, "Normal?": "No"},
            #     {"Column": "Error", "Statistic": 0.31, "Critical Value": 0.56, "Normal?": "Yes"},
            # ]
            result = df_soc.run_anderson_test(selected_columns)

            for row in result:
                anderson_tree.insert("", "end", values=(
                    row["GPR"],
                    f"{row['Statistic']:.3f}",
                    f"{row['Critical Value (5%)']:.3f}",
                    row["Normality"]
                ))


        except Exception as e:
            messagebox.showerror("Error", f"Failed to run Anderson test:\n{e}")


def save_anderson_to_csv():
    file_path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv")],
        title="Save Anderson-Darling Results"
    )
    if not file_path:
        return

    try:
        with open(file_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["GPR", "A² Statistic", "Critical Value (5%)", "Normality"])  # headers

            for child in anderson_tree.get_children():
                row = anderson_tree.item(child)["values"]
                writer.writerow(row)

        messagebox.showinfo("Success", f"Anderson-Darling results saved to:\n{file_path}")
    except Exception as e:
        messagebox.showerror("Failed to save file", f"{e}")


def setup_spc_tab(tab, run_command, listbox_ref_dict, canvas_ref_dict, label="Run SPC"):
    # Column selection frame
    column_select_frame = ttk.Frame(tab)
    column_select_frame.pack(pady=10)

    listbox, listbox_frame = create_listbox_with_scrollbar(column_select_frame)
    listbox_frame.pack()

    ttk.Button(column_select_frame, text=label, command=run_command).pack(pady=(5, 0))

    # Scrollable plot area
    plot_frame = ttk.Frame(tab)
    plot_frame.pack(expand=True, fill="both")

    canvas = tk.Canvas(plot_frame)
    scrollbar = tk.Scrollbar(plot_frame, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", expand=True, fill="both")

    canvas_container = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=canvas_container, anchor="nw")

    def on_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    canvas_container.bind("<Configure>", on_configure)

    # Save widget references for outside use
    listbox_ref_dict["listbox"] = listbox
    canvas_ref_dict["canvas"] = canvas
    canvas_ref_dict["container"] = canvas_container


def run_spc_test(method, refs, canvas_refs):
    global df_soc

    method_tab_indices = {
        "shewhart": 0,
        "wsd": 1,
        "sc": 2,
        "swv": 3
    }

    current_tab_index = method_tab_indices[method]

    for i in range(4):
        if i != current_tab_index:
            sub_tabs.tab(i, state="disabled")

    sub_tabs.select(current_tab_index)

    for widget in canvas_refs["container"].winfo_children():
        widget.destroy()

    if df_soc is not None:
        selected_indices = refs["listbox"].curselection()
        selected_columns = [refs["listbox"].get(i) for i in selected_indices]

        if not selected_columns:
            messagebox.showwarning("No Selection", f"⚠️ Please select at least one column for {method.upper()}.")
            return

        try:
            method_func = {
                "shewhart": df_soc.get_shewhart_x_chart_figs,
                "wsd": df_soc.get_wsd_x_chart_figs,
                "sc": df_soc.get_sc_x_chart_figs,
                "swv": df_soc.get_swv_x_chart_figs
            }.get(method)

            if method_func is None:
                messagebox.showerror("Error", f"Unsupported SPC method: {method}")
                return

            figs, outlier_dict, results_list = method_func(selected_columns=selected_columns)

            for fig in figs:
                frame = ttk.Frame(canvas_refs["container"])
                frame.pack(expand=True, fill='both', padx=10, pady=5)
                chart = FigureCanvasTkAgg(fig, master=frame)
                chart.draw()
                chart.get_tk_widget().pack(expand=True, fill='both')

            open_outliers_window(outlier_dict, method)
            show_spc_stats(results_list, method)

        except Exception as e:
            messagebox.showerror("Error", f"Could not display charts for {method.upper()}:\n\n{e}")

        def save_all_figures_to_pdf():
            file_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                title="Save All Plots as PDF"
            )

            if not file_path:
                return

            try:
                with PdfPages(file_path) as pdf:
                    for fig in figs:
                        pdf.savefig(fig, bbox_inches="tight")
                messagebox.showinfo("Success", f"All plots saved to:\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save PDF:\n{e}")

        save_button_frame = ttk.Frame(canvas_refs["container"])
        save_button_frame.pack(pady=10)
        ttk.Button(save_button_frame, text="💾 Save All Plots to One PDF", command=save_all_figures_to_pdf).pack()


def exit_spc_method():
    global df_soc, summary_window_counters, elimination_round_counters

    if df_soc is not None and file_path:

        method = active_method.get()

        # Prompt to save elimination log
        if method in elimination_logs and elimination_logs[method]:
            save = messagebox.askyesno(
                "Save Eliminations?",
                f"Do you want to save the elimination log for {method.upper()}?"
            )
            if save:
                log_df = pd.DataFrame(elimination_logs[method],
                                      columns=["#elimination", "Criterion", "ID", "Eliminated Value"])
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".csv",
                    filetypes=[("CSV files", "*.csv")],
                    title="Save Elimination Log"
                )
                if save_path:
                    log_df.to_csv(save_path, index=False, encoding="utf-8-sig")
                    messagebox.showinfo("Saved", f"Elimination log saved to:\n{save_path}")

        # Reset state
        if method in elimination_logs:
            elimination_logs[method] = []

        elimination_round_counters = {}

        df_soc = DataFrame_soc.from_file(file_path)
        messagebox.showinfo("SPC Reset", "The original data has been restored.")

        for i in range(4):
            sub_tabs.tab(i, state="normal")

        for m in summary_window_counters:
            summary_window_counters[m] = 1

        # Close summary window
        for window in root.winfo_children():
            if isinstance(window, tk.Toplevel) and "SPC Summary" in window.title():
                window.destroy()

        # Close outliers window
        for window in root.winfo_children():
            if isinstance(window, tk.Toplevel) and "Select Outliers" in window.title():
                window.destroy()

        # Clear all SPC plots
        for canvas_refs in [shewhart_canvas_refs, wsd_canvas_refs, sc_canvas_refs, swv_canvas_refs]:
            if "container" in canvas_refs:
                for widget in canvas_refs["container"].winfo_children():
                    widget.destroy()


def show_spc_stats(results_list, method):
    global summary_window_counters

    stats_window = tk.Toplevel()
    method_name = "Shewhart" if method == "shewhart" else method.upper()
    count = summary_window_counters.get(method, 1)
    stats_window.title(f"{method_name} SPC Summary {count}")

    x_offset = 50 + (count - 1) * 20
    y_offset = 100 + (count - 1) * 20
    stats_window.geometry(f"+{x_offset}+{y_offset}")
    stats_window.lift()
    stats_window.focus_force()
    stats_window.attributes('-topmost', 1)
    stats_window.after(500, lambda: stats_window.attributes('-topmost', 0))

    summary_window_counters[method] = count + 1

    # Updated fonts for Windows style
    style = ttk.Style()
    style.configure("Treeview", font=(default_font, font_size))
    style.configure("Treeview.Heading", font=(default_font, font_size, "bold"))

    # Define columns
    tree_columns = ["GPR Column", "Mean", "Count", "LCL", "UCL", "LSL", "USL", "Outliers"]

    tree = ttk.Treeview(stats_window, columns=tree_columns, show="headings")

    # Setup headings and column widths
    tree.heading("GPR Column", text="GPR Column")
    tree.column("GPR Column", anchor="center", width=180)

    for col in ["Mean", "Count", "LCL", "UCL", "LSL", "USL", "Outliers"]:
        tree.heading(col, text=col)
        tree.column(col, anchor="center", width=90)

    # Insert rows
    for row in results_list:
        tree.insert("", "end", values=(
            row["GPR Column"],
            row["Mean (X̄)"],
            row["Counts"],
            row["LCL"],
            row["UCL"],
            row["LSL"] if row["LSL"] is not None else "-",
            row["USL"] if row["USL"] is not None else "-",
            "Yes" if row["Out-of-Control IDs"] else "No"
        ))

    tree.pack(expand=True, fill="both", padx=10, pady=10)

    def save_to_csv():
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save SPC Summary"
        )
        if not file_path:
            return

        try:
            with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(tree_columns)  # Header row

                for child in tree.get_children():
                    writer.writerow(tree.item(child)["values"])

            messagebox.showinfo("Success", f"Results saved to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Failed to Save", f"{e}")

    ttk.Button(stats_window, text="💾 Save as CSV", command=save_to_csv).pack(pady=10)


def open_outliers_window(outlier_dict, method):
    global outlier_window_ref

    # outlier_dict: {"criterion1": [ID1, ID2], "criterion2": [ID3, ID4], ...}
    def eliminate_selected_ids():
        criterion = criterion_combo.get()  # from the dropdown
        selected_indices = id_listbox.curselection()
        selected_ids = [id_listbox.get(i) for i in selected_indices]

        if not criterion or not selected_ids:
            messagebox.showwarning("Missing Data", "Please select a criterion and at least one ID.")
            return

        if method not in elimination_logs:
            elimination_logs[method] = []

        if method not in elimination_round_counters:
            elimination_round_counters[method] = 1

        round_num = elimination_round_counters[method]

        # Now call your modified method
        log = df_soc.elimination_recalculate_gui(
            method=method,
            confidence_level="99.73%",
            selected_criterion=criterion,
            selected_ids=selected_ids,
            round_num=round_num
        )

        elimination_round_counters[method] += 1

        elimination_logs[method].extend(log)

        messagebox.showinfo("Success", "Selected outliers have been eliminated.")
        outlier_window.destroy()
        # stats_window.destroy()

        method_refs = {
            "shewhart": (shewhart_refs, shewhart_canvas_refs),
            "wsd": (wsd_refs, wsd_canvas_refs),
            "sc": (sc_refs, sc_canvas_refs),
            "swv": (swv_refs, swv_canvas_refs)
        }

        run_spc_test(method, *method_refs[method])

    def on_criterion_selected(event):
        selected = criterion_combo.get()
        if selected:
            # Populate ID listbox
            id_listbox.delete(0, tk.END)
            for id_ in outlier_dict[selected]:
                id_listbox.insert(tk.END, id_)
            eliminate_button.config(state="disabled")

    def on_id_selected(event):
        # Enable the button only if at least one ID is selected
        if id_listbox.curselection():
            eliminate_button.config(state="normal")
        else:
            eliminate_button.config(state="disabled")

    if outlier_window_ref is not None and outlier_window_ref.winfo_exists():
        outlier_window_ref.destroy()

    # Create a new Toplevel window
    outlier_window = tk.Toplevel()
    outlier_window_ref = outlier_window  # store the new reference
    outlier_window.title("Select Outliers to Eliminate")

    outlier_window.geometry("+900+600")
    outlier_window.lift()
    outlier_window.focus_force()
    outlier_window.attributes('-topmost', 1)
    outlier_window.after(500, lambda: outlier_window.attributes('-topmost', 0))

    # Criterion label + combobox
    tk.Label(outlier_window, text="Select Criterion:").grid(row=0, column=0, sticky="w")
    criterion_combo = ttk.Combobox(outlier_window, values=list(outlier_dict.keys()), state="readonly", width=30)
    criterion_combo.grid(row=0, column=1, padx=5, pady=5)
    criterion_combo.bind("<<ComboboxSelected>>", on_criterion_selected)

    # ID listbox with scrollbar
    tk.Label(outlier_window, text="Select IDs:").grid(row=1, column=0, sticky="nw")
    id_frame = tk.Frame(outlier_window)
    id_scroll = tk.Scrollbar(id_frame, orient="vertical")
    id_listbox = tk.Listbox(id_frame, height=10, selectmode="multiple", yscrollcommand=id_scroll.set,
                            exportselection=False)
    id_scroll.config(command=id_listbox.yview)
    id_scroll.pack(side="right", fill="y")
    id_listbox.pack(side="left", fill="both", expand=True)
    id_frame.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")
    id_listbox.bind("<<ListboxSelect>>", on_id_selected)

    # Eliminate button (initially disabled)
    eliminate_button = ttk.Button(outlier_window, text="Eliminate", state="disabled", command=eliminate_selected_ids)
    eliminate_button.grid(row=2, column=1, pady=10, sticky="e")

    # Optional: configure resizing
    outlier_window.grid_rowconfigure(1, weight=1)
    outlier_window.grid_columnconfigure(1, weight=1)


def create_listbox_with_scrollbar(parent, width=50, height=6, selectmode=tk.MULTIPLE):
    frame = tk.Frame(parent)
    listbox = tk.Listbox(
        frame,
        selectmode=selectmode,
        width=width,
        height=height,
        exportselection=False
    )
    scrollbar = tk.Scrollbar(frame, orient="vertical", command=listbox.yview)
    listbox.config(yscrollcommand=scrollbar.set)

    listbox.pack(side="left", fill="y")
    scrollbar.pack(side="right", fill="y")

    return listbox, frame


# ======================================================== #
# ==================== Main GUI setup ==================== #
# ======================================================== #
if __name__ == "__main__":

    # DPI awareness for Windows
    if sys.platform == "win32":
        try:
            from ctypes import windll

            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass


    def on_closing():
        if messagebox.askokcancel("Quit", "Do you really want to exit the application?"):
            plt.close('all')
            root.destroy()
            root.quit()  # Ensure the mainloop ends completely
            sys.exit(0)


    root = ThemedTk(theme="arc")  # Modern theme that mimics macOS

    # Set consistent font
    default_font = font.nametofont("TkDefaultFont")
    default_font.configure(family=font_family, size=font_size)
    root.option_add("*Font", default_font)

    active_method = tk.StringVar(value="shewhart")
    root.title("Statistical Process Control Analysis")
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    window_width = min(1000, screen_width - 570)
    window_height = min(800, screen_height - 170)

    root.geometry(f"{window_width}x{window_height}")

    # root.geometry("850x600+100+50")

    # ==================== Main Tabs Creation ==================== #
    tab_control = ttk.Notebook(root)

    tab1 = ttk.Frame(tab_control)
    tab2 = ttk.Frame(tab_control)
    tab3 = ttk.Frame(tab_control)

    tab_control.add(tab1, text='Import File')
    tab_control.add(tab2, text='Analyze', state="disabled")
    tab_control.add(tab3, text='Statistical Process Control', state="disabled")

    tab_control.pack(expand=1, fill='both')

    # ==================== Tab 1 layout ==================== #

    # Themed title label
    ttk.Label(
        tab1,
        text="Upload a file to inspect its contents and begin the analysis.",
        font=(default_font, font_size)
    ).pack(pady=10)

    # Warning label (requires tk.Label for color, optional ttk.Label with style)
    warning_label = tk.Label(
        tab1,
        text="⚠️ Use the recommended Excel layout to ensure proper functionality.",
        fg="red",
        font=(default_font, font_size - 1, "italic"),
        wraplength=800,
        justify="center"
    )
    warning_label.pack(pady=(0, 5))

    # Themed button
    ttk.Button(
        tab1,
        text="📁 Load File",
        command=load_file
    ).pack(padx=10, pady=10)

    # File loaded label (you can use tk.Label if color is needed)
    file_label = tk.Label(tab1, text="", fg="green", font=(default_font, font_size - 1, "bold"))
    file_label.pack(pady=5)

    # ==================== Sub-tabs for tab1 ==================== #
    sub_tabs_tab1 = ttk.Notebook(tab1)
    sub_tabs_tab1.pack(expand=1, fill="both", padx=10, pady=10)

    # Use themed frames
    summary_tab = ttk.Frame(sub_tabs_tab1)
    info_tab = ttk.Frame(sub_tabs_tab1)
    head_tab = ttk.Frame(sub_tabs_tab1)
    tail_tab = ttk.Frame(sub_tabs_tab1)

    sub_tabs_tab1.add(summary_tab, text="Summary")
    sub_tabs_tab1.add(info_tab, text="Info")
    sub_tabs_tab1.add(head_tab, text="Head")
    sub_tabs_tab1.add(tail_tab, text="Tail")

    for i in range(4):
        sub_tabs_tab1.tab(i, state="disabled")

    summary_output = create_text_area_with_scrollbars_and_spinbox(summary_tab)
    info_output = create_text_area_with_scrollbars_and_spinbox(info_tab)

    head_row_count = tk.IntVar(value=5)
    head_output = create_text_area_with_scrollbars_and_spinbox(
        head_tab,
        spinbox_var=head_row_count,
        spinbox_command=show_head
    )

    tail_row_count = tk.IntVar(value=5)
    tail_output = create_text_area_with_scrollbars_and_spinbox(
        tail_tab,
        spinbox_var=tail_row_count,
        spinbox_command=show_tail
    )

    # ==================== Tab 2 layout ==================== #
    ttk.Label(tab2, text="Summary statistics and distribution check.", font=(default_font, font_size)).pack(pady=10)

    file_label_2 = tk.Label(tab2, text="", fg="green", font=(default_font, font_size - 1, "bold"))
    file_label_2.pack(pady=5)

    # === Sub-tabs for tab2 === #
    sub_tabs = ttk.Notebook(tab2)
    sub_tabs.pack(expand=1, fill="both", padx=10, pady=10)

    stats_tab = ttk.Frame(sub_tabs)
    hists_tab = ttk.Frame(sub_tabs)
    anderson_tab = ttk.Frame(sub_tabs)

    sub_tabs.add(stats_tab, text="Statistics")
    sub_tabs.add(hists_tab, text="Histograms")
    sub_tabs.add(anderson_tab, text="Normality Test")

    stats_output = create_text_area_with_scrollbars_and_spinbox(stats_tab)

    # == Hists_output == #
    selected_hist_columns = []

    # Widgets
    column_select_frame = ttk.Frame(hists_tab)
    column_select_frame.pack(pady=10)

    # Scrollable plot area
    plot_frame = ttk.Frame(hists_tab)
    plot_frame.pack(expand=True, fill="both")

    canvas = tk.Canvas(plot_frame)
    scrollbar = tk.Scrollbar(plot_frame, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scrollbar.pack(side="right", fill="y")
    canvas.pack(side="left", expand=True, fill="both")

    # Inner frame in canvas
    hist_canvas_container = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=hist_canvas_container, anchor="nw")


    def on_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))


    hist_canvas_container.bind("<Configure>", on_configure)

    ttk.Label(column_select_frame, text="Select columns to plot:").pack()

    listbox, listbox_frame = create_listbox_with_scrollbar(column_select_frame)
    listbox_frame.pack()

    ttk.Button(column_select_frame, text="📊 Plot Selected Histograms", command=show_histograms).pack(pady=5)

    # == Anderson_output == #
    selected_anderson_columns = []

    column_select_frame = ttk.Frame(anderson_tab)
    column_select_frame.pack(pady=10)

    anderson_frame = ttk.Frame(anderson_tab)
    anderson_frame.pack(expand=True, fill="both")

    listbox_anderson, listbox_frame_anderson = create_listbox_with_scrollbar(column_select_frame)
    listbox_frame_anderson.pack()

    ttk.Button(column_select_frame, text="Run Anderson-Darling Test", command=run_anderson_test).pack(pady=5)

    # Treeview (keep styling)
    style = ttk.Style()
    style.configure("Treeview", font=(default_font, font_size))
    style.configure("Treeview.Heading", font=(default_font, font_size, "bold"))

    anderson_tree = ttk.Treeview(
        anderson_frame,
        columns=("GPR", "Statistic", "Critical Value (5%)", "Normality"),
        show="headings"
    )

    anderson_tree.heading("GPR", text="Criterion")
    anderson_tree.heading("Statistic", text="A² Statistic")
    anderson_tree.heading("Critical Value (5%)", text="Critical Value (5%)")
    anderson_tree.heading("Normality", text="Normality")

    anderson_tree.pack(expand=True, fill="both", padx=10, pady=10)

    ttk.Button(anderson_frame, text="💾 Save as CSV", command=save_anderson_to_csv).pack(pady=5)

    # ==================== Tab 3 layout ==================== #
    ttk.Label(tab3, text="Perform Statistical Process Control Analysis.", font=(default_font, font_size)).pack(pady=10)

    file_label_3 = tk.Label(tab3, text="", fg="green", font=(default_font, font_size - 1, "bold"))
    file_label_3.pack(pady=5)

    sub_tabs = ttk.Notebook(tab3)
    sub_tabs.pack(expand=1, fill="both", padx=10, pady=10)

    shewhart_tab = ttk.Frame(sub_tabs)
    wsd_tab = ttk.Frame(sub_tabs)
    sc_tab = ttk.Frame(sub_tabs)
    swv_tab = ttk.Frame(sub_tabs)

    sub_tabs.add(shewhart_tab, text="Shewhart")
    sub_tabs.add(wsd_tab, text="Weighted Standard Deviation")
    sub_tabs.add(sc_tab, text="Skewness Correction")
    sub_tabs.add(swv_tab, text="Scaled Weighted Variance")


    # Bind to tab change
    def on_tab_changed(event):
        selected_tab = event.widget.select()
        tab_text = event.widget.tab(selected_tab, "text").strip().lower()

        mapping = {
            "shewhart": "shewhart",
            "weighted standard deviation": "wsd",
            "skewness correction": "sc",
            "scaled weighted variance": "swv"
        }
        active_method.set(mapping.get(tab_text, "shewhart"))


    sub_tabs.bind("<<NotebookTabChanged>>", on_tab_changed)

    # == Shewhart ==
    shewhart_refs = {}
    shewhart_canvas_refs = {}

    setup_spc_tab(
        shewhart_tab,
        run_command=lambda: run_spc_test("shewhart", shewhart_refs, shewhart_canvas_refs),
        listbox_ref_dict=shewhart_refs,
        canvas_ref_dict=shewhart_canvas_refs,
        label="Run Shewhart SPC"
    )

    ttk.Button(shewhart_tab, text="❌ Exit SPC Session", command=exit_spc_method).pack(pady=5)

    # == Weighted Standard Deviation ==
    wsd_refs = {}
    wsd_canvas_refs = {}

    setup_spc_tab(
        wsd_tab,
        run_command=lambda: run_spc_test("wsd", wsd_refs, wsd_canvas_refs),
        listbox_ref_dict=wsd_refs,
        canvas_ref_dict=wsd_canvas_refs,
        label="Run WSD SPC"
    )

    ttk.Button(wsd_tab, text="❌ Exit SPC Session", command=exit_spc_method).pack(pady=5)

    # == Skewness Correction ==
    sc_refs = {}
    sc_canvas_refs = {}

    setup_spc_tab(
        sc_tab,
        run_command=lambda: run_spc_test("sc", sc_refs, sc_canvas_refs),
        listbox_ref_dict=sc_refs,
        canvas_ref_dict=sc_canvas_refs,
        label="Run SC SPC"
    )

    ttk.Button(sc_tab, text="❌ Exit SPC Session", command=exit_spc_method).pack(pady=5)

    # == Scaled Weighted Variance ==
    swv_refs = {}
    swv_canvas_refs = {}

    setup_spc_tab(
        swv_tab,
        run_command=lambda: run_spc_test("swv", swv_refs, swv_canvas_refs),
        listbox_ref_dict=swv_refs,
        canvas_ref_dict=swv_canvas_refs,
        label="Run SWV SPC"
    )

    ttk.Button(swv_tab, text="❌ Exit SPC Session", command=exit_spc_method).pack(pady=5)

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()