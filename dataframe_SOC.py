import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from scipy.stats import anderson
from matplotlib.backends.backend_pdf import PdfPages
from scipy.stats import norm
from scipy.stats import skew
from scipy.stats import iqr
from scipy.interpolate import interp1d
import os
import ast
from decimal import Decimal, ROUND_HALF_UP
from tkinter import messagebox


class DataFrame_soc(pd.DataFrame):
    _metadata = ["site_of_cancer", "gamma"]  # This tells Pandas to treat it as a real attribute

    b = 6

    GPRs_n_Names = ["Name", "ID", "QA Date", "Global 3%3mm", "Global 3%2mm", "Global 3%1mm",
                    "Global 2%2mm", "Global 2%1mm", "Global 1%2mm", "Global 1%1mm",
                    "Local 3%3mm", "Local 3%2mm", "Local 3%1mm", "Local 2%2mm",
                    "Local 2%1mm", "Local 1%2mm", "Local 1%1mm"]

    mean_max = ["Global Mean Gamma Index", "Global Max Gamma Index", "Local Mean Gamma Index",
                "Local Max Gamma Index"]

    data_for_x_charts = GPRs_n_Names + ["Global Mean Gamma Index"]

    criteria = ["Global 3%3mm", "Global 3%2mm", "Global 3%1mm",
                "Global 2%2mm", "Global 2%1mm", "Global 1%2mm", "Global 1%1mm",
                "Local 3%3mm", "Local 3%2mm", "Local 3%1mm", "Local 2%2mm",
                "Local 2%1mm", "Local 1%2mm", "Local 1%1mm", "Global Mean Gamma Index"]

    z_table = {
        "90%": {"alpha": 0.10, "z_score": 1.645},
        "95%": {"alpha": 0.05, "z_score": 1.960},
        "95.45%": {"alpha": 0.0455, "z_score": 2.000},
        "99%": {"alpha": 0.01, "z_score": 2.576},
        "99.73%": {"alpha": 0.0027, "z_score": 3.000},
    }

    def __init__(self, data=None, file_path=None, *args, **kwargs):
        """
        Initialize the DataFrame_soc object. Subclass of pandas.
        """
        self.site_of_cancer = None  # Initialize attribute

        if file_path:
            ext = os.path.splitext(file_path)[1].lower()
            id_converter = {"ID": str}
            if ext == ".csv":
                data = pd.read_csv(file_path, converters=id_converter)
            elif ext in (".xls", ".xlsx"):
                try:
                    data = pd.read_excel(file_path, sheet_name="data", converters=id_converter)
                except ValueError:
                    raise ValueError("Worksheet named 'data' not found in Excel file.")
            else:
                raise ValueError("Unsupported file type. Please use .csv, .xlsx or .xls")

            super().__init__(data, *args, **kwargs)

            # Check for required columns
            required_columns = ["ID", "Site of cancer", "QA Date"]
            missing_columns = [col for col in required_columns if col not in data.columns]
            if missing_columns:
                raise ValueError(f"❌ Missing required column(s): {', '.join(missing_columns)}")

            if "ID" in self.columns:
                self["ID"] = self["ID"].astype(str)

            if "Site of cancer" in self.columns:
                self["Site of cancer"] = self["Site of cancer"].astype(str)
                self.site_of_cancer = sorted(self["Site of cancer"].dropna().unique().tolist())

            self.load_warnings = []  # Create an attribute to store any warnings

            exclude_mean_gamma = False
            if "MedianDoseDev" not in data.columns:
                self.load_warnings.append(
                    "• Column 'MedianDoseDev' not found. Global mean γ will be skipped in the SPC analysis.")
                exclude_mean_gamma = True

            self.present_criteria = [
                col for col in self.criteria
                if col in data.columns and not (exclude_mean_gamma and col == "Global Mean Gamma Index")
            ]
            self.missing_criteria = [col for col in self.criteria if col not in data.columns]

            numeric_criteria = []
            non_numeric_columns = []
            for col in self.present_criteria:
                if pd.api.types.is_numeric_dtype(self[col]):
                    numeric_criteria.append(col)
                else:
                    non_numeric_columns.append(col)

            # Add warnings for missing criteria
            if hasattr(self, "load_warnings"):
                self.load_warnings += [f"• Missing criterion: {col}" for col in self.missing_criteria]
            else:
                self.load_warnings = [f"• Missing criterion: {col}" for col in self.missing_criteria]

            if non_numeric_columns:
                if hasattr(self, "load_warnings"):
                    self.load_warnings += [f"• Non-numeric input in criterion: {col}" for col in non_numeric_columns]
                else:
                    self.load_warnings = [f"• Non-numeric input in criterion: {col}" for col in non_numeric_columns]

            self.data_for_analysis = ["ID", "QA Date"] + numeric_criteria

        self.gamma = kwargs.pop("gamma", None)

        for col in self.columns:
            if str(col).strip().lower() == "qa date":
                self.sort_by_QA_Date()

        if self.gamma is None and "MedianDoseDev" in self.columns:
            mean_dosedev = np.mean(self["MedianDoseDev"])
            self.gamma = np.sqrt((0.5 ** 2) / (2 ** 2) + (mean_dosedev ** 2) / (0.03 ** 2))  # for Global3%/2mm

    @classmethod  # This method belongs to the class (cls), not the instance (self).
    def from_file(cls, file_path, sheet_name=0):  # sheet_name=0, overwrites by sheet_name="data" above
        """Alternative constructor: Create an object by loading from an Excel file."""
        return cls(file_path=file_path)

    def round_half_up(self, value, ndigits=0):
        rounding_format = f'1.{"0" * ndigits}'
        return float(Decimal(str(value)).quantize(Decimal(rounding_format), rounding=ROUND_HALF_UP))

    def get_z_info(self, confidence_level):
        if isinstance(confidence_level, float):
            confidence_level = f"{round(confidence_level * 100, 2)}%"
        try:
            info = self.z_table[confidence_level]
            return info["alpha"], info["z_score"]
        except KeyError:
            raise ValueError(
                f"Unsupported confidence level '{confidence_level}'. "
                f"Try one of: {list(self.z_table.keys())}"
            )

    def sort_by_QA_Date(self):
        """
        Sorts the DataFrame by 'QA Date' in ascending order.
        - Converts 'QA Date' to datetime format if necessary.
        - Moves NaN values to the bottom.
        - Resets the index after sorting.
        """
        # Convert 'QA Date' column to datetime format, coercing errors to NaT
        self['QA Date'] = pd.to_datetime(self['QA Date'], errors='coerce')

        # Sort by 'QA Date' (ascending), placing NaN values at the end
        self.sort_values(by='QA Date', ascending=True, na_position='last', inplace=True)

        # Reset index
        self.reset_index(drop=True, inplace=True)

    # ===================== GUI METHODS ===================== #

    def get_summary_data(self):
        """
        Returns summary information in a structured format for GUI display.
        """
        summary = {
            "Site of Cancer": self.site_of_cancer,
            "Shape": f"{self.shape[0]} rows x {self.shape[1]} columns",
            "Columns": list(self.columns),
            "Data for analysis": list(self[self.data_for_analysis].columns)
        }

        # Clean column names for search
        clean_columns = [col.strip().lower() for col in self.columns]

        if "qa date" in clean_columns:
            col_idx = clean_columns.index("qa date")
            col_name = self.columns[col_idx]  # use original casing
            dates = pd.to_datetime(self[col_name], errors="coerce").dropna()
            if not dates.empty:
                summary["Date Range"] = f"{dates.min().date()} → {dates.max().date()}"
            else:
                summary["Date Range"] = f"'{col_name}' exists but is not parseable"

            if not self['QA Date'].dropna().is_monotonic_increasing:
                message = "⚠️ Warning: Data is not sorted by 'QA Date'. This may affect time-based control charts."
            else:
                message = "✅ Data is sorted by 'QA Date'!"
            summary["sorted data"] = f"{message}"

        return summary

    def get_statistics(self, ndigits=2):
        numeric_df = self.select_dtypes(include='number')
        stats_df = numeric_df.describe()

        # Apply custom rounding
        for col in stats_df.columns:
            stats_df[col] = stats_df[col].apply(lambda x: self.round_half_up(x, ndigits))

        return stats_df.to_string()

    def plot_histograms_gui(self, selected_columns=None, pdf=None, return_fig=False):
        numeric_data = self[self.columns.intersection(self.data_for_analysis)].select_dtypes(include=['number'])

        if selected_columns is not None:
            numeric_data = numeric_data[selected_columns]

        columns = numeric_data.columns
        sns.set_style("darkgrid")
        figs = []

        for feature in columns:
            fig, ax = plt.subplots(figsize=(8, 4))
            sns.histplot(numeric_data[feature], kde=True, color='skyblue', ax=ax)
            ax.set_title(
                f"{feature} | Skewness: {round(numeric_data[feature].skew(), 2)} | Kurtosis: {round(numeric_data[feature].kurt(), 2)}"
            )
            ax.yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
            ax.set_ylabel('Counts', fontsize=10)
            x_label = "Mean γ" if feature == "Global Mean Gamma Index" else "GPR(%)"
            ax.set_xlabel(x_label, fontsize=10)
            fig.tight_layout()

            if pdf is not None:
                pdf.savefig(fig, dpi=300, bbox_inches="tight")
                plt.close(fig)
            else:
                figs.append(fig)

        if return_fig:
            return figs if figs else None

    def run_anderson_test(self, selected_columns=None):
        """
        Perform the Anderson-Darling test for normality on numerical GPR columns.
        - If selected_columns is provided, only those are tested.
        - Uses the 5% significance level to determine normality.
        - Returns a list of dictionaries suitable for GUI Treeview display.
        """
        data_GPRs_numerical = self[self.columns.intersection(self.data_for_analysis)].select_dtypes(include=['number'])

        if selected_columns is not None:
            data_GPRs_numerical = data_GPRs_numerical[selected_columns]

        results = []
        for column in data_GPRs_numerical.columns:
            valid_data = data_GPRs_numerical[column].dropna()
            result = anderson(valid_data, dist='norm')
            is_normal = result.statistic < result.critical_values[2]  # 5% significance level
            results.append({
                'GPR': column,
                'Statistic': round(result.statistic, 3),
                'Critical Value (5%)': round(result.critical_values[2], 3),
                'Normality': 'Likely Normal' if is_normal else 'Not Normal'
            })

        df_results = pd.DataFrame(results)

        return df_results.to_dict(orient='records')

    def plot_x_chart(self, pdf=None, column=None, CL=None, UCL=None, LCL=None,
                     USL=None, LSL=None, data_to_plot=None, out_of_control=None,
                     confidence_level=None, method_name=None):
        """
        Plots X-Chart.
        """
        data_to_plot = data_to_plot.reset_index(drop=False)
        index_map = data_to_plot["index"]  # Keep original indices to recover ID later
        y_values = data_to_plot[column]

        sns.set_style("darkgrid")
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.plot(data_to_plot.index, y_values.values, marker='o', linestyle='-', color='b', label="GPR Data")
        ax.axhline(y=CL, color='green', linestyle='--', linewidth=2.5, label="CL")
        ax.axhline(y=UCL, color='red', linestyle='--', linewidth=2.3, label="UCL")
        ax.axhline(y=LCL, color='red', linestyle='--', linewidth=2.5, label="LCL")
        if USL is not None:
            ax.axhline(y=USL, color='orange', linestyle='--', linewidth=2.0, label="USL")
        if LSL is not None:
            ax.axhline(y=LSL, color='orange', linestyle='--', linewidth=2.0, label="LSL")
        # Map out-of-control positions to new 0-based index
        outlier_positions = [data_to_plot.index[data_to_plot["index"] == idx][0] for idx in out_of_control]
        ax.scatter(outlier_positions, y_values.loc[outlier_positions],
                   color='red', marker='o', s=100, edgecolors='black', zorder=3, label="Out-of-Control")

        ax.set_xlabel("Time Ordered Observations")
        ylabel = "GPR (%)" if column in self.GPRs_n_Names else "Mean γ"
        ax.set_ylabel(ylabel)
        ax.set_title(f"{method_name}: I-Chart for {column}")
        ax.legend(loc='upper left', fontsize=8)
        ax.grid(True)

        def on_click(event):
            if event.inaxes and event.xdata is not None and event.ydata is not None:
                for i, y in enumerate(y_values):
                    if abs(event.xdata - i) < 0.25 and abs(event.ydata - y) < 0.25:
                        true_index = index_map[i]
                        patient_id = self.loc[true_index, "ID"]
                        messagebox.showinfo("Point Info", f"ID: {patient_id}\nValue: {y}")
                        break

        fig.canvas.mpl_connect("button_press_event", on_click)

        fig.tight_layout()

        if pdf is not None:
            pdf.savefig(fig)
            plt.close(fig)
            return None
        else:
            return fig

    def define_outliers(self, valid_data, column, LCL, UCL, LSL, USL):
        valid_data_rounded = valid_data.apply(lambda x: self.round_half_up(x, 2))
        UCL = self.round_half_up(UCL, 2)
        LCL = self.round_half_up(LCL, 1)

        if column == "Global Mean Gamma Index":
            target = UCL + 0.01
            target = self.round_half_up(target, 2)
            UCL = target
            target2 = USL + 0.01
            target2 = self.round_half_up(target2, 2)
            USL = target2
        else:
            target = LCL - 0.1
            target = self.round_half_up(target, 1)
            LCL = target
            target2 = LSL - 0.1
            target2 = self.round_half_up(target2, 1)
            LSL = target2

        out_of_control_mask = (valid_data_rounded > UCL) | (valid_data_rounded < LCL)
        out_of_control = valid_data_rounded.index[out_of_control_mask]
        out_of_control_info = self.loc[out_of_control, ['Name', 'ID']]

        return LCL, UCL, LSL, USL, out_of_control, out_of_control_info, valid_data_rounded

    def get_shewhart_x_chart_figs(self, confidence_level="99.73%", selected_columns=None):
        """
        Return matplotlib figures for GUI display for selected columns.
        """
        alpha, Z_alpha = self.get_z_info(confidence_level)
        bita = self.b
        d2 = 1.128

        figures = []
        results_list = []

        outlier_dict = {}
        for column in selected_columns:
            valid_data = self[column].dropna()
            USL = None
            LSL = None
            MR = np.abs(np.diff(valid_data))
            mean_MR = np.mean(MR)
            CL = np.mean(valid_data)
            sigma = mean_MR / d2

            if column == "Global Mean Gamma Index":
                # TOLERANCE
                UCL = CL + Z_alpha * sigma
                LCL = 0
                # ACTIONS
                T = self.gamma
                DA = bita * np.sqrt(sigma ** 2 + (CL - T) ** 2)
                USL = CL + DA / 2
            else:
                # TOLERANCE
                UCL = 100
                LCL = CL - Z_alpha * sigma
                # ACTIONS
                T = 100
                DA = bita * np.sqrt(sigma ** 2 + (CL - T) ** 2)
                LSL = CL - DA / 2

            LCL, UCL, LSL, USL, out_of_control, out_of_control_info, valid_data_rounded = self.define_outliers(
                valid_data, column, LCL, UCL, LSL, USL)

            outlier_dict[column] = list(out_of_control_info["ID"].values)

            fig = self.plot_x_chart(
                pdf=None,
                column=column,
                CL=CL,
                UCL=UCL,
                LCL=LCL,
                USL=USL,
                LSL=LSL,
                data_to_plot=valid_data_rounded,
                out_of_control=out_of_control,
                confidence_level=confidence_level,
                method_name="Shewhart"
            )

            figures.append(fig)

            results_list.append({
                "GPR Column": column,
                "Mean (X̄)": self.round_half_up(CL, 1),
                "Counts": valid_data_rounded.count(),
                "LCL": LCL,
                "UCL": UCL,
                "LSL": LSL if LSL is not None else None,
                "USL": USL if USL is not None else None,
                "Out-of-Control IDs": list(out_of_control_info["ID"].values)
            })

        return figures, outlier_dict, results_list

    def get_swv_x_chart_figs(self, confidence_level="99.73%", selected_columns=None):
        """
        Return matplotlib figures for GUI display for selected columns.
        """
        alpha, Z_alpha = self.get_z_info(confidence_level)
        bita = self.b
        d2 = 1.128

        PX_values = np.array([0.30, 0.32, 0.34, 0.36, 0.38, 0.40, 0.42, 0.44, 0.46, 0.48,
                              0.50, 0.52, 0.54, 0.56, 0.58, 0.60, 0.62, 0.64, 0.66, 0.68, 0.70])
        WL_values = np.array([3.26, 2.98, 2.74, 2.53, 2.36, 2.26, 2.14, 2.04, 1.97, 1.93, 1.88,
                              1.86, 1.83, 1.81, 1.82, 1.84, 1.85, 1.89, 1.96, 2.04, 2.13])
        WU_values = np.flip(WL_values)

        interp_WL = interp1d(PX_values, WL_values, kind='linear',
                             fill_value="extrapolate")  # "extrapolate": If someone gives me an input outside the range of PX, just extend the line and estimate it.
        interp_WU = interp1d(PX_values, WU_values, kind='linear', fill_value="extrapolate")

        figures = []
        results_list = []

        outlier_dict = {}
        for column in selected_columns:
            valid_data = self[column].dropna()
            USL = None
            LSL = None
            MR = np.abs(np.diff(valid_data))
            mean_MR = np.mean(MR)
            CL = np.mean(valid_data)

            P_X = np.mean(valid_data <= CL)
            W_L = interp_WL(P_X)
            W_U = interp_WU(P_X)

            Z_alpha_U = norm.ppf(1 - alpha / (4 * (1 - P_X)))
            Z_alpha_L = norm.ppf(1 - alpha / (4 * P_X))

            if column == "Global Mean Gamma Index":
                # TOLERANCE
                UCL = CL + (W_U / 3) * np.sqrt(1 / (2 * (1 - P_X))) * Z_alpha_U * mean_MR
                LCL = 0
                # ACTION
                T = self.gamma
                DA_2 = np.sqrt((3 * (W_U / 3) * np.sqrt(1 / (2 * (1 - P_X))) * mean_MR) ** 2 + (3 * (CL - T)) ** 2)
                USL = CL + DA_2
            else:
                # TOLERANCE
                UCL = 100
                LCL = CL - (W_L / 3) * np.sqrt(1 / (2 * P_X)) * Z_alpha_L * mean_MR
                # ACTION
                T = 100
                DA_2 = np.sqrt((3 * (W_L / 3) * np.sqrt(1 / (2 * P_X)) * mean_MR) ** 2 + (3 * (CL - T)) ** 2)
                LSL = CL - DA_2

            LCL, UCL, LSL, USL, out_of_control, out_of_control_info, valid_data_rounded = self.define_outliers(
                valid_data, column, LCL, UCL, LSL, USL)
            outlier_dict[column] = list(out_of_control_info["ID"].values)

            fig = self.plot_x_chart(
                pdf=None,
                column=column,
                CL=CL,
                UCL=UCL,
                LCL=LCL,
                USL=USL,
                LSL=LSL,
                data_to_plot=valid_data_rounded,
                out_of_control=out_of_control,
                confidence_level=confidence_level,
                method_name="SWV"
            )

            figures.append(fig)

            results_list.append({
                "GPR Column": column,
                "Mean (X̄)": self.round_half_up(CL, 1),
                "Counts": valid_data_rounded.count(),
                "LCL": LCL,
                "UCL": UCL,
                "LSL": LSL if LSL is not None else None,
                "USL": USL if USL is not None else None,
                "Out-of-Control IDs": list(out_of_control_info["ID"].values)
            })

        return figures, outlier_dict, results_list

    def get_wsd_x_chart_figs(self, confidence_level="99.73%", selected_columns=None):
        """
        Return matplotlib figures for GUI display for selected columns.
        """
        alpha, Z_alpha = self.get_z_info(confidence_level)
        bita = self.b
        d2 = 1.128

        PX_values = np.array([0.30, 0.32, 0.34, 0.36, 0.38, 0.40, 0.42, 0.44, 0.46, 0.48,
                              0.50, 0.52, 0.54, 0.56, 0.58, 0.60, 0.62, 0.64, 0.66, 0.68, 0.70])

        d2_WSD_values = np.array([0.947, 0.982, 1.012, 1.039, 1.063, 1.083, 1.099, 1.112, 1.121,
                                  1.126, 1.128, 1.126, 1.121, 1.112, 1.099, 1.083, 1.063,
                                  1.039, 1.012, 0.982, 0.947])

        interp_d2_WSD = interp1d(PX_values, d2_WSD_values, kind='linear', fill_value="extrapolate")

        figures = []
        results_list = []

        outlier_dict = {}
        for column in selected_columns:
            valid_data = self[column].dropna()
            USL = None
            LSL = None
            MR = np.abs(np.diff(valid_data))
            mean_MR = np.mean(MR)
            CL = np.mean(valid_data)

            P_X = np.mean(valid_data <= CL)  # Probability that X ≤ X̄
            d2_WSD = interp_d2_WSD(P_X)  # Interpolate d2_WSD for given P_X

            if column == "Global Mean Gamma Index":
                # TOLERANCE
                UCL = CL + (Z_alpha * mean_MR / d2_WSD) * 2 * P_X
                LCL = 0
                # ACTION
                T = self.gamma
                DA_2 = np.sqrt(((3 * mean_MR / d2_WSD) * 2 * P_X) ** 2 + (3 * (CL - T)) ** 2)
                USL = CL + DA_2
            else:
                # TOLERANCE
                UCL = 100
                LCL = CL - (Z_alpha * mean_MR / d2_WSD) * 2 * (1 - P_X)
                # ACTION
                T = 100
                DA_2 = np.sqrt(((3 * mean_MR / d2_WSD) * 2 * (1 - P_X)) ** 2 + (3 * (CL - T)) ** 2)
                LSL = CL - DA_2

            LCL, UCL, LSL, USL, out_of_control, out_of_control_info, valid_data_rounded = self.define_outliers(
                valid_data, column, LCL, UCL, LSL, USL)

            outlier_dict[column] = list(out_of_control_info["ID"].values)

            fig = self.plot_x_chart(
                pdf=None,
                column=column,
                CL=CL,
                UCL=UCL,
                LCL=LCL,
                USL=USL,
                LSL=LSL,
                data_to_plot=valid_data_rounded,
                out_of_control=out_of_control,
                confidence_level=confidence_level,
                method_name="WSD"
            )

            figures.append(fig)

            results_list.append({
                "GPR Column": column,
                "Mean (X̄)": self.round_half_up(CL, 1),
                "Counts": valid_data_rounded.count(),
                "LCL": LCL,
                "UCL": UCL,
                "LSL": LSL if LSL is not None else None,
                "USL": USL if USL is not None else None,
                "Out-of-Control IDs": list(out_of_control_info["ID"].values)
            })

        return figures, outlier_dict, results_list

    def get_sc_x_chart_figs(self, confidence_level="99.73%", selected_columns=None):
        """
        Return matplotlib figures for GUI display for selected columns.
        """
        alpha, Z_alpha = self.get_z_info(confidence_level)
        bita = self.b
        d2 = 1.128

        k3_values = np.array([0.00, 0.40, 0.80, 1.20, 1.60, 2.00, 2.40, 2.80, 3.20, 3.60, 4.00])
        d2_sc_values = np.array([1.12, 1.12, 1.11, 1.08, 1.05, 1.02, 0.98, 0.95, 0.92, 0.90, 0.88])

        interp_d2_sc = interp1d(k3_values, d2_sc_values, kind='linear', fill_value="extrapolate")

        figures = []
        results_list = []

        outlier_dict = {}
        for column in selected_columns:
            valid_data = self[column].dropna()
            USL = None
            LSL = None
            MR = np.abs(np.diff(valid_data))
            mean_MR = np.mean(MR)
            CL = np.mean(valid_data)

            k3 = skew(valid_data)  # Using built-in skewness function
            d2_sc = interp_d2_sc(abs(k3))
            skew_factor = (1 / 6) * (Z_alpha ** 2 - 1) * k3 / (1 + 0.2 * k3 ** 2)

            if column == "Global Mean Gamma Index":
                # TOLERANCE
                UCL = CL + (Z_alpha + skew_factor) * (mean_MR / d2_sc)
                LCL = 0
                # ACTION
                skew_factor_action = (4 / 3) * k3 / (1 + 0.2 * k3 ** 2)
                T = self.gamma
                DA_2 = np.sqrt(((3 + skew_factor_action) * (mean_MR / d2_sc)) ** 2 + (3 * (CL - T)) ** 2)
                USL = CL + DA_2
            else:
                # TOLERANCE
                UCL = 100
                LCL = CL + (-Z_alpha + skew_factor) * (mean_MR / d2_sc)
                # ACTION
                skew_factor_action = (4 / 3) * k3 / (1 + 0.2 * k3 ** 2)
                T = 100
                DA_2 = np.sqrt(((-3 + skew_factor_action) * (mean_MR / d2_sc)) ** 2 + (3 * (CL - T)) ** 2)
                LSL = CL - DA_2

            LCL, UCL, LSL, USL, out_of_control, out_of_control_info, valid_data_rounded = self.define_outliers(
                valid_data, column, LCL, UCL, LSL, USL)

            outlier_dict[column] = list(out_of_control_info["ID"].values)

            fig = self.plot_x_chart(
                pdf=None,
                column=column,
                CL=CL,
                UCL=UCL,
                LCL=LCL,
                USL=USL,
                LSL=LSL,
                data_to_plot=valid_data_rounded,
                out_of_control=out_of_control,
                confidence_level=confidence_level,
                method_name="SC"
            )

            figures.append(fig)

            results_list.append({
                "GPR Column": column,
                "Mean (X̄)": self.round_half_up(CL, 1),
                "Counts": valid_data_rounded.count(),
                "LCL": LCL,
                "UCL": UCL,
                "LSL": LSL if LSL is not None else None,
                "USL": USL if USL is not None else None,
                "Out-of-Control IDs": list(out_of_control_info["ID"].values)
            })

        return figures, outlier_dict, results_list

    def elimination_recalculate_gui(self, method="shewhart", confidence_level="99.73%",
                                    selected_criterion=None, selected_ids=None, round_num=1):

        eliminated_log = []  # To store tuples like (criterion, ID, value)

        criterion = selected_criterion
        IDs_input = selected_ids

        valid_methods = {
            "shewhart": self.get_shewhart_x_chart_figs,
            "swv": self.get_swv_x_chart_figs,
            "wsd": self.get_wsd_x_chart_figs,
            "sc": self.get_sc_x_chart_figs
        }

        if criterion == "Global 3%2mm":
            for ID in IDs_input:
                row = self[self["ID"] == ID]
                if not row.empty:
                    row_index = row.index[0]  # 🔹 Only one row should match
                    for crit in self.criteria:
                        if crit in self.columns:
                            value = self.at[row_index, crit]
                            self.at[row_index, crit] = np.nan
                            eliminated_log.append((round_num, crit, f"'{ID}", value))
        else:
            for ID in IDs_input:
                row = self[self["ID"] == ID]
                if not row.empty:
                    row_index = row.index[0]
                    value = self.at[row_index, criterion]
                    self.at[row_index, criterion] = np.nan
                    eliminated_log.append((round_num, criterion, f"'{ID}", value))

        return eliminated_log