import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import pyautogui
import time
import sys
import shutil
from datetime import datetime

selected_file = None
processed_file = None

REQUIRED_COLUMNS = [
    "Campaign", "Platform", "Content Type", "Impressions", "Engagements",
    "Clicks", "Conversions", "Cost", "Revenue", "Status"
]


def get_base_path():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def classify_performance(roi):
    if pd.isna(roi):
        return "Needs Review"
    if roi > 1:
        return "High Performing"
    elif roi >= 0.3:
        return "Average Performing"
    else:
        return "Needs Review"


def recommend_action(row):
    roi = row["ROI"]
    ctr = row["CTR"]
    conversion_rate = row["Conversions Rate%"]
    cost_per_conversion = row["Cost Per Conversion"]

    if pd.isna(roi) or pd.isna(ctr) or pd.isna(conversion_rate) or pd.isna(cost_per_conversion):
        return "Review Data"

    if roi > 1 and ctr > 0.02:
        return "Increase Budget"
    elif roi < 0.3 and cost_per_conversion > 20:
        return "Reduce Budget"
    elif ctr < 0.01 and conversion_rate < 0.10:
        return "Review Campaign Content"
    elif ctr >= 0.02 and conversion_rate < 0.10:
        return "Improve Call-to-Action"
    else:
        return "Maintain and Monitor"


def priority_score(row):
    score = 0

    if pd.isna(row["ROI"]) or pd.isna(row["CTR"]) or pd.isna(row["Conversions Rate%"]) or pd.isna(row["Cost Per Conversion"]):
        return 6

    if row["ROI"] < 0.3:
        score += 3
    if row["CTR"] < 0.01:
        score += 2
    if row["Conversions Rate%"] < 0.10:
        score += 2
    if row["Cost Per Conversion"] > 20:
        score += 3

    return score


def priority_label(score):
    if score >= 6:
        return "High Priority"
    elif score >= 3:
        return "Medium Priority"
    else:
        return "Low Priority"


def alert_flag(row):
    roi = row["ROI"]
    ctr = row["CTR"]
    conversion_rate = row["Conversions Rate%"]

    if pd.isna(roi) or pd.isna(ctr) or pd.isna(conversion_rate):
        return "Data Issue"
    if roi < 0:
        return "Loss Making"
    elif ctr < 0.01:
        return "Low Engagement"
    elif conversion_rate < 0.10:
        return "Low Conversion"
    else:
        return "Normal"


def set_status(text, color="#1F2937"):
    status_label.config(text=f"Status: {text}", fg=color)
    root.update_idletasks()


def add_hover_effect(button, normal_bg, hover_bg):
    def on_enter(_):
        button["bg"] = hover_bg

    def on_leave(_):
        button["bg"] = normal_bg

    button.bind("<Enter>", on_enter)
    button.bind("<Leave>", on_leave)


def apply_icon(window):
    icon_path = os.path.join(get_base_path(), "app_icon.ico")
    if os.path.exists(icon_path):
        try:
            window.iconbitmap(icon_path)
        except Exception:
            pass


def show_splash():
    splash = tk.Tk()
    splash.overrideredirect(True)
    splash.geometry("520x240+420+220")
    splash.configure(bg="#1E3A8A")

    apply_icon(splash)

    splash_title = tk.Label(
        splash,
        text="SME ROI Decision Support System",
        font=("Segoe UI", 22, "bold"),
        fg="white",
        bg="#1E3A8A"
    )
    splash_title.pack(pady=(55, 10))

    splash_subtitle = tk.Label(
        splash,
        text="Initializing integrated campaign analytics environment...",
        font=("Segoe UI", 11),
        fg="#DBEAFE",
        bg="#1E3A8A"
    )
    splash_subtitle.pack()

    splash.update()
    time.sleep(2)
    splash.destroy()


def create_template_if_missing():
    template_path = os.path.join(get_base_path(), "Template.xlsx")
    if not os.path.exists(template_path):
        template_df = pd.DataFrame(columns=REQUIRED_COLUMNS)
        template_df.to_excel(template_path, index=False, engine="openpyxl")


def download_template():
    try:
        create_template_if_missing()
        src = os.path.join(get_base_path(), "Template.xlsx")

        save_path = filedialog.asksaveasfilename(
            title="Save Template As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="SME_Template.xlsx"
        )

        if save_path:
            shutil.copy(src, save_path)
            messagebox.showinfo("Template Downloaded", f"Template saved successfully:\n{save_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


def select_file():
    global selected_file

    file_path = filedialog.askopenfilename(
        title="Select Excel Dataset",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if file_path:
        selected_file = file_path
        file_label.config(
            text=f"Selected File: {os.path.basename(file_path)}",
            fg="#111827"
        )
        set_status("File selected. Ready for validation and processing.", "#1D4ED8")
        summary_box.delete("1.0", tk.END)


def process_file():
    global selected_file, processed_file

    if not selected_file:
        messagebox.showwarning("No File Selected", "Please select a dataset first.")
        return

    try:
        set_status("Reading dataset...", "#1D4ED8")

        df = pd.read_excel(selected_file)
        original_rows = len(df)

        df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
        df.columns = df.columns.astype(str).str.strip()

        found_columns = list(df.columns)
        missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]

        summary_box.delete("1.0", tk.END)
        summary_box.insert(tk.END, "VALIDATION REPORT\n")
        summary_box.insert(tk.END, "------------------------------\n")
        summary_box.insert(tk.END, f"Selected File: {os.path.basename(selected_file)}\n")
        summary_box.insert(tk.END, f"Rows Before Processing: {original_rows}\n")
        summary_box.insert(tk.END, f"Columns Found: {len(found_columns)}\n\n")

        if missing_columns:
            summary_box.insert(tk.END, "Missing Required Columns:\n")
            for col in missing_columns:
                summary_box.insert(tk.END, f"- {col}\n")

            summary_box.insert(tk.END, "\nPlease use the system template to prepare the dataset correctly.")
            set_status("Validation failed. Missing required columns.", "#DC2626")
            messagebox.showerror(
                "Missing Columns",
                f"The following required columns are missing:\n{missing_columns}"
            )
            return

        summary_box.insert(tk.END, "Validation Passed: All required columns found.\n\n")
        set_status("Validation passed. Cleaning and processing data...", "#1D4ED8")

        numeric_columns = [
            "Impressions", "Engagements", "Clicks",
            "Conversions", "Cost", "Revenue"
        ]

        for col in numeric_columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

        before_cleaning_rows = len(df)

        df = df.dropna(subset=REQUIRED_COLUMNS)
        df = df.dropna(subset=numeric_columns)
        df = df[df["Impressions"] > 0]
        df = df[df["Clicks"] > 0]
        df = df[df["Conversions"] > 0]
        df = df[df["Cost"] > 0]

        after_cleaning_rows = len(df)
        removed_rows = before_cleaning_rows - after_cleaning_rows

        df["Engagement Rate%"] = df["Engagements"] / df["Impressions"].replace(0, pd.NA)
        df["Conversions Rate%"] = df["Conversions"] / df["Clicks"].replace(0, pd.NA)
        df["ROI"] = (df["Revenue"] - df["Cost"]) / df["Cost"].replace(0, pd.NA)
        df["CTR"] = df["Clicks"] / df["Impressions"].replace(0, pd.NA)
        df["Cost Per Conversion"] = df["Cost"] / df["Conversions"].replace(0, pd.NA)

        df["Performance Category"] = df["ROI"].apply(classify_performance)
        df["Recommended Action"] = df.apply(recommend_action, axis=1)
        df["Priority Score"] = df.apply(priority_score, axis=1)
        df["Priority Level"] = df["Priority Score"].apply(priority_label)
        df["Alert"] = df.apply(alert_flag, axis=1)

        df["Processed Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df["Source File"] = os.path.basename(selected_file)

        processed_file = os.path.join(get_base_path(), "Processed_Output.xlsx")
        df.to_excel(processed_file, index=False, engine="openpyxl")

        total_campaigns = len(df)
        avg_roi = df["ROI"].mean()
        high_priority = (df["Priority Level"] == "High Priority").sum()
        alerts_count = (df["Alert"] != "Normal").sum()
        high_performing = (df["Performance Category"] == "High Performing").sum()

        summary_box.insert(tk.END, "PROCESSING SUMMARY\n")
        summary_box.insert(tk.END, "------------------------------\n")
        summary_box.insert(tk.END, f"Rows After Processing: {after_cleaning_rows}\n")
        summary_box.insert(tk.END, f"Rows Removed: {removed_rows}\n")
        summary_box.insert(tk.END, f"Total Campaigns: {total_campaigns}\n")
        summary_box.insert(tk.END, f"Average ROI: {avg_roi:.2%}\n")
        summary_box.insert(tk.END, f"High Performing Campaigns: {high_performing}\n")
        summary_box.insert(tk.END, f"High Priority Campaigns: {high_priority}\n")
        summary_box.insert(tk.END, f"Campaigns with Alerts: {alerts_count}\n\n")
        summary_box.insert(tk.END, f"Processed file saved as:\n{processed_file}\n\n")
        summary_box.insert(tk.END, "Next Step:\nClick 'Open Dashboard' to view updated insights.")

        set_status("Processing complete. System ready.", "#059669")
        messagebox.showinfo("Success", "Dataset processed successfully.")

    except Exception as e:
        set_status("An error occurred during processing.", "#DC2626")
        messagebox.showerror("Error", str(e))


def open_dashboard():
    base_path = get_base_path()
    pbix_path = os.path.join(base_path, "SME_Dashboard.pbix")

    if not os.path.exists(pbix_path):
        messagebox.showerror(
            "Error",
            f"Dashboard not found.\n\nExpected location:\n{pbix_path}"
        )
        return

    try:
        set_status("Opening Power BI dashboard...", "#1D4ED8")
        os.startfile(pbix_path)

        time.sleep(20)

        try:
            pyautogui.hotkey("ctrl", "r")
        except Exception:
            pass

        set_status("Dashboard opened and refresh command sent.", "#059669")
        messagebox.showinfo(
            "Dashboard Opened",
            "Power BI dashboard opened.\nA refresh command was sent automatically."
        )

    except Exception as e:
        set_status("Dashboard opened, but refresh may have failed.", "#D97706")
        messagebox.showwarning(
            "Warning",
            f"Dashboard opened, but automatic refresh may not have worked.\n\nDetails:\n{e}"
        )


def exit_app():
    root.destroy()


show_splash()
create_template_if_missing()

root = tk.Tk()
root.title("SME ROI Decision Support System")

# Opens maximized on Windows
root.state("zoomed")

# Backup size if zoomed does not work
root.geometry("1200x800")

root.configure(bg="#EEF2F7")
root.resizable(True, True)

apply_icon(root)

header_frame = tk.Frame(root, bg="#173A8C", height=110)
header_frame.pack(fill="x")

title_label = tk.Label(
    header_frame,
    text="SME ROI Decision Support System",
    font=("Segoe UI", 28, "bold"),
    fg="white",
    bg="#173A8C"
)
title_label.pack(pady=(22, 6))

subtitle_label = tk.Label(
    header_frame,
    text="Upload, validate, process, and analyze campaign data through one integrated system.",
    font=("Segoe UI", 13),
    fg="#DCEAFE",
    bg="#173A8C"
)
subtitle_label.pack()

content_frame = tk.Frame(root, bg="#EEF2F7")
content_frame.pack(fill="both", expand=True, padx=35, pady=25)

controls_frame = tk.Frame(content_frame, bg="white", bd=1, relief="solid")
controls_frame.pack(fill="x", pady=(0, 18))

controls_title = tk.Label(
    controls_frame,
    text="System Controls",
    font=("Segoe UI", 17, "bold"),
    bg="white",
    fg="#111827"
)
controls_title.grid(row=0, column=0, columnspan=2, pady=(18, 14))

button_style = {
    "font": ("Segoe UI", 13, "bold"),
    "width": 22,
    "height": 2,
    "bd": 0,
    "cursor": "hand2"
}

select_button = tk.Button(
    controls_frame,
    text="Select Dataset",
    command=select_file,
    bg="#2563EB",
    fg="white",
    activebackground="#1D4ED8",
    activeforeground="white",
    **button_style
)
select_button.grid(row=1, column=0, padx=18, pady=12)

process_button = tk.Button(
    controls_frame,
    text="Process Dataset",
    command=process_file,
    bg="#059669",
    fg="white",
    activebackground="#047857",
    activeforeground="white",
    **button_style
)
process_button.grid(row=1, column=1, padx=18, pady=12)

template_button = tk.Button(
    controls_frame,
    text="Download Template",
    command=download_template,
    bg="#F59E0B",
    fg="white",
    activebackground="#D97706",
    activeforeground="white",
    **button_style
)
template_button.grid(row=2, column=0, padx=18, pady=12)

open_dashboard_button = tk.Button(
    controls_frame,
    text="Open Dashboard",
    command=open_dashboard,
    bg="#7C3AED",
    fg="white",
    activebackground="#6D28D9",
    activeforeground="white",
    **button_style
)
open_dashboard_button.grid(row=2, column=1, padx=18, pady=12)

controls_frame.grid_columnconfigure(0, weight=1)
controls_frame.grid_columnconfigure(1, weight=1)

add_hover_effect(select_button, "#2563EB", "#1D4ED8")
add_hover_effect(process_button, "#059669", "#047857")
add_hover_effect(template_button, "#F59E0B", "#D97706")
add_hover_effect(open_dashboard_button, "#7C3AED", "#6D28D9")

file_label = tk.Label(
    controls_frame,
    text="Selected File: None",
    font=("Segoe UI", 12),
    bg="white",
    fg="#374151"
)
file_label.grid(row=3, column=0, columnspan=2, pady=(10, 8))

status_label = tk.Label(
    controls_frame,
    text="Status: Waiting for user input.",
    font=("Segoe UI", 12, "italic"),
    bg="white",
    fg="#374151"
)
status_label.grid(row=4, column=0, columnspan=2, pady=(0, 18))

summary_frame = tk.Frame(content_frame, bg="white", bd=1, relief="solid")
summary_frame.pack(fill="both", expand=True)

summary_title = tk.Label(
    summary_frame,
    text="Validation and Processing Summary",
    font=("Segoe UI", 17, "bold"),
    bg="white",
    fg="#111827"
)
summary_title.pack(pady=(16, 12))

summary_box = tk.Text(
    summary_frame,
    height=18,
    width=120,
    font=("Consolas", 11),
    bg="#F8FAFC",
    fg="#111827",
    bd=0,
    wrap="word",
    insertbackground="#111827"
)
summary_box.pack(fill="both", expand=True, padx=18, pady=(0, 18))

exit_button = tk.Button(
    root,
    text="Exit System",
    command=exit_app,
    font=("Segoe UI", 12, "bold"),
    width=22,
    height=2,
    bg="#DC2626",
    fg="white",
    activebackground="#B91C1C",
    activeforeground="white",
    bd=0,
    cursor="hand2"
)
exit_button.pack(pady=(0, 18))

add_hover_effect(exit_button, "#DC2626", "#B91C1C")

root.mainloop()