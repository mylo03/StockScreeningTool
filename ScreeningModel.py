import pandas as pd
import tkinter as tk
from tkinter import ttk
import openpyxl

def fetch_data():
    data = pd.read_excel('/Users/mylofarman/Documents/2R_Capital/Tickers.xlsx')
    return data

def apply_filters(data, filters):
    # Apply filters, however do nothing if empty
    for metric, condition in filters.items():
        min_val, max_val = condition
        if min_val != '':
            data = data[data[metric] >= float(min_val)]
        if max_val != '':
            data = data[data[metric] <= float(max_val)]
    return data


def rank_companies(data, weights):
    # Create a composite score, by calculating a score based on the index for that metric
    metrics = ['ROIC', 'EBIT Margin', '52-Week Price Low Relative',
               'Revenue Growth Rate', 'EBIT Margin Improvement']
    for metric in metrics:
        data.loc[:, f'{metric}_rank'] = data[metric].rank(ascending=True, method='dense')

    data.loc[:, 'composite_score'] = (
            data['ROIC_rank'] * weights['ROIC'] +
            data['EBIT Margin_rank'] * weights['EBIT Margin'] +
            data['52-Week Price Low Relative_rank'] * weights['52-Week Price Low Relative'] +
            data['Revenue Growth Rate_rank'] * weights['Revenue Growth Rate'] +
            data['EBIT Margin Improvement_rank'] * weights['EBIT Margin Improvement']
    )

    ranked_data = data.sort_values(by='composite_score', ascending=False)
    return ranked_data


def run_screening():
    # Fetch filters from the UI
    filters = {
        'ROIC': (roic_min_entry.get(), roic_max_entry.get()),
        'EBIT Margin': (ebit_margin_min_entry.get(), ebit_margin_max_entry.get()),
        '52-Week Price Low Relative': (week_52_low_min_entry.get(), week_52_low_max_entry.get()),
        'Revenue Growth Rate': (revenue_growth_min_entry.get(), revenue_growth_max_entry.get()),
        'EBIT Margin Improvement': (ebit_margin_improvement_min_entry.get(), ebit_margin_improvement_max_entry.get())
    }

    # Collect the data
    data = fetch_data()
    filtered_data = apply_filters(data, filters)

    # Fetch weights from the UI
    weights = {
        'ROIC': float(roic_weight_entry.get()),
        'EBIT Margin': float(ebit_margin_weight_entry.get()),
        '52-Week Price Low Relative': float(week_52_low_weight_entry.get()),
        'Revenue Growth Rate': float(revenue_growth_weight_entry.get()),
        'EBIT Margin Improvement': float(ebit_margin_improvement_weight_entry.get())
    }

    ranked_data = rank_companies(filtered_data, weights)
    ranked_data = ranked_data[:10]

    # Empty the results section
    for row in tree.get_children():
        tree.delete(row)

    # Add the new results
    for index, row in ranked_data.iterrows():
        tree.insert("", "end", values=(row['Company Name'], f"{row['composite_score']:.2f}"))

    # Highlighting the top company, with the option to highlight additional such as 2-5 etc.
    for i in range(min(10, len(ranked_data))):
        if i == 0:
            tree.item(tree.get_children()[i], tags=("highlight1",))

    tree.tag_configure("highlight1", background="gold")

# Create the main window
root = tk.Tk()
root.title("Stock Screening Tool")

# For filters and weights frame
left_frame = ttk.Frame(root)
left_frame.grid(row=0, column=0, padx=20, pady=20, sticky='nsew')
filter_frame = ttk.Frame(left_frame)
filter_frame.grid(row=0, column=0, sticky='w')

# Weights  label
weights_label = ttk.Label(filter_frame, text="Weights", font=("Arial", 10, "bold"))
weights_label.grid(row=0, column=3, padx=5, pady=5)

# Maximum filter label
filter_label = ttk.Label(filter_frame, text="Maximum Filter", font=("Arial", 10, "bold"))
filter_label.grid(row=0, column=2, padx=5, pady=5)

# Minimum filter label
filter_label = ttk.Label(filter_frame, text="Minimum Filter", font=("Arial", 10, "bold"))
filter_label.grid(row=0, column=1, padx=5, pady=5)

for i, (label_text, weight_default) in enumerate([
    ("ROIC (%)", 0.25),
    ("EBIT Margin (%)", 0.25),
    ("52-Week Price Low Relative (%)", 0.20),
    ("Revenue Growth (5Y Avg) (%)", 0.15),
    ("EBIT Margin Improvement (5Y Avg) (%)", 0.15)]):
    label = ttk.Label(filter_frame, text=label_text)
    label.grid(row=i+1, column=0, sticky='w', padx=5, pady=20)

    # Setting a minimum filter, if empty will not be applied
    min_entry = ttk.Entry(filter_frame, width=6)
    min_entry.grid(row=i+1, column=1, padx=5, pady=(2, 2))
    if label_text == "ROIC (%)":
        roic_min_entry = min_entry
    elif label_text == "EBIT Margin (%)":
        ebit_margin_min_entry = min_entry
    elif label_text == "52-Week Price Low Relative (%)":
        week_52_low_min_entry = min_entry
    elif label_text == "Revenue Growth (5Y Avg) (%)":
        revenue_growth_min_entry = min_entry
    elif label_text == "EBIT Margin Improvement (5Y Avg) (%)":
        ebit_margin_improvement_min_entry = min_entry

    # Setting a maximum filter, if empty will not be applied
    max_entry = ttk.Entry(filter_frame, width=6)
    max_entry.grid(row=i+1, column=2, padx=5, pady=(2, 2))
    if label_text == "ROIC (%)":
        roic_max_entry = max_entry
    elif label_text == "EBIT Margin (%)":
        ebit_margin_max_entry = max_entry
    elif label_text == "52-Week Price Low Relative (%)":
        week_52_low_max_entry = max_entry
    elif label_text == "Revenue Growth (5Y Avg) (%)":
        revenue_growth_max_entry = max_entry
    elif label_text == "EBIT Margin Improvement (5Y Avg) (%)":
        ebit_margin_improvement_max_entry = max_entry

    # Weight entry next to entering the filters minimum and maximum
    weight_entry = ttk.Entry(filter_frame, width=6)
    weight_entry.insert(0, str(weight_default))
    weight_entry.grid(row=i+1, column=3, padx=5, pady=5)

    # Assign weight entries to the different metrics
    if label_text == "ROIC (%)":
        roic_weight_entry = weight_entry
    elif label_text == "EBIT Margin (%)":
        ebit_margin_weight_entry = weight_entry
    elif label_text == "52-Week Price Low Relative (%)":
        week_52_low_weight_entry = weight_entry
    elif label_text == "Revenue Growth (5Y Avg) (%)":
        revenue_growth_weight_entry = weight_entry
    elif label_text == "EBIT Margin Improvement (5Y Avg) (%)":
        ebit_margin_improvement_weight_entry = weight_entry


# Button to run screening
button_frame = ttk.Frame(left_frame)
button_frame.grid(row=len(filter_frame.grid_info()), column=0, columnspan=3, pady=20)

run_button = ttk.Button(button_frame, text="Run Screening", command=run_screening)
run_button.grid(row=0, column=0, padx=5, pady=5)

# For Displaying the top securities after filtering and weighting
result_frame = ttk.Frame(root, width=600, height=400)
result_frame.grid(row=0, column=1, padx=20, pady=20, sticky='nsew')
tree = ttk.Treeview(result_frame, columns=("Company", "Score"), show="headings")
tree.heading("Company", text="Company")
tree.heading("Score", text="Score")
tree.column("Company", anchor='w', width=250)
tree.column("Score", anchor='center', width=100)
tree.pack(expand=True, fill='both')

# Configure the grid weights for responsive resizing
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=2)

# Start the main loop
root.mainloop()