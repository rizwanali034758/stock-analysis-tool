import tkinter as tk
from tkinter import Canvas, messagebox
import logging

# Configure logging
logging.basicConfig(filename='stock_analysis.log', level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

def create_gui():
    root = tk.Tk()
    root.title("Comprehensive Stock Valuation Tool")
    root.geometry("900x600")

    # Create a main frame with a grid layout
    main_frame = tk.Frame(root)
    main_frame.pack(fill=tk.BOTH, expand=True)

    # Create a canvas to allow for scrolling
    canvas = tk.Canvas(main_frame)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Add a vertical scrollbar
    scrollbar = tk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # Create a frame to contain all the fields
    scrollable_frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')
    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    # Define field sections
    fields = {
        "Profitability": [
            ("Revenue:", "revenue"),
            ("Net Income:", "net_income"),
            ("Operating Profit:", "operating_profit"),
            ("Gross Profit:", "gross_profit"),
            ("Stock Price:", "stock_price"),
            ("Book Value:", "book_value"),
            ("Free Cash Flow:", "free_cash_flow"),
            ("Total Assets:", "total_assets"),
            ("Total Liabilities:", "total_liabilities"),
            ("Interest Expense:", "interest_expense")
        ],
        
        "Growth and Dividends": [
            ("Estimated Growth Rate (%):", "growth_rate"),
            ("Annual Dividends Per Share:", "dividends")
        ]
    }

    # Layout fields in grid
    row = 0
    col = 0
    for section, items in fields.items():
        tk.Label(scrollable_frame, text=section, font=("Arial", 14, "bold")).grid(row=row, column=col, padx=20, pady=10, sticky='w')
        row += 1
        for label_text, entry_name in items:
            tk.Label(scrollable_frame, text=label_text, font=("Arial", 12)).grid(row=row, column=col, padx=20, pady=5, sticky='w')
            entry = tk.Entry(scrollable_frame, width=30)
            entry.grid(row=row, column=col+1, padx=20, pady=5)
            globals()[f"{entry_name}_entry"] = entry
            row += 1
        row = 0
        col += 2

    # Add a large submit button at the bottom
    submit_button = tk.Button(scrollable_frame, text="Submit", command=submit_data, font=("Arial", 16), bg="blue", fg="white", width=20, height=2, relief=tk.RAISED)
    submit_button.grid(row=row, column=0, columnspan=4, pady=20)

    root.mainloop()

def submit_data():
    try:
        # Get user input and validate
        data = {}
        for entry_name in ["revenue", "net_income", "operating_profit", "gross_profit", "stock_price", "book_value", "free_cash_flow", "total_assets", "total_liabilities", "interest_expense", "growth_rate", "dividends"]:
            entry = globals()[f"{entry_name}_entry"]
            data[entry_name] = float(entry.get())

        revenue = data["revenue"]
        net_income = data["net_income"]
        operating_profit = data["operating_profit"]
        gross_profit = data["gross_profit"]
        stock_price = data["stock_price"]
        book_value = data["book_value"]
        free_cash_flow = data["free_cash_flow"]
        total_assets = data["total_assets"]
        total_liabilities = data["total_liabilities"]
        interest_expense = data["interest_expense"]
        growth_rate = data["growth_rate"]
        dividends = data["dividends"]

        # Check for zero values to prevent division errors
        if total_assets - total_liabilities == 0:
            raise ValueError("The difference between Total Assets and Total Liabilities cannot be zero.")

        # ----- Financial Ratios -----
        gross_profit_margin = (gross_profit / revenue) * 100
        operating_profit_margin = (operating_profit / revenue) * 100
        net_profit_margin = (net_income / revenue) * 100

        # Handle division by zero for ROA and ROE
        roa = (net_income / total_assets) * 100 if total_assets != 0 else float('inf')
        roe = (net_income / (total_assets - total_liabilities)) * 100 if (total_assets - total_liabilities) != 0 else float('inf')

        # ----- Valuation Ratios -----
        pe_ratio = stock_price / (net_income / (total_assets - total_liabilities)) if (total_assets - total_liabilities) != 0 else float('inf')
        pb_ratio = stock_price / (book_value / (total_assets - total_liabilities)) if (total_assets - total_liabilities) != 0 else float('inf')
        ps_ratio = stock_price / (revenue / (total_assets - total_liabilities)) if (total_assets - total_liabilities) != 0 else float('inf')

        # ----- Solvency Ratios -----
        debt_to_equity_ratio = total_liabilities / (total_assets - total_liabilities) if (total_assets - total_liabilities) != 0 else float('inf')
        interest_coverage_ratio = operating_profit / interest_expense if interest_expense != 0 else float('inf')

        # ----- Growth Ratios -----
        eps = net_income / (total_assets - total_liabilities) if (total_assets - total_liabilities) != 0 else float('inf')
        free_cash_flow_yield = free_cash_flow / (stock_price * (total_assets - total_liabilities)) if (stock_price * (total_assets - total_liabilities)) != 0 else float('inf')
        dividend_yield = (dividends / stock_price) * 100 if stock_price != 0 else float('inf')

        # ----- Final Investment Suggestions -----
        suggestions = []

        # Analyze profitability
        if gross_profit_margin > 40:
            suggestions.append("High gross profit margin indicates strong efficiency.")
        if operating_profit_margin > 20:
            suggestions.append("Operating profit margin is excellent for long-term investments.")
        if net_profit_margin > 10:
            suggestions.append("Net profit margin is healthy.")

        # Analyze valuation
        if pe_ratio < 15:
            suggestions.append("P/E ratio indicates the stock is undervalued.")
        elif pe_ratio > 25:
            suggestions.append("P/E ratio suggests overvaluation.")

        if pb_ratio < 1:
            suggestions.append("P/B ratio indicates the stock is undervalued.")
        elif pb_ratio > 3:
            suggestions.append("P/B ratio suggests the stock might be overvalued.")

        if ps_ratio < 1:
            suggestions.append("P/S ratio indicates good valuation for sales.")

        # Analyze solvency
        if debt_to_equity_ratio < 0.5:
            suggestions.append("Low debt-to-equity ratio indicates low financial risk.")
        elif debt_to_equity_ratio > 1:
            suggestions.append("High debt-to-equity ratio may indicate higher financial risk.")

        if interest_coverage_ratio > 5:
            suggestions.append("High interest coverage ratio indicates the company can easily cover its debt.")
        elif interest_coverage_ratio < 1.5:
            suggestions.append("Low interest coverage ratio indicates the company may struggle with its debt.")

        # Analyze growth and dividend
        if dividend_yield > 3:
            suggestions.append("High dividend yield indicates a good source of income for long-term investors.")
        if free_cash_flow_yield > 5:
            suggestions.append("Free cash flow yield indicates good potential for growth.")

        # Final decision based on overall analysis
        if len(suggestions) > 5:
            final_decision = "The stock appears to be a strong buy for long-term investment."
        elif 3 <= len(suggestions) <= 5:
            final_decision = "The stock is fairly valued. It may be a hold or cautious buy."
        else:
            final_decision = "The stock may not be a good long-term investment based on the data."

        # Display the results in a new window or popup with enhanced styling
        result_window = tk.Toplevel(root)
        result_window.title("Investment Analysis Result")
        result_window.geometry("600x800")
        result_window.configure(bg="#f0f0f0")

        # Create a styled result label
        result_text = (
            "----- Financial Ratios Analysis -----\n\n"
            f"Gross Profit Margin: {gross_profit_margin:.2f}%\n"
            f"Operating Profit Margin: {operating_profit_margin:.2f}%\n"
            f"Net Profit Margin: {net_profit_margin:.2f}%\n"
            f"Return on Assets (ROA): {roa:.2f}%\n"
            f"Return on Equity (ROE): {roe:.2f}%\n\n"
            f"Price-to-Earnings Ratio (P/E): {pe_ratio:.2f}\n"
            f"Price-to-Book Ratio (P/B): {pb_ratio:.2f}\n"
            f"Price-to-Sales Ratio (P/S): {ps_ratio:.2f}\n\n"
            f"Debt-to-Equity Ratio: {debt_to_equity_ratio:.2f}\n"
            f"Interest Coverage Ratio: {interest_coverage_ratio:.2f}\n\n"
            f"Earnings Per Share (EPS): {eps:.2f}\n"
            f"Free Cash Flow Yield: {free_cash_flow_yield:.2f}%\n"
            f"Dividend Yield: {dividend_yield:.2f}%\n\n"
            "----- Investment Suggestions -----\n"
            f"{' '.join(suggestions)}\n\n"
            "----- Final Decision -----\n"
            f"{final_decision}"
        )

        color_result = "green" if len(suggestions) > 5 else "red" if len(suggestions) <= 3 else "orange"

        result_label = tk.Label(result_window, text=result_text, bg="#f0f0f0", font=("Arial", 12), justify=tk.LEFT)
        result_label.pack(pady=20, padx=20)

        result_label.config(fg=color_result)

        # Adjust the position of the scrollbar
        logging.root.update_idletasks()
        Canvas.yview_moveto(1.0)  # Scroll to bottom

    except ValueError as e:
        logging.error("ValueError: %s", e)
        messagebox.showerror("Input Error", "Please enter valid numeric values!")
    except ZeroDivisionError as e:
        logging.error("ZeroDivisionError: %s", e)
        messagebox.showerror("Calculation Error", "A division by zero error occurred. Please check your input values.")
    except Exception as e:
        logging.error("Unexpected error: %s", e)
        messagebox.showerror("Unexpected Error", "An unexpected error occurred. Please try again.")

create_gui()
