import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from tkinter import ttk
import pandas as pd

class FinancialAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Financial Data Entry")
        self.create_widgets()

    def create_widgets(self):
        instructions = tk.Label(self.root, text="Enter financial data (without commas)", font=('Arial', 12), pady=10)
        instructions.grid(row=0, column=0, columnspan=4)

        self.entries = []
        field_labels = [
            ("Revenue", "Enter total revenue from income statement"),
            ("Net Profit", "Enter net profit after tax from income statement"),
            ("Total Assets", "Enter total assets from balance sheet"),
            ("Equity", "Enter shareholders' equity from balance sheet"),
            ("Current Assets", "Enter current assets from balance sheet"),
            ("Current Liabilities", "Enter current liabilities from balance sheet"),
            ("Cash Flow", "Enter cash flow from cash flow statement"),
            ("Total Liabilities", "Enter total liabilities from balance sheet"),
            ("Capital Expenditure", "Enter capital expenditure from cash flow statement"),
            ("EBITDA", "Enter EBITDA from financial statement"),
            ("Market Cap", "Enter market capitalization of the company"),
            ("Dividend", "Enter dividend paid from financial statement"),
            ("Cost of Goods Sold", "Enter cost of goods sold from income statement"),
            ("Inventory", "Enter inventory from balance sheet"),
            ("Receivables", "Enter receivables from balance sheet"),
            ("Payables", "Enter payables from balance sheet"),
            ("Number of Shares", "Enter number of shares outstanding"),
            ("Previous Net Profit", "Enter previous net profit for growth calculation"),
            ("Previous Revenue", "Enter previous revenue for growth calculation"),
            ("Previous Dividend", "Enter previous dividend for growth calculation"),
            ("Previous Total Assets", "Enter previous total assets for growth calculation")
        ]

        for i, (label, tooltip) in enumerate(field_labels):
            row = i // 2 + 1
            col = (i % 2) * 2
            entry_label = tk.Label(self.root, text=label + ":", anchor="w", font=('Arial', 10, 'bold'))
            entry_label.grid(row=row, column=col, sticky=tk.W, padx=10, pady=5)
            
            entry = tk.Entry(self.root, width=30)
            entry.grid(row=row, column=col + 1, padx=10, pady=5)
            entry.bind("<FocusIn>", lambda e, t=tooltip: self.show_tooltip(t))
            entry.bind("<FocusOut>", lambda e: self.hide_tooltip())
            entry.bind("<Return>", lambda e, idx=i: self.focus_next_entry(idx))
            
            self.entries.append(entry)

        submit_button = tk.Button(self.root, text="Submit", command=self.calculate_ratios)
        submit_button.grid(row=len(field_labels) // 2 + 1, column=0, columnspan=4, pady=20)

        self.entries[-1].bind("<Return>", lambda e: submit_button.invoke())

        self.tooltip_label = tk.Label(self.root, text="", font=('Arial', 10, 'italic'))
        self.tooltip_label.grid(row=len(field_labels) // 2 + 2, column=0, columnspan=4)

    def focus_next_entry(self, idx):
        if idx + 1 < len(self.entries):
            self.entries[idx + 1].focus_set()

    def show_tooltip(self, text):
        self.tooltip_label.config(text=text)

    def hide_tooltip(self):
        self.tooltip_label.config(text="")

    def calculate_ratios(self):
        try:
            # Retrieve entered data
            data = [float(entry.get()) for entry in self.entries]
            (revenue, net_profit, total_assets, equity, current_assets, 
             current_liabilities, cash_flow, total_liabilities, capex, ebitda, 
             market_cap, dividend, cogs, inventory, receivables, payables, 
             num_shares, prev_net_profit, prev_revenue, prev_dividend, prev_total_assets) = data

            # Calculate ratios
            ratios = {
                'Profit Margin': net_profit / revenue,
                'Return on Assets (ROA)': net_profit / total_assets,
                'Return on Equity (ROE)': net_profit / equity,
                'Gross Margin': (revenue - cogs) / revenue,
                'Operating Margin': ebitda / revenue,
                'Net Profit Margin': net_profit / revenue,
                'Current Ratio': current_assets / current_liabilities,
                'Quick Ratio': (current_assets - inventory) / current_liabilities,
                'Cash Ratio': cash_flow / current_liabilities,
                'Debt to Equity Ratio': total_liabilities / equity,
                'Debt to Assets Ratio': total_liabilities / total_assets,
                'Interest Coverage Ratio': ebitda / (cogs * 0.05),  # Estimate interest expense
                'Equity Ratio': equity / total_assets,
                'Asset Turnover': revenue / total_assets,
                'Inventory Turnover': revenue / inventory,
                'Receivables Turnover': revenue / receivables,
                'Payables Turnover': revenue / payables,
                'Earnings Per Share (EPS)': net_profit / num_shares,
                'P/E Ratio': market_cap / net_profit,
                'Dividend Yield': dividend / market_cap,
                'EV to EBITDA Ratio': market_cap / ebitda,
                'Earnings Yield': net_profit / market_cap,
                'PEG Ratio': (market_cap / net_profit) / ((net_profit - prev_net_profit) / prev_net_profit),
                'EV to Sales Ratio': market_cap / revenue,
                'Earnings Growth': (net_profit - prev_net_profit) / prev_net_profit,
                'Revenue Growth': (revenue - prev_revenue) / prev_revenue,
                'Dividend Growth Rate': (dividend - prev_dividend) / prev_dividend,
                'Asset Growth': (total_assets - prev_total_assets) / prev_total_assets,
                'Operating Cash Flow to Net Income': cash_flow / net_profit,
                'Free Cash Flow': cash_flow - capex,
                'Operating Cash Flow to Sales': cash_flow / revenue,
                'Cash Flow Coverage Ratio': cash_flow / total_liabilities,
                'Cash Flow Margin': cash_flow / revenue,
                'Retention Ratio': (net_profit - dividend) / net_profit,
                'Capital Gearing Ratio': total_liabilities / equity,
                'Financial Leverage Ratio': total_assets / equity,
                'Debt to Capital Ratio': total_liabilities / (total_liabilities + equity),
                'Book Value per Share': equity / num_shares,
                'Market to Book Ratio': market_cap / equity,
                'Free Cash Flow Yield': (cash_flow - capex) / market_cap,
                'Net Profit Ratio': net_profit / revenue,
                'Company Worth': market_cap / (net_profit / num_shares),
                'Liquidation Value': total_assets - total_liabilities,
                'Recovery Percentage': (total_assets - total_liabilities) / total_liabilities
            }

            company_worth = ratios['Company Worth']
            liquidation_value = ratios['Liquidation Value']
            stock_decision = "BUY" if market_cap < company_worth and ratios['Profit Margin'] > 0.1 else "DO NOT BUY"
            recovery_decision = "SAFE" if liquidation_value > total_liabilities else "RISKY"

            self.show_results(ratios, stock_decision, recovery_decision, company_worth, liquidation_value)

        except ValueError:
            messagebox.showerror("Input Error", "Please ensure all fields contain valid numbers.")

    def show_results(self, ratios, stock_decision, recovery_decision, company_worth, liquidation_value):
        result_window = tk.Toplevel(self.root)
        result_window.title("Financial Analysis Results")

        canvas = tk.Canvas(result_window)
        scrollbar = ttk.Scrollbar(result_window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        result_label = tk.Label(scrollable_frame, text="Financial Ratios", font=('Arial', 16, 'bold'), pady=10)
        result_label.grid(row=0, column=0, columnspan=2)

        row_count = 1
        for i, (ratio, value) in enumerate(ratios.items()):
            color = "green" if self.is_good(ratio, value) else "red"
            result_label = tk.Label(scrollable_frame, text=f"{ratio}: {value:.2f}", font=('Arial', 14, 'bold'),fg=color)
            col = i % 2
            result_label.grid(row=row_count + i // 2, column=col, padx=10, pady=5)

        final_decision_label = tk.Label(scrollable_frame, text=f"Final Decision: {stock_decision}", font=('Arial', 16, 'bold'), fg="blue", pady=20)
        final_decision_label.grid(row=row_count + len(ratios) // 2 + 2, column=0, columnspan=2)

        recovery_label = tk.Label(scrollable_frame, text=f"Recovery Decision: {recovery_decision}", font=('Arial', 16, 'bold'), fg="blue", pady=10)
        recovery_label.grid(row=row_count + len(ratios) // 2 + 3, column=0, columnspan=2)

        save_button = tk.Button(scrollable_frame, text="Save Results to Excel", command=lambda: self.save_to_excel(ratios, stock_decision, recovery_decision, company_worth, liquidation_value))
        save_button.grid(row=row_count + len(ratios) // 2 + 4, column=0, columnspan=2, pady=10)

    def is_good(self, ratio, value):
        thresholds = {
            'Profit Margin': 0.1,
            'Return on Assets (ROA)': 0.05,
            'Return on Equity (ROE)': 0.15,
            'Current Ratio': 1.5,
            'Debt to Equity Ratio': 2,
            'Gross Margin': 0.2,
            'Operating Margin': 0.1,
            'Net Profit Margin': 0.1,
            'Quick Ratio': 1,
            'Cash Ratio': 0.2,
            'Debt to Assets Ratio': 0.5,
            'Interest Coverage Ratio': 3,
            'Equity Ratio': 0.3,
            'Asset Turnover': 1,
            'Inventory Turnover': 5,
            'Receivables Turnover': 8,
            'Payables Turnover': 10,
            'Earnings Per Share (EPS)': 1,
            'P/E Ratio': 20,
            'Dividend Yield': 0.03,
            'EV to EBITDA Ratio': 10,
            'Earnings Yield': 0.05,
            'PEG Ratio': 1,
            'EV to Sales Ratio': 2,
            'Earnings Growth': 0.1,
            'Revenue Growth': 0.1,
            'Dividend Growth Rate': 0.05,
            'Asset Growth': 0.1,
            'Operating Cash Flow to Net Income': 1,
            'Free Cash Flow': 0,
            'Operating Cash Flow to Sales': 0.1,
            'Cash Flow Coverage Ratio': 1,
            'Cash Flow Margin': 0.1,
            'Retention Ratio': 0.5,
            'Capital Gearing Ratio': 0.5,
            'Financial Leverage Ratio': 2,
            'Debt to Capital Ratio': 0.5,
            'Book Value per Share': 10,
            'Market to Book Ratio': 1.5,
            'Free Cash Flow Yield': 0.05,
            'Net Profit Ratio': 0.1,
            'Company Worth': 1,
            'Liquidation Value': 0
        }
        return value >= thresholds.get(ratio, 0)

    def save_to_excel(self, ratios, stock_decision, recovery_decision, company_worth, liquidation_value):
        company_name = simpledialog.askstring("Input", "Enter the company name:")
        if company_name:
            data = { "Financial Data": [entry.get() for entry in self.entries] }
            data.update(ratios)
            data['Company Worth'] = company_worth
            data['Liquidation Value'] = liquidation_value
            data['Stock Decision'] = stock_decision
            data['Recovery Decision'] = recovery_decision

            df = pd.DataFrame(data, index=[0])

            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile=f"{company_name}_analysis.xlsx")
            if file_path:
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Success", f"Financial data saved to {file_path}")
            else:
                messagebox.showerror("Error", "File path not specified!")

# Running the application
if __name__ == "__main__":
    root = tk.Tk()
    app = FinancialAnalysisApp(root)
    root.mainloop()
