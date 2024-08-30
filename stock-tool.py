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
            ("Profit", "Enter net profit after tax from income statement"),
            ("Current Assets", "Enter current assets from balance sheet"),
            ("Non-current Assets", "Enter non-current assets from balance sheet"),
            ("Total Assets", "Enter total assets from balance sheet"),
            ("Current Liabilities", "Enter current liabilities from balance sheet"),
            ("Non-current Liabilities", "Enter non-current liabilities from balance sheet"),
            ("Total Liabilities", "Enter total liabilities from balance sheet"),
            ("Equity", "Enter shareholders' equity from balance sheet"),
            ("Operating Expenses", "Enter operating expenses from income statement"),
            ("Cash Flow", "Enter cash flow from cash flow statement"),
            ("Depreciation", "Enter depreciation and amortization from income statement"),
            ("Interest Expense", "Enter interest expense from income statement"),
            ("Inventory", "Enter inventory from balance sheet"),
            ("Receivables", "Enter receivables from balance sheet"),
            ("Payables", "Enter payables from balance sheet"),
            ("Capital Expenditure", "Enter capital expenditure from cash flow statement"),
            ("EBIT", "Enter earnings before interest and taxes from income statement"),
            ("EBT", "Enter earnings before tax from income statement"),
            ("Tax", "Enter tax paid from income statement"),
            ("Market Cap", "Enter market capitalization of the company"),
            ("Dividend", "Enter dividend paid from financial statement")
        ]

        # Create field labels and text entry widgets
        for i, (label, tooltip) in enumerate(field_labels):
            row = i // 2 + 1
            col = (i % 2) * 2
            entry_label = tk.Label(self.root, text=label + ":", anchor="w", font=('Arial', 10, 'bold'))
            entry_label.grid(row=row, column=col, sticky=tk.W, padx=10, pady=5)
            
            entry = tk.Entry(self.root, width=30)
            entry.grid(row=row, column=col + 1, padx=10, pady=5)
            entry.bind("<FocusIn>", lambda e, t=tooltip: self.show_tooltip(t))  # Show tooltip on focus
            entry.bind("<FocusOut>", lambda e: self.hide_tooltip())  # Hide tooltip when unfocused
            entry.bind("<Return>", lambda e, idx=i: self.focus_next_entry(idx))  # Handle Enter key for field navigation
            
            self.entries.append(entry)

        # Submit button
        submit_button = tk.Button(self.root, text="Submit", command=self.calculate_ratios)
        submit_button.grid(row=len(field_labels) // 2 + 1, column=0, columnspan=4, pady=20)

        self.entries[-1].bind("<Return>", lambda e: submit_button.invoke())  # Submit on Enter in last field

        # Tooltip label at the bottom
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
            # Retrieve entered data and perform ratio calculations
            revenue = float(self.entries[0].get())
            profit = float(self.entries[1].get())
            current_assets = float(self.entries[2].get())
            non_current_assets = float(self.entries[3].get())
            total_assets = float(self.entries[4].get())
            current_liabilities = float(self.entries[5].get())
            non_current_liabilities = float(self.entries[6].get())
            total_liabilities = float(self.entries[7].get())
            equity = float(self.entries[8].get())
            operating_expenses = float(self.entries[9].get())
            cash_flow = float(self.entries[10].get())
            depreciation = float(self.entries[11].get())
            interest_expense = float(self.entries[12].get())
            inventory = float(self.entries[13].get())
            receivables = float(self.entries[14].get())
            payables = float(self.entries[15].get())
            capex = float(self.entries[16].get())
            ebit = float(self.entries[17].get())
            ebt = float(self.entries[18].get())
            tax = float(self.entries[19].get())
            market_cap = float(self.entries[20].get())
            dividend = float(self.entries[21].get())

            # Financial Ratio Calculations
            ratios = {}
            ratios['Profit Margin'] = profit / revenue
            ratios['Return on Assets (ROA)'] = profit / total_assets
            ratios['Return on Equity (ROE)'] = profit / equity
            ratios['Gross Margin'] = (revenue - operating_expenses) / revenue
            ratios['Operating Margin'] = ebit / revenue
            ratios['Net Profit Margin'] = profit / revenue
            ratios['Current Ratio'] = current_assets / current_liabilities
            ratios['Quick Ratio'] = (current_assets - inventory) / current_liabilities
            ratios['Cash Ratio'] = cash_flow / current_liabilities
            ratios['Debt to Equity Ratio'] = total_liabilities / equity
            ratios['Debt to Assets Ratio'] = total_liabilities / total_assets
            ratios['Interest Coverage Ratio'] = ebit / interest_expense
            ratios['Equity Ratio'] = equity / total_assets
            ratios['Asset Turnover'] = revenue / total_assets
            ratios['Inventory Turnover'] = revenue / inventory
            ratios['Receivables Turnover'] = revenue / receivables
            ratios['Payables Turnover'] = revenue / payables
            ratios['Earnings Per Share (EPS)'] = profit / market_cap
            ratios['P/E Ratio'] = market_cap / profit
            ratios['Dividend Yield'] = dividend / market_cap

            # New Features: Company Worth and Recovery Calculation
            company_worth = ratios['P/E Ratio'] * profit
            liquidation_value = current_assets - current_liabilities

            # Buy stock decision based on whether it's below worth
            stock_decision = "BUY" if market_cap < company_worth and ratios['Profit Margin'] > 0.1 else "DO NOT BUY"

            # Recovery decision based on liquidation value
            recovery_decision = "SAFE" if liquidation_value > total_liabilities else "RISKY"

            # Display the results in a new window
            self.show_results(ratios, stock_decision, recovery_decision, company_worth, liquidation_value)

        except ValueError:
            messagebox.showerror("Input Error", "Please ensure all fields contain valid numbers.")

    def show_results(self, ratios, stock_decision, recovery_decision, company_worth, liquidation_value):
        result_window = tk.Toplevel(self.root)
        result_window.title("Financial Analysis Results")

        # Scrollable results window with two columns
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

        result_label = tk.Label(scrollable_frame, text="Financial Ratios", font=('Arial', 14, 'bold'), pady=10)
        result_label.grid(row=0, column=0, columnspan=2)

        # Display ratios in two columns
        row_count = 1
        for i, (ratio, value) in enumerate(ratios.items()):
            color = "green" if self.is_good(ratio, value) else "red"
            result_label = tk.Label(scrollable_frame, text=f"{ratio}: {value:.2f}", fg=color)
            col = i % 2  # Alternate between two columns
            result_label.grid(row=row_count + i // 2, column=col, padx=10, pady=5)

        # Final decision in center
        final_decision_label = tk.Label(scrollable_frame, text=f"Final Decision: {stock_decision}", font=('Arial', 16, 'bold'), fg="blue", pady=20)
        final_decision_label.grid(row=row_count + len(ratios) // 2 + 2, column=0, columnspan=2)

        recovery_label = tk.Label(scrollable_frame, text=f"Recovery Decision: {recovery_decision}", font=('Arial', 16, 'bold'), fg="blue", pady=10)
        recovery_label.grid(row=row_count + len(ratios) // 2 + 3, column=0, columnspan=2)

        # Save results button
        save_button = tk.Button(scrollable_frame, text="Save Results to Excel", command=lambda: self.save_to_excel(ratios, stock_decision, recovery_decision, company_worth, liquidation_value))
        save_button.grid(row=row_count + len(ratios) // 2 + 4, column=0, columnspan=2, pady=10)

    def is_good(self, ratio, value):
        # Example thresholds for "good" and "bad" interpretation
        thresholds = {
            'Profit Margin': 0.1,
            'Return on Assets (ROA)': 0.05,
            'Return on Equity (ROE)': 0.15,
            'Current Ratio': 1.5,
            'Debt to Equity Ratio': 2
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
