import json
import forecastTemplate
import salesReport
import expenseReport
import userReport


################### PATH MANAGER ###################
JSON_FILE_PATH = "C:/Users/Public/paths.json"

# Read the JSON file and load its content into the 'paths' dictionary
try:
    with open(JSON_FILE_PATH, 'r', encoding='utf-8') as f:
        paths = json.load(f)
except FileNotFoundError:
    print(f"Error: The JSON file was not found at '{JSON_FILE_PATH}'. Please check the path.")
    exit()
except json.JSONDecodeError:
    print(f"Error: Could not decode JSON from '{JSON_FILE_PATH}'. Please check the file's content.")
    exit()
except Exception as e:
    print(f"An unexpected error occurred while loading the JSON file: {e}")
    exit()

######################################################################



# --- Main execution block ---
if __name__ == "__main__":
    original_excel_file_path = paths['ExpenseReport']
    template_excel_file_path = paths['templateExpenseReport']
    output_folder = paths['outputExpenseReport']
    history_folder = paths['expenseReportHistory']
    sales_report_path = paths['salesDataStorage']
    expense_report_data_actual_path = paths['expenseReportDataActual']
    forecast_access_db_path = paths['statementsFP&A']
    forecast_template_excel_path = paths['expenseForecastTemplate']
    forecast_output_folder = paths['outputForecastTemplate']
    REPORT_DIRECTORY = paths['usersReport']

    #---------- EXPENSE REPORT ----------
    # expenseReport.move_files_to_history(output_folder, history_folder, template_excel_file_path)
    # expenseReport.process_expense_reports(original_excel_file_path, template_excel_file_path, output_folder)
    # expenseReport.refresh_excel_files_in_folder(output_folder)
    # final_total_expense_soles = expenseReport.calculate_total_saldo_soles(output_folder, template_excel_file_path)
    # print(f"\nFINAL TOTAL EXPENSE FOR THE PERIOD: {final_total_expense_soles:,.2f}")

    #---------- SALES REPORT ----------
    # period_input = str(input("Please enter the period (e.g., '2023-12'): "))
    # salesReport.process_sales_reports(expense_report_data_actual_path, sales_report_path, period_input)
    # expenseReport.move_files_to_history(sales_report_path, history_folder)

    #---------- FORECAST TEMPLATE GENERATOR ----------
    # forecastTemplate.template_forecast_generator(access_db_path=forecast_access_db_path, excel_template_path=forecast_template_excel_path, output_folder_path=forecast_output_folder)


    #---------- USER REPORT ----------
    userReport.process_and_alert_client_data(REPORT_DIRECTORY)