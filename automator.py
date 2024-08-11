# automator.py

from pathlib import Path
import create_new_sales_update
import process_estimate_summary
import update_scoreboard_24
import update_scoreboard_23
import update_sales_report

def run_automation(import_path):
    print("Starting Sales Update Automation")

    print("\n1. Creating new Sales Update file")
    new_file_path = create_new_sales_update.create_new_sales_update(import_path)
    print(f"New Sales Update file created: {new_file_path}")

    print("\n2. Processing Estimate Summary")
    process_estimate_summary.process_estimate_summary(import_path)
    print("Estimate Summary processing completed")

    print("\n3. Updating 2024 Scoreboard")
    update_scoreboard_24.update_scoreboard(import_path)
    print("2024 Scoreboard update completed")

    print("\n4. Updating 2023 Scoreboard")
    update_scoreboard_23.update_scoreboard(import_path)
    print("2023 Scoreboard update completed")

    print("\n5. Updating Sales Report")
    update_sales_report.update_sales_report(import_path)
    print("Customer Count Updated")

    print("\nSales Update Automation completed successfully!")

if __name__ == "__main__":
    # Set the path to your directory containing the Sales Update files
    import_path = Path("C:/Users/Shadow/projects/analyitics_automator")
    
    run_automation(import_path)