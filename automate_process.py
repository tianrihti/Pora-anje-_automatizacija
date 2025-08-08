import pandas as pd
import logging
import os
from datetime import datetime, timedelta
from openpyxl import load_workbook
import win32com.client as win32
import time
import pythoncom
import xlwings as xw
import psutil


# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class ExcelAutomation:
    def __init__(self):
        self.pregled_file = os.path.abspath('Pregled.xls')
        self.porocanje_file = os.path.abspath("poročanje proizvodnje2025.xlsm")
        self.plan_file = os.path.abspath("plan brizganja 2025 mesečni.xlsx")

        # Validate files exist
        self._validate_files()

    def _validate_files(self):
        """Validate that all required files exist"""
        files = [self.pregled_file, self.porocanje_file, self.plan_file]
        for file in files:
            if not os.path.exists(file):
                raise FileNotFoundError(f"File not found: {file}")
        
    def step1_copy_pregled_data(self):
        """Step 1: Copy data from Pregled.xls"""
        logger.info("Step 1: Copying data from Pregled.xls")
        
        try:
            # Read Pregled.xls
            try:
                df = pd.read_excel(self.pregled_file, sheet_name="Sheet1", engine='xlrd')
            except ImportError:
                df = pd.read_excel(self.pregled_file, sheet_name="Sheet1", engine='openpyxl')
            
            logger.info(f"Successfully read Pregled.xls with {len(df)} rows")
            return df
            
        except Exception as e:
            logger.error(f"Error reading Pregled.xls: {e}")
            raise
    
    def step2_paste_to_porocanje(self, data_df):
        """Step 2: Paste data into poročanje proizvodnje2025.xlsm sheet 'prilepi gosoft' using openpyxl (safe for macros)"""
        logger.info("Step 2: Pasting data into poročanje proizvodnje2025.xlsm (safe method)")
        try:
            from openpyxl import load_workbook

            # Load the workbook with macros preserved
            wb = load_workbook(self.porocanje_file, keep_vba=True)
            if 'prilepi gosoft' not in wb.sheetnames:
                ws = wb.create_sheet('prilepi gosoft')
            else:
                ws = wb['prilepi gosoft']
            
            # Clear the sheet
            ws.delete_rows(1, ws.max_row)

            # Write headers
            for c_idx, col_name in enumerate(data_df.columns, 1):
                ws.cell(row=1, column=c_idx, value=col_name)

            # Write DataFrame to sheet
            for r_idx, row in enumerate(data_df.values, 2):
                for c_idx, value in enumerate(row, 1):
                    ws.cell(row=r_idx, column=c_idx, value=value)

            wb.save(self.porocanje_file)
            wb.close()
            logger.info(f"Successfully pasted {len(data_df)} rows to 'prilepi gosoft' sheet (safe method)")
            return True
        except Exception as e:
            logger.error(f"Error pasting data to poročanje proizvodnje2025.xlsm (safe method): {e}")
            raise
    
    def get_target_date(self) -> datetime:
        """Get the target date based on the rules"""
        today = datetime.now()
        if today.weekday() == 0:  # Monday
            # Go back to Friday
            target_date = today - timedelta(days=3)
        else:
            # Go back to yesterday
            target_date = today - timedelta(days=1)
        return target_date

    def step3_find_date_in_plan(self):
        """Step 3: Find the correct date in plan sheet"""
        logger.info("Step 3: Finding date in plan sheet")
        
        target_date = self.get_target_date()
        logger.info(f"Looking for date: {target_date.strftime('%Y-%m-%d')}")
        
        try:
            wb = load_workbook(self.plan_file, data_only=True)
            ws = wb["plan"]
            
            # Check row 4 for dates
            date_found = False
            target_col = None
            
            for col in range(1, ws.max_column + 1):
                cell_value = ws.cell(row=4, column=col).value
                if cell_value:
                    # Try to parse as date
                    if isinstance(cell_value, datetime):
                        cell_date = cell_value.date()
                    else:
                        try:
                            cell_date = pd.to_datetime(cell_value).date()
                        except:
                            continue
                    
                    if cell_date == target_date.date():
                        target_col = col
                        date_found = True
                        logger.info(f"Found target date at column {col}")
                        break
            
            if not date_found:
                raise ValueError(f"Target date {target_date.strftime('%Y-%m-%d')} not found in row 4")
            
            return target_col
            
        except Exception as e:
            logger.error(f"Error finding date in plan: {e}")
            raise

    def step4_copy_plan_range(self, start_col):
        """Step 4: Copy range from plan sheet"""
        logger.info(f"Step 4: Copying range from column {start_col}")

        try:
            wb = load_workbook(self.plan_file, data_only=True)
            ws = wb["plan"]

            copied_data = []  # Will store list of rows, each row is a list of 3 cell values

            for row in range(6, 45):  # Rows 6 to 44 inclusive
                row_data = []
                for col_offset in range(3):  # Copy 3 columns: start_col, start_col+1, start_col+2
                    cell = ws.cell(row=row, column=start_col + col_offset)
                    row_data.append(cell.value)
                copied_data.append(row_data)

            logger.info(f"Copied range from column {start_col} to {start_col+2}, rows 6 to 44")
            return copied_data

        except Exception as e:
            logger.error(f"Error copying plan range: {e}")
            raise
    
    def step5_paste_to_brizganje(self, copied_data):
        try:
            logger.info("Step 5: Pasting data into 'brizganje izracun' sheet")

            # Load workbook
            wb = load_workbook(self.porocanje_file, keep_vba=True)
            ws = wb["brizganje izračun"]

            """Get the target date based on the rules"""
            today = datetime.now()
            if today.weekday() == 0:  # Monday
                # Go back to Friday
                target_date = today - timedelta(days=4)
            else:
                # Go back to yesterday
                target_date = today - timedelta(days=2)

            logger.info(f"Looking for date: {target_date.strftime('%Y-%m-%d')}")

            # Search row 4, starting from column D (index 4)
            target_col = None
            for col in range(1, ws.max_column + 1):
                cell_value = ws.cell(row=4, column=col).value
                if cell_value:
                    if isinstance(cell_value, datetime):
                        cell_date = cell_value.date()
                    else:
                        try:
                            cell_date = pd.to_datetime(cell_value).date()
                        except:
                            continue

                    if cell_date == target_date.date():
                        target_col = col
                        logger.info(f"Found target date in column {col}")
                        break

            if target_col is None:
                raise ValueError(f"Target date {target_date.strftime('%Y-%m-%d')} not found in row 4")

            paste_col = target_col - 1  # One column to the left

            logger.info(f"Pasting values into column {paste_col}")

            # Paste copied_data into brizganje izracun sheet
            start_row = 4  # Assuming we start from row 4 (just like when copying)
            for i, row_data in enumerate(copied_data):
                for j, value in enumerate(row_data):
                    ws.cell(row=start_row + i, column=paste_col + j, value=value)

            # Save the workbook
            wb.save(self.porocanje_file)
            wb.close()
            logger.info("Successfully pasted values to 'brizganje izracun' sheet")

        except Exception as e:
            logger.error(f"Error in Step 5: {e}")
            raise

    def step6_analyze_brizganje(self):
        try:
            logger.info("Step 6: Analyzing rows 7 to 46 in 'brizganje izračun'")

            # Load workbook with formula results
            wb = load_workbook(self.porocanje_file, data_only=True)
            ws = wb["brizganje izračun"]

            saved_texts = []
            
            logger.setLevel(logging.DEBUG)
            for row in range(7, 47): # Rows 7 to 46 inclusive
                col_a = ws.cell(row=row, column=1).value  # Column A
                col_l = ws.cell(row=row, column=12).value  # Column L
                col_m = ws.cell(row=row, column=13).value  # Column M
                col_x = ws.cell(row=row, column=24).value  # Column X
                col_y = ws.cell(row=row, column=25).value  # Column Y

                # Log what we’re reading:
                logger.debug(f"Row {row}: A='{col_a}', L='{col_l}', M='{col_m}', X='{col_x}', Y='{col_y}'")

                cell_a = ws.cell(row=row, column=1).value  # Column A
                if not cell_a:
                    continue  # Skip if column A is empty

                cell_l = ws.cell(row=row, column=12).value  # Column L (12)
                if cell_l in (None, ""):
                    continue  # Skip if column L is empty or zero

                cell_m = ws.cell(row=row, column=13).value  # Column M (13)
                # Try converting to float if possible
                try:
                    if isinstance(cell_m, str):
                        # Remove euro sign and replace comma with dot, if needed
                        cleaned = cell_m.replace("€", "").replace(",", ".").strip()
                        value_m = float(cleaned)
                    else:
                        value_m = float(cell_m)
                except (TypeError, ValueError):
                    continue  # Skip if not a number

                if value_m > 50:
                    saved_texts.append(str(cell_a))  # Save value from column A

            logger.info(f"Saved texts: {saved_texts}")
            return saved_texts

        except Exception as e:
            logger.error(f"Error in Step 6: {e}")
            raise

    def recalc_excel(self):
        try:
            app = xw.App(visible=False)
            wb = app.books.open(self.porocanje_file)
            wb.app.calculate()
            wb.save()
            wb.close()
            app.quit()
        except Exception as e:
            logger.error(f"Error in recalc_excel: {e}")
        finally:
            try:
                app.quit()
            except:
                pass
            self.kill_excel_processes()

    def kill_excel_processes(self):
        time.sleep(2)  # Wait for 2 seconds before killing processes
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] in ['EXCEL.EXE', 'excel.exe']:
                try:
                    proc.terminate()
                    proc.wait(timeout=3)  # Wait for up to 3 seconds for the process to terminate
                    logger.info(f"Terminated Excel process: {proc.pid}")
                except psutil.TimeoutExpired:
                    proc.kill()  # Force kill if it doesn't terminate
                    logger.warning(f"Force killed Excel process: {proc.pid}")
                except:
                    logger.warning(f"Failed to terminate Excel process: {proc.pid}")


if __name__ == "__main__":
    # Create automation instance
    automation = ExcelAutomation()
    try:
        # Run step 1
        pregled_data = automation.step1_copy_pregled_data()

        # Run step 2
        automation.step2_paste_to_porocanje(pregled_data)

        # Run step 3
        target_col = automation.step3_find_date_in_plan()

        # Run step 4
        plan_range_data = automation.step4_copy_plan_range(target_col)

        # Run step 5
        automation.step5_paste_to_brizganje(plan_range_data)

        automation.recalc_excel()
        # Run step 6
        saved_texts = automation.step6_analyze_brizganje()
    except Exception as e:
        logger.error((f"An error occurred: {e}"))
    finally:
        automation.kill_excel_processes()
        logger.info("Script execution finished")