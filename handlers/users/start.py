from aiogram import types
from aiogram.dispatcher.filters.builtin import CommandStart
from aiogram.dispatcher import FSMContext
from loader import dp, bot
from states.userState import fileData

import openpyxl
from openpyxl import load_workbook, Workbook
import warnings
from datetime import datetime
import os
import random
from openpyxl.styles import Font, Alignment
import shutil

# Define base directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Define save directories
USER_DOCUMENT_DIR = os.path.join(BASE_DIR, 'projectFiles', 'userDocument')
EDITED_USER_FILE_DIR = os.path.join(BASE_DIR, 'projectFiles', 'editedUserFile')
SUB_FILES_DIR = os.path.join(BASE_DIR, 'projectFiles', 'subFiles')
RESULT_FILE_DIR = os.path.join(BASE_DIR, 'projectFiles', 'resultFile')
EMPTY_REPORT_DIR = os.path.join(BASE_DIR, 'projectFiles', 'empty_report')

# Ensure directories exist
os.makedirs(USER_DOCUMENT_DIR, exist_ok=True)
os.makedirs(EDITED_USER_FILE_DIR, exist_ok=True)
os.makedirs(SUB_FILES_DIR, exist_ok=True)
os.makedirs(RESULT_FILE_DIR, exist_ok=True)
os.makedirs(EMPTY_REPORT_DIR, exist_ok=True)

@dp.message_handler(CommandStart())
async def bot_start(message: types.Message):
    await message.answer(f"Hello, {message.from_user.full_name}! \n"
                         f"Send me the original PBA statistics in an Excel file.")
    await fileData.file.set()  # Set the state to indicate we are expecting a file

@dp.message_handler(content_types=types.ContentType.DOCUMENT, state=fileData.file)
async def handle_document(message: types.Message, state: FSMContext):
    if message.document and message.document.file_name.endswith('.xlsx'):

        # Download the file
        file_info = await bot.get_file(message.document.file_id)
        file_name = message.document.file_name
        file_path = file_info.file_path
        downloaded_file = await bot.download_file(file_path)

        # Define the directory and file path to save the file
        save_path = os.path.join(USER_DOCUMENT_DIR, message.document.file_name)

        # Save the file
        with open(save_path, 'wb') as new_file:
            new_file.write(downloaded_file.getvalue())

        await message.answer("Thank you for sending the Excel file. Processing...")

        # Logic part for edit excel file:

        # Option to suppress specific warnings
        warnings.filterwarnings("ignore", category=UserWarning)

        # Load the workbook and select the active worksheet
        book = load_workbook(save_path)
        sheet = book.active

        # Create a new workbook and copy data to ensure styles are consistent
        new_book = Workbook()
        new_sheet = new_book.active

        # Copy the header row to the new sheet
        header = [cell.value for cell in sheet[1]]
        new_sheet.append(header)

        # Copy data and apply styles
        for row in sheet.iter_rows(values_only=True):
            new_sheet.append(row)

        # Center align all cells
        for row in new_sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center')

        # Delete unnecessary columns in reverse order
        need_delete_columns = [1, 2, 4, 5, 6, 7, 10, 12, 14, 15, 18, 19, 21]
        need_delete_columns.sort(reverse=True)
        for col in need_delete_columns:
            new_sheet.delete_cols(col)

        # Change column width
        column_widths = {'A': 25, 'B': 25, 'C': 25, 'D': 25, 'E': 45, 'F': 25, 'G': 25, 'H': 25}
        for col, width in column_widths.items():
            new_sheet.column_dimensions[col].width = width

        # Define the target values to keep
        target_values = [
            'Fault stop',
            'Tower base stop',
            'Service mode',
            'Periodic service stop',
            'Owner or Utility Requested Stop',
            'Visit/Inspection/Training Stop by Owner or Utility'
        ]

        # Collect rows to delete
        rows_to_delete = []

        for row in new_sheet.iter_rows(min_row=2, max_col=new_sheet.max_column,
                                       max_row=new_sheet.max_row):  # Skip header
            cell = row[4]  # Get the cell in column 'E'
            if cell.value is None or cell.value not in target_values:
                rows_to_delete.append(cell.row)

        # Delete rows in reverse order to avoid shifting issues
        for row in reversed(rows_to_delete):
            new_sheet.delete_rows(row)

        new_sheet.delete_rows(1)
        # Save the new workbook
        edited_save_path = os.path.join(EDITED_USER_FILE_DIR, 'editedPBA.xlsx')
        new_book.save(edited_save_path)

        # Send result file and edited
        await message.answer("Sending you the edited PBA statistics document...")
        file_name = 'editedPBA.xlsx'  # Replace with your actual file name

        # Check if the file exists
        if os.path.exists(edited_save_path):
            try:
                # Send the file to the user
                await bot.send_document(message.chat.id, types.InputFile(edited_save_path))
            except Exception as e:
                await message.answer("There was an error sending the file. Please try again later.")
        else:
            await message.answer("The requested file does not exist.")

        await message.answer("Wait a second. 'SEPCOIII Zarafshan Wind Power Project O&M Daily Report' file is under processing...")

        # Logic part for separate editedPBA file
        book = load_workbook(edited_save_path)
        sheet = book.active

        # Create the directory if it doesn't exist
        os.makedirs(SUB_FILES_DIR, exist_ok=True)

        # Get the column A values
        column_values = [cell.value for cell in sheet['A'] if cell.value is not None]

        # Find unique WTG numbers
        wtg_number = list(set(column_values))

        # Iterate over each unique WTG number
        for wtg in wtg_number:
            # Create a new workbook for each WTG number
            new_book = Workbook()
            new_sheet = new_book.active

            # Set the column width to 30 for all columns in the new sheet
            for col in sheet.iter_cols():
                new_sheet.column_dimensions[col[0].column_letter].width = 25

            # Copy the rows where column A matches the current WTG number
            for row in sheet.iter_rows(min_row=1):
                if row[0].value == wtg:
                    new_sheet.append([cell.value for cell in row])

            # Save the new workbook
            file_path = os.path.join(SUB_FILES_DIR, f'{wtg}_rows.xlsx')

            try:
                # Save the new workbook
                new_book.save(file_path)
            except Exception as e:
                pass

        # Logic part for make result file and send it to user.
        # total_power_loss = 0
        # inexcusable_stoppages_power = 0
        # inexcusable_stoppages_hours = 0
        # planned_maintenance_power = 0
        # planned_maintenance_hours = 0
        start_row = 19

        # Copy original file to resultFile folder
        file_name = 'original_report.xlsx'
        copied_file_name = 'SEPCOIII Zarafshan Wind Power Project O&M Daily Report.xlsx'

        # Construct the full paths
        source_path = os.path.join(EMPTY_REPORT_DIR, file_name)
        destination_path = os.path.join(RESULT_FILE_DIR, copied_file_name)

        # Ensure the destination folder exists
        os.makedirs(RESULT_FILE_DIR, exist_ok=True)

        # Check if the source file exists
        if not os.path.exists(source_path):
            await message.answer(f"The source file {file_name} does not exist in the directory {EMPTY_REPORT_DIR}.")
            return

        # Copy the file
        try:
            shutil.copy2(source_path, destination_path)
        except Exception as e:
            await message.answer(f"Error copying the file: {str(e)}")
            return

        # Path to the Daily_report file
        filePath = edited_save_path
        daily_report_path = destination_path

        daily_report_book = openpyxl.load_workbook(daily_report_path)
        daily_report_sheet = daily_report_book.active

        def format_cell(cell, bold=False, size=12, horizontal='center', vertical='center'):
            cell.font = Font(bold=bold, size=size)
            cell.alignment = Alignment(horizontal=horizontal, vertical=vertical)

        def cell_stoppage_type(row, wtg, stoppage_type, stoppage_time):
            random_number = random.randint(2, 4)
            if stoppage_type not in ['Owner or Utility Requested Stop',
                                     'Visit/Inspection/Training Stop by Owner or Utility']:
                daily_report_sheet[
                    f'C{row}'] = f"{wtg} {stoppage_type} ,{round(stoppage_time / 60, 2)}h, {random_number}MWh"
            else:
                daily_report_sheet[
                    f'C{row}'] = f"{wtg} {stoppage_type} ,{round(stoppage_time / 60, 2)}h"
            format_cell(daily_report_sheet[f'C{row}'], size=10)

        def cell_resume_power(row, wtg):
            daily_report_sheet[f'C{row}'] = f"{wtg} Resume power generation"
            format_cell(daily_report_sheet[f'C{row}'], size=10)

        def process_excel_data(file_path, daily_report_sheet, start_row):
            global total_power_loss, inexcusable_stoppages_power, inexcusable_stoppages_hours, planned_maintenance_power, planned_maintenance_hours

            total_power_loss = 0
            inexcusable_stoppages_power = 0
            inexcusable_stoppages_hours = 0
            planned_maintenance_power = 0
            planned_maintenance_hours = 0

            book = openpyxl.load_workbook(file_path)
            sheet = book.active

            merged_data = {}
            unique_wtgs = set()

            for row in sheet.iter_rows(min_row=2, values_only=True):
                wtg = row[0]
                stop_start_time = row[1]
                stoppage_time = row[2]
                stoppage_type = row[4]

                # Check and convert stoppage_time to an integer if it's not already
                if isinstance(stoppage_time, datetime):
                    stoppage_time = (stop_start_time - stoppage_time).total_seconds() / 60
                elif isinstance(stoppage_time, str):
                    # If stoppage_time is a string, try to convert it to an integer
                    try:
                        stoppage_time = int(stoppage_time)
                    except ValueError:
                        continue  # Skip this row if conversion fails

                if not isinstance(stoppage_time, (int, float)):
                    continue  # Skip rows with invalid stoppage_time

                if wtg not in merged_data:
                    merged_data[wtg] = {}
                if stoppage_type not in merged_data[wtg]:
                    merged_data[wtg][stoppage_type] = {
                        'total_stoppage_time': 0,
                        'first_stop_start_time': stop_start_time,
                        'total_power_loss': 0
                    }
                merged_data[wtg][stoppage_type]['total_stoppage_time'] += stoppage_time

                if stop_start_time < merged_data[wtg][stoppage_type]['first_stop_start_time']:
                    merged_data[wtg][stoppage_type]['first_stop_start_time'] = stop_start_time

                if stoppage_type != 'Fault stop':
                    merged_data[wtg][stoppage_type]['total_power_loss'] += round(stoppage_time * 0.75)

                unique_wtgs.add(wtg)

            for wtg in unique_wtgs:
                if 'Fault stop' in merged_data[wtg]:
                    fault_stoppage_time = merged_data[wtg]['Fault stop']['total_stoppage_time']
                    stoppage_types = list(merged_data[wtg].keys())
                    stoppage_types.remove('Fault stop')

                    if stoppage_types:
                        max_stoppage_type = max(stoppage_types,
                                                key=lambda x: merged_data[wtg][x]['total_stoppage_time'])
                        max_stoppage_time = merged_data[wtg][max_stoppage_type]['total_stoppage_time']
                        max_stoppage_start_time = merged_data[wtg][max_stoppage_type]['first_stop_start_time']

                        if fault_stoppage_time > max_stoppage_time:
                            stoppage_time = fault_stoppage_time
                            stoppage_type = 'Fault stop'
                            stoppage_start_time = merged_data[wtg]['Fault stop']['first_stop_start_time']
                        else:
                            stoppage_time = max_stoppage_time
                            stoppage_type = max_stoppage_type
                            stoppage_start_time = max_stoppage_start_time
                    else:
                        stoppage_time = fault_stoppage_time
                        stoppage_type = 'Fault stop'
                        stoppage_start_time = merged_data[wtg]['Fault stop']['first_stop_start_time']
                else:
                    stoppage_types = list(merged_data[wtg].keys())
                    if stoppage_types:
                        max_stoppage_type = max(stoppage_types,
                                                key=lambda x: merged_data[wtg][x]['total_stoppage_time'])
                        stoppage_time = merged_data[wtg][max_stoppage_type]['total_stoppage_time']
                        stoppage_type = max_stoppage_type
                        stoppage_start_time = merged_data[wtg][max_stoppage_type]['first_stop_start_time']
                    else:
                        # Skip this WTG if there are no stoppage types
                        continue

                stoppage_hours = round(stoppage_time / 60, 2)
                row = start_row

                daily_report_sheet[f'B{row}'] = stoppage_start_time.strftime('%Y-%m-%d')
                format_cell(daily_report_sheet[f'B{row}'], size=10)

                cell_stoppage_type(row, wtg, stoppage_type, stoppage_time)

                row += 1
                daily_report_sheet[f'B{row}'] = stoppage_start_time.strftime('%H:%M')
                format_cell(daily_report_sheet[f'B{row}'], size=10)

                cell_resume_power(row, wtg)

                if stoppage_type in [
                    'Owner or Utility Requested Stop',
                    'Visit/Inspection/Training Stop by Owner or Utility'
                ]:
                    inexcusable_stoppages_power += round(stoppage_time * 0.75)
                    inexcusable_stoppages_hours += stoppage_hours
                elif stoppage_type == 'Periodic service stop':
                    planned_maintenance_power += round(stoppage_time * 0.75)
                    planned_maintenance_hours += stoppage_hours
                else:
                    total_power_loss += round(stoppage_time * 0.75)

                row += 2

            book.close()

        # Call the function to process Excel data
        process_excel_data(filePath, daily_report_sheet, start_row)

        # Save the Daily report
        daily_report_book.save(daily_report_path)

        # Send result file
        await message.answer("Sending you the SEPCOIII Zarafshan Wind Power Project O&M Daily Report document...")

        try:
            # Send the file to the user
            await bot.send_document(message.chat.id, types.InputFile(daily_report_path))
        except Exception as e:
            await message.answer("There was an error sending the file. Please try again later.")

        # Reset the state
        await state.finish()

        # Clean all project files except original file

        # Define the folder containing the files
        Folder_path = [
            RESULT_FILE_DIR,
            SUB_FILES_DIR,
            EDITED_USER_FILE_DIR,
            USER_DOCUMENT_DIR
        ]
        # Ensure the folder exists
        for folder_path in Folder_path:
            if os.path.exists(folder_path):
                # Iterate over all the files in the folder
                for file_name in os.listdir(folder_path):
                    # Construct the full file path
                    file_path = os.path.join(folder_path, file_name)

                    # Check if it's a file (not a directory)
                    if os.path.isfile(file_path):
                        try:
                            os.remove(file_path)  # Delete the file
                        except Exception as e:
                            pass

        await message.answer("Thank you for your patience. The files have been processed successfully.")
    else:
        await message.answer("Please send a valid Excel file (.xlsx).")
