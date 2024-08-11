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

# Download and save the file
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

# Send result file and edited:
        await message.answer("Sending you the edited PBA statistics document...")

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


# Logic part for seperate editedPBA file:

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


#Logic part for make result file and send it to user.

        total_power_loss = 0
        inexcusable_stoppages_power = 0
        inexcusable_stoppages_hours = 0
        planned_maintenance_power = 0
        planned_maintenance_hours = 0
        start_row = 19

        # Copy original file to resultFile folder
        file_name = 'original_report.xlsx'
        copied_file_name = 'SEPCOIII Zarafshan Wind Power Project O&M Daily Report.xlsx'

        source_folder = os.path.join(EMPTY_REPORT_DIR, file_name)
        destination_folder = os.path.join(RESULT_FILE_DIR, copied_file_name)

        # Ensure the destination folder exists
        os.makedirs(RESULT_FILE_DIR, exist_ok=True)

        # Copy the file
        try:
            shutil.copy2(source_folder, destination_folder)
        except Exception as e:
            pass

        # Path to the Daily_report file
        filePath = edited_save_path
        daily_report_path = destination_folder

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
                    f'C{row}'] = f"{wtg} {stoppage_type} ,{round(stoppage_time / 60, 2)}h, {random_number} Engineers."
            else:
                daily_report_sheet[f'C{row}'] = f"{wtg} {stoppage_type} ,{round(stoppage_time / 60, 2)}h."
            format_cell(daily_report_sheet[f'C{row}'], horizontal='left')

        def cell_resume_power(row, wtg):
            daily_report_sheet[f'C{row}'] = f"{wtg} resume power generation"
            format_cell(daily_report_sheet[f'C{row}'], horizontal='left')

        def process_excel_data(file_path):
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active

            data = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                data.append(row)

            grouped_data = {}
            for row in data:
                event_type = row[4]  # Assuming event type is in column E
                if event_type not in grouped_data:
                    grouped_data[event_type] = []
                grouped_data[event_type].append(row[0])  # Assuming identifier is in column A

            output_list = []
            for event_type, identifiers in grouped_data.items():
                output_string = f"{', '.join(identifiers)} - {event_type}"
                output_list.append(output_string)

            return output_list

        # Iterate over each WTG number
        for wtg in wtg_number:
            # File path for the current WTG
            file_path = os.path.join(SUB_FILES_DIR,f'{wtg}_rows.xlsx')
            book = openpyxl.load_workbook(file_path)
            sheet = book.active

            start_dates = []
            end_dates = []
            stoppage_summary = {}
            power = 0

            # Iterate through the rows to extract the dates and calculate total power loss
            for row in sheet.iter_rows(min_row=sheet.min_row, values_only=True):  # Start from row 2 to skip the header
                start_date = row[1] if len(row) > 1 else None  # Column B
                end_date = row[2] if len(row) > 2 else None  # Column C
                power = row[7] if len(row) > 7 else None  # Column H
                stoppage_time = row[3] if len(row) > 3 else 0  # Column D
                stoppage_type = row[4] if len(row) > 4 else ''  # Column E

                if isinstance(start_date, datetime):
                    start_dates.append(start_date)
                if isinstance(end_date, datetime):
                    end_dates.append(end_date)
                if isinstance(power, (int, float)):  # Ensure power_loss is numeric
                    total_power_loss += power

                # Accumulate stoppage times by type
                if isinstance(stoppage_time, (int, float)):
                    if stoppage_type in stoppage_summary:
                        stoppage_summary[stoppage_type] += stoppage_time
                    else:
                        stoppage_summary[stoppage_type] = stoppage_time

                # Accumulate stoppage times
                if isinstance(stoppage_time, (int, float)):
                    if 'Fault stop' in stoppage_type:
                        inexcusable_stoppages_hours += stoppage_time
                        inexcusable_stoppages_power += power
                    else:
                        planned_maintenance_hours += stoppage_time
                        planned_maintenance_power += power

            # Calculate the minimum start date and maximum end date
            min_start_date = min(start_dates) if start_dates else None
            max_end_date = max(end_dates) if end_dates else None

            # Determine the maximum stoppage time excluding 'Fault stop'
            max_stoppage_type = None
            max_stoppage_time = 0
            fault_stop_time = stoppage_summary.get('Fault stop', 0)

            if stoppage_summary:
                max_stoppage_type = max(
                    (key for key in stoppage_summary if key != 'Fault stop'),
                    key=lambda k: stoppage_summary[k],
                    default=None
                )
                if max_stoppage_type:
                    max_stoppage_time = stoppage_summary[max_stoppage_type]

            # Write to daily report
            if min_start_date:
                daily_report_sheet[f'B{start_row}'] = min_start_date.strftime("%H:%M")
                format_cell(daily_report_sheet[f'B{start_row}'])

            # Include 'Fault stop' and the next most significant stoppage type
            if fault_stop_time > 0:
                cell_stoppage_type(start_row, wtg, 'Fault stop', fault_stop_time)
                start_row += 1
                if max_stoppage_type:
                    cell_stoppage_type(start_row, wtg, max_stoppage_type, max_stoppage_time)
                    start_row += 1
            elif max_stoppage_type:
                cell_stoppage_type(start_row, wtg, max_stoppage_type, max_stoppage_time)
                start_row += 1

            if max_end_date:
                daily_report_sheet[f'B{start_row}'] = max_end_date.strftime("%H:%M")
                format_cell(daily_report_sheet[f'B{start_row}'])
            cell_resume_power(start_row, wtg)
            start_row += 1
            book.close()

        inexcusable_stoppages_hours /= 60
        planned_maintenance_hours /= 60

        total_power_loss = round(total_power_loss / 1000, 2)
        inexcusable_stoppages_power = round(inexcusable_stoppages_power / 1000, 2)
        planned_maintenance_power = round(planned_maintenance_power / 1000, 2)

        inexcusable_stoppages_hours = round(inexcusable_stoppages_hours, 1)
        planned_maintenance_hours = round(planned_maintenance_hours, 1)

        results = process_excel_data(filePath)
        str = ''
        for result in results:
            str += f'{result} \n'

        if os.access(daily_report_path, os.W_OK):
            daily_report_sheet['H17'] = str

            daily_report_sheet['G17'] = total_power_loss
            if inexcusable_stoppages_power != 0:
                daily_report_sheet['E17'] = f'{inexcusable_stoppages_hours} / {inexcusable_stoppages_power}'
            else:
                daily_report_sheet['E17'] = inexcusable_stoppages_hours
            daily_report_sheet['F17'] = f'{planned_maintenance_hours} / {planned_maintenance_power}'

            daily_report_book.save(daily_report_path)

        # Send result file and edited
        await message.answer("Sending you the Daily report document...")
        file_name = 'SEPCOIII Zarafshan Wind Power Project O&M Daily Report.xlsx'
        file_path = os.path.join(RESULT_FILE_DIR, file_name)

        # Check if the file exists
        if os.path.exists(file_path):
            try:
                # Send the file to the user
                await bot.send_document(message.chat.id, types.InputFile(file_path))
            except Exception as e:
                await message.answer("There was an error sending the file. Please try again later.")
        else:
            await message.answer("The requested file does not exist.")

        # Optionally reset the state after processing
        await state.finish()

#Clean all project files except original file

        # Define the folder containing the files
        Folder_path = [
            SUB_FILES_DIR,
            RESULT_FILE_DIR,
            USER_DOCUMENT_DIR,
            EDITED_USER_FILE_DIR
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

    else:
        await message.answer("Please send a valid Excel file (.xlsx).")
