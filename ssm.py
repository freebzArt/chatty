"""
Script Name: Shot Sheet Maker
Script Version: 3.13.0
Flame Version: 2025
Written by: Michael Vaglienty

Custom Action Type: Media Panel

Description:
    Create shot sheets from selected sequence clips that can be loaded into Excel, Google Sheets, or Numbers.

    Sequence should have all clips on one version/track.

    *** First time script is run it will need to install xlsxWriter - System password required for this ***
    This will need to happen for each new version of Flame.

Menus:
    Right-click on selected sequences in media panel -> Export Shot Sheet

To install:
    Copy script into /opt/Autodesk/shared/python/shot_sheet_maker
"""

#-------------------------------------
# [Imports]
#-------------------------------------

import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from collections import OrderedDict

import flame
from lib.pyflame_lib_shot_sheet_maker import *
from xlsxwriter.utility import xl_rowcol_to_cell

#-------------------------------------
# [Constants]
#-------------------------------------

SCRIPT_NAME = 'Shot Sheet Maker'
SCRIPT_VERSION = 'v3.13.0'
SCRIPT_PATH = os.path.abspath(os.path.dirname(__file__))

#-------------------------------------
# [Main Script]
#-------------------------------------

class ShotSheetMaker:

    def __init__(self, selection):

        pyflame.print_title(f'{SCRIPT_NAME} {SCRIPT_VERSION}')

        # Check script path, if path is incorrect, stop script.
        if not pyflame.verify_script_install():
            return

        # Create/Load config file settings.
        self.settings = self.load_config()

        self.selection = selection

        # Make sure sequences only have one version/track
        for item in self.selection:
            if len(item.versions) > 1:
                PyFlameMessageWindow(
                    message='Sequences can only have one version/track',
                    message_type=MessageType.ERROR,
                    parent=None,
                    )
                return
            elif len(item.versions[0].tracks) > 1:
                PyFlameMessageWindow(
                    message='Sequences can only have one version/track',
                    message_type=MessageType.ERROR,
                    parent=None,
                    )
                return

        # Make sure no two sequences are named the same
        sequence_names = [str(seq.name)[1:-1] for seq in self.selection]
        duplicate_names = [name for name in sequence_names if sequence_names.count(name) > 1]
        if duplicate_names:
            PyFlameMessageWindow(
                message='No two sequences can have the same name.\n\nRename sequences and try again.',
                message_type=MessageType.ERROR,
                parent=None,
                )
            return

        # Define paths
        self.temp_path = os.path.join(SCRIPT_PATH, 'temp')

        # Create temp directory. If it already exists, delete it and create a new one.
        if os.path.isdir(self.temp_path):
            shutil.rmtree(self.temp_path)
        os.mkdir(self.temp_path)

        # Get Flame variables
        self.flame_project_name = flame.projects.current_project.name
        self.current_flame_version = flame.get_version()

        # Initialize misc. variables
        self.thumb_nail_height = ''
        self.thumb_nail_width = ''
        self.x_offset = ''
        self.y_offset = ''
        self.column_width = ''
        self.row_height = ''
        self.temp_image_path = ''

        # Initialize exporter
        self.exporter = flame.PyExporter()
        self.exporter.foreground = True
        self.exporter.export_between_marks = True

        # Get jpeg export preset
        preset_dir = flame.PyExporter.get_presets_dir(
            flame.PyExporter.PresetVisibility.Autodesk,
            flame.PyExporter.PresetType.Image_Sequence)
        jpg_preset_path = os.path.join(preset_dir, "Jpeg", "Jpeg (8-bit).xml")

        # Copy jpeg export preset to temp directory
        self.temp_export_preset = os.path.join(self.temp_path, 'Temp_Export_Preset.xml')
        shutil.copy(jpg_preset_path, self.temp_export_preset)

        # Make sure xlsxwriter is installed, if not, install it. Otherwise, open window.
        xlsxwriter_installed = self.xlsxwriter_check()

        if xlsxwriter_installed:
            return self.main_window()
        return self.install_xlsxwriter()

    def load_config(self) -> PyFlameConfig:
        settings = PyFlameConfig(
            config_values={
                'export_path': '/opt/Autodesk',
                'thumbnail_size': 'Medium',
                'one_workbook': True,
                'reveal_in_finder': False,
                'save_images': False,
                }
            )

        return settings

    def xlsxwriter_check(self) -> bool:
        """
        XlsxWriter Check

        Check if xlsxWriter is installed by attempting to import it.

        Returns:
        --------
            bool:
                True if xlsxWriter is installed, False if not.
        """

        try:
            import xlsxwriter
            pyflame.print('XlsxWriter Successfully Imported')
            return True
        except:
            pyflame.print('XlsxWriter Not Found, Installing...')

            return False

    def install_xlsxwriter(self) -> None:
        """
        Install XlsxWriter
        ==================

        Install xlsxWriter python package.
        """

        password_window = PyFlamePasswordWindow(
            text='System password is required to install xlsxwriter python package.',
            parent=None,
            )
        #password_window.show()
        system_password = password_window.password

        if system_password:
            python_install_dir = pyflame.get_flame_python_packages_path()
            print('python Install Directory:', python_install_dir)

            # Untar xlsxwriter
            xlsxwriter_tar = os.path.join(SCRIPT_PATH, 'assets/xlsxwriter/xlsxwriter-3.0.3.tgz')

            pyflame.untar(
                tar_file_path=xlsxwriter_tar,
                untar_path=python_install_dir,
                sudo_password=system_password,
                )

            install_dir = os.path.join(python_install_dir, 'xlsxwriter')

            if os.path.isdir(install_dir):
                files = os.listdir(install_dir)
                if files:
                    PyFlameMessageWindow(
                        message='Python xlsxWriter module installed.',
                        title=f'{SCRIPT_NAME}: Operation Complete',
                        parent=None,
                        )
                    flame.execute_shortcut('Rescan Python Hooks')
                    self.main_window()
            else:
                PyFlameMessageWindow(
                    message='Python xlsxWriter module install failed.',
                    parent=None,
                    )

    #-------------------------------------
    # [Windows]
    #-------------------------------------

    def main_window(self):

        def export_path_browse():
            export_path = pyflame.file_browser(
                path=self.export_path_entry.text,
                title='Select Export Path',
                select_directory=True,
                window_to_hide=[self.window],
            )
            if export_path:
                self.export_path_entry.text = export_path

        def save_config():
            # Persist only the export path
            self.settings.export_path = self.export_path_entry.text

            # Path checks
            if not os.path.isdir(self.settings.export_path):
                PyFlameMessageWindow(
                    message='Export path not found - Select new path.',
                    message_type=MessageType.ERROR,
                    parent=self.window,
                )
                return
            if not os.access(self.settings.export_path, os.W_OK):
                PyFlameMessageWindow(
                    message='Unable to export to selected path - Select new path.',
                    message_type=MessageType.ERROR,
                    parent=self.window,
                )
                return

            # If you still want to persist a config file, do it here (path only)
            self.settings.save_config(config_values={
                'export_path': self.export_path_entry.text,
            })

            # Hide → run → close
            self.window.hide()
            self.create_shot_sheets()
            self.window.close()

        def close_window():
            self.window.close()

        # ---- UI (minimal) ----
        self.window = PyFlameWindow(
            title=f'{SCRIPT_NAME} <small>{SCRIPT_VERSION}</small>',
            return_pressed=save_config,
            escape_pressed=close_window,
            grid_layout_columns=5,
            grid_layout_rows=3,  # smaller grid now
            parent=None,
        )

        self.export_path_label = PyFlameLabel(text='Export Path')
        self.export_path_entry = PyFlameEntry(text=self.settings.export_path)

        self.export_path_browse_button = PyFlameButton(text='Browse', connect=export_path_browse)
        self.create_button = PyFlameButton(text='Create', connect=save_config, color=Color.BLUE)
        self.cancel_button = PyFlameButton(text='Cancel', connect=self.window.close)

        # Layout
        self.window.grid_layout.addWidget(self.export_path_label, 0, 0)
        self.window.grid_layout.addWidget(self.export_path_entry, 0, 1, 1, 3)
        self.window.grid_layout.addWidget(self.export_path_browse_button, 0, 4)

        self.window.grid_layout.addWidget(self.cancel_button, 2, 3)
        self.window.grid_layout.addWidget(self.create_button, 2, 4)

        self.export_path_entry.set_focus()

    #-------------------------------------

    def create_shot_sheets(self):
        """
        Create Shot Sheets
        ==================

        Create shot sheets from selected sequences. Export to xlsx format. Open in Finder if selected.
        """
        import xlsxwriter

        def show_error(message):
            PyFlameMessageWindow(
                message=message,
                title=f'{SCRIPT_NAME}: Export Error',
                parent=self.window,
            )

        def safe_log(msg):
            try:
                pyflame.print(msg)
            except Exception:
                print(msg)

        def edit_xlsx_file(xlsx_path):
            """
            Edit XLSX File
            ==============

            Edit the xlsx file to add image links. Links the images to cell in the xlsx file.

            This only works in Excel. It does not work in Google Sheets or Numbers.

            Args
            ----
                xlsx_path (str): Path to the xlsx file to edit.
            """



            def unzip_xlsx(xlsx_path, extract_dir):
                with zipfile.ZipFile(xlsx_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_dir)

            def zip_dir(source_dir, zip_path):
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root, dirs, files in os.walk(source_dir):
                        for file in files:
                            full_path = os.path.join(root, file)
                            arcname = os.path.relpath(full_path, source_dir)
                            zipf.write(full_path, arcname)

            def modify_drawing_xml(drawing_xml_path):
                ns = {
                    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                }
                ET.register_namespace('', ns['xdr'])

                tree = ET.parse(drawing_xml_path)
                root = tree.getroot()

                for anchor_tag in ['{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}twoCellAnchor',
                                '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}oneCellAnchor']:
                    for anchor in root.findall(anchor_tag):
                        anchor.set('editAs', 'twoCell')  # Options: 'absolute', 'oneCell', 'twoCell'

                tree.write(drawing_xml_path, encoding='utf-8', xml_declaration=True)

            def patch_xlsx_images(xlsx_path, output_path):
                temp_dir = os.path.join(os.path.dirname(xlsx_path), 'temp_xlsx')

                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                os.makedirs(temp_dir)

                # Step 1: Unzip the .xlsx file
                unzip_xlsx(xlsx_path, temp_dir)

                # Step 2: Modify drawing XML files
                drawings_dir = os.path.join(temp_dir, 'xl', 'drawings')
                if os.path.isdir(drawings_dir):
                    for file in os.listdir(drawings_dir):
                        if file.startswith('drawing') and file.endswith('.xml'):
                            full_path = os.path.join(drawings_dir, file)
                            modify_drawing_xml(full_path)

                # Step 3: Zip folder back to .xlsx
                temp_zip = output_path + '.zip'
                zip_dir(temp_dir, temp_zip)

                # Step 4: Rename .zip back to .xlsx
                if os.path.exists(output_path):
                    os.remove(output_path)
                os.rename(temp_zip, output_path)

                # Cleanup
                shutil.rmtree(temp_dir)

            patch_xlsx_images(xlsx_path, xlsx_path)

        def save_images():
            """
            Save Images
            ===========

            Save images to export path if selected.
            """

            # Copy images to export path
            if self.settings.save_images:
                image_path = os.path.join(self.export_path_entry.text, f'{seq_name}_images')
                if not os.path.exists(image_path):
                    os.makedirs(image_path)
                for image in os.listdir(self.temp_image_path):
                    if image.endswith('.jpg'):
                        shutil.copy(os.path.join(self.temp_image_path, image), os.path.join(image_path, image))


        # Sort selected sequences by name
        try:
            sequence_names = [str(seq.name)[1:-1] for seq in self.selection]
            sequence_names.sort()
        except Exception as e:
            show_error(f'Failed to sort sequences: {e}')
            safe_log(f'Error sorting sequences: {e}')
            return

        sorted_sequences = []
        try:
            for name in sequence_names:
                for seq in self.selection:
                    if str(seq.name)[1:-1] == name:
                        sorted_sequences.append(seq)
        except Exception as e:
            show_error(f'Failed to build sorted sequence list: {e}')
            safe_log(f'Error building sorted sequence list: {e}')
            return

        # Create one workbook per sequence
        for sequence in sorted_sequences:
            seq_name = str(sequence.name)[1:-1]
            xlsx_path = os.path.join(self.export_path_entry.text, f'{seq_name}.xlsx')
            safe_log(f'Creating workbook at {xlsx_path}')

            try:
                workbook = xlsxwriter.Workbook(xlsx_path)
                self.get_shots(sequence)
                self.create_sequence_worksheet(workbook, seq_name)
                workbook.close()
                edit_xlsx_file(xlsx_path)
            except Exception as e:
                show_error(f'Error creating workbook for {seq_name}: {e}')
                safe_log(f'Error creating workbook for {seq_name}: {e}')
                continue


        # Delete temp directory
        try:
            shutil.rmtree(self.temp_path)
        except Exception as e:
            safe_log(f'Error deleting temp directory: {e}')

        # Close window
        try:
            self.window.close()
        except Exception as e:
            safe_log(f'Error closing window: {e}')

        # Show message window
        try:
            PyFlameMessageWindow(
                message=f'Shot Sheet(s) Exported:\n\n{self.export_path_entry.text}',
                title=f'{SCRIPT_NAME}: Export Complete',
                parent=self.window,
            )
        except Exception as e:
            safe_log(f'Error showing export complete message: {e}')

        if self.settings.reveal_in_finder:
            try:
                pyflame.open_in_finder(
                    path=self.export_path_entry.text,
                )
                safe_log('Finder opened')
            except Exception as e:
                safe_log(f'Error opening Finder: {e}')

        safe_log('Done.')

    def get_shots(self, sequence):
        """
        Get Shots
        =========

        Export thumbnails and get shot info for all shots in selected sequence
        """

        def thumbnail_res():
            seq_height = sequence.height
            seq_width = sequence.width
            seq_ratio = float(seq_width) / float(seq_height)

            self.thumb_nail_height = 100
            self.thumb_nail_width = int(self.thumb_nail_height * seq_ratio)
            self.x_offset = 30

            self.row_height = self.thumb_nail_height + (self.thumb_nail_height * .2)
            self.column_width = (self.thumb_nail_width + (self.x_offset * 2)) / 7.83
            self.y_offset = ((self.row_height * 1.333) - self.thumb_nail_height) / 2

        def export_thumbnail(self, sequence, segment, shot_name):

            # Create list of existing exported thumbnails to check for extra frames
            temp_image_path_files = os.listdir(self.temp_image_path)

            # Modify export preset with selected resolution
            edit_preset = open(self.temp_export_preset, 'r')
            contents = edit_preset.readlines()
            edit_preset.close()

            contents[8] = f'  <namePattern>{shot_name}</namePattern>\n'
            contents[15] = f'   <width>{self.thumb_nail_width}</width>\n'
            contents[16] = f'   <height>{self.thumb_nail_height}</height>\n'
            contents[26] = '  <framePadding>0</framePadding>\n'

            edit_preset = open(self.temp_export_preset, 'w')
            contents = ''.join(contents)
            edit_preset.write(contents)
            edit_preset.close()

            # Mark in and out in sequence for segment frame to export
            sequence.in_mark = segment.record_in
            sequence.out_mark = segment.record_in + 1

            # Export thumbnail
            self.exporter.export(sequence, self.temp_export_preset, self.temp_image_path)

            # Clear sequence in and out marks
            sequence.in_mark = None
            sequence.out_mark = None

            # Fix for extra frames being exported when Inclusive Out Marks is selected in Flame Timeline Prefs.
            # Check if more than one thumbnail was exported. If so, delete the second frame and rename the first frame to remove the frame number.
            updated_temp_image_path_files = os.listdir(self.temp_image_path)

            if len(updated_temp_image_path_files) - len(temp_image_path_files) == 2:
                new_images = [image for image in updated_temp_image_path_files if image not in temp_image_path_files]

                # Delete first image in new_images list
                os.remove(os.path.join(self.temp_image_path, new_images[0]))

                # Rename second image in new_images list to remove frame number and add to temp_image_path_files list.
                new_image = new_images[1]

                if new_image.count('.') > 1:
                    os.rename(os.path.join(self.temp_image_path, new_image), os.path.join(self.temp_image_path, new_image.split('.')[0] + '.jpg'))
                temp_image_path_files.append(new_image)

        # Create temp directory to store thumbnails
        self.temp_image_path = os.path.join(self.temp_path, str(sequence.name)[1:-1])
        if os.path.exists(self.temp_image_path):
            shutil.rmtree(self.temp_image_path)
        os.makedirs(self.temp_image_path)

        # Set thumbnail size
        thumbnail_res()

        self.shot_dict = OrderedDict()

        # Create dictionary for all shots containing clip info
        for segment in sequence.versions[0].tracks[0].segments:
            if segment.type == 'Video Segment':
                shot_name = str(segment.shot_name)[1:-1]
                if not shot_name:
                    shot_name = str(segment.name)[1:-1]

                # Export thumbnail for segment
                export_thumbnail(self, sequence, segment, shot_name)

                self.clip_info_list = []

                self.clip_info_list.append(f'Shot Name: {shot_name}')
                self.clip_info_list.append(f'Source Name: {str(segment.source_name)}')
                self.clip_info_list.append(f'Source Path: {segment.file_path}')
                self.clip_info_list.append(f'Source TC: {str(segment.source_in)[1:-1]} - {str(segment.source_out)[1:-1]}')
                self.clip_info_list.append(f'Source TC In: {str(segment.source_in)[1:-1]}')
                self.clip_info_list.append(f'Source TC Out: {str(segment.source_out)[1:-1]}')
                self.clip_info_list.append(f'Record TC: {str(segment.record_in)[1:-1]} - {str(segment.record_out)[1:-1]}')
                self.clip_info_list.append(f'Record TC In: {str(segment.record_in)[1:-1]}')
                self.clip_info_list.append(f'Record TC Out: {str(segment.record_out)[1:-1]}')
                self.clip_info_list.append(f'Shot Frame In: {str(flame.PyTime(segment.record_in.relative_frame))}')
                self.clip_info_list.append(f'Shot Frame Out: {str(flame.PyTime(segment.record_out.relative_frame))}')
                self.clip_info_list.append(f'Shot Length: {str(segment.record_duration)[1:-1]} - {str(flame.PyTime(segment.record_duration.frame))} Frames')

                # Source length can be inifinite for things like solid colors, check for this.
                if segment.source_duration != 'infinite':
                    self.clip_info_list.append(f'Source Length: {str(segment.source_duration)[1:-1]} - {str(flame.PyTime(segment.source_duration.frame))} Frames')
                else:
                    self.clip_info_list.append('Source Length: Infinite')

                self.clip_info_list.append(f'Comment: {str(segment.comment)[1:-1]}')

                print(f'{shot_name} Clip Info:\n')
                for info in self.clip_info_list:
                    print(f'    {info}')
                print('\n')

                self.shot_dict.update({shot_name : self.clip_info_list})

    def create_sequence_worksheet(self, workbook, seq_name):

        def add_clip_info_column(clip_info) -> None:
            """
            This will add clip info as additional rows below each shot's 4-row block.
            """

            if not clip_info:
                return

            # Set wider column for clip info
            worksheet.set_column('B:B', 50)

            # Start from first shot row and increment by 4 for each shot plus any additional info rows
            current_row = 1
            row_offset = 0

            for shot in self.shot_dict:
                clip_info_text = []

                # Always collect clip info
                clip_info_text.append(str(self.shot_dict[shot][1]))  # Source Name
                clip_info_text.append(str(self.shot_dict[shot][3]))  # Source TC
                clip_info_text.append(str(self.shot_dict[shot][6]))  # Record TC
                clip_info_text.append(str(self.shot_dict[shot][11])) # Shot Length

                if clip_info_text:
                    # Calculate actual row with offset
                    actual_row = current_row + row_offset
                    
                    # Write clip info in additional rows below the 4-row block
                    clip_info = '\n'.join(clip_info_text)
                    worksheet.merge_range(actual_row + 4, 1, actual_row + 5, 7, clip_info, task_format)
                    
                    # Set appropriate height for info rows
                    worksheet.set_row(actual_row + 4, 30)
                    worksheet.set_row(actual_row + 5, 30)
                    
                    # Increment offset for additional rows
                    row_offset += 2

                current_row += 4  # Move to next shot's block

        def add_token_clip_info() -> None:
            """
            Add the clip info to appropriate fields in the shot info column based on column name tokens.
            """

            # List of clip info tokens to check against column names
            info_tokens = [
                'Shot Name',
                'Source Name',
                'Source Path',
                'Source TC',
                'Source TC In',
                'Source TC Out',
                'Record TC',
                'Record TC In',
                'Record TC Out',
                'Shot Frame In',
                'Shot Frame Out',
                'Shot Length',
                'Source Length',
                'Comment'
            ]

            # Keep track of extra rows added for clip info
            row_offset = 0
            base_row = 1

            for shot_name in self.shot_dict:
                # Calculate actual row with offset
                current_row = base_row + row_offset
                
                # Add relevant clip info under each shot's task/status/artist section
                additional_info = []
                
                for name in info_tokens:
                    if name in column_names and name != 'Shot Name':  # Skip shot name as it's already shown
                        info = self.shot_dict[shot_name][info_tokens.index(name)].split(': ', 1)[1]
                        additional_info.append(f"{name}: {info}")

                if additional_info:
                    # Write the additional info in merged cells below the artist field
                    info_text = '\n'.join(additional_info)
                    worksheet.merge_range(current_row + 4, 1, current_row + 5, 7, info_text, task_format)
                    worksheet.set_row(current_row + 4, 30)
                    worksheet.set_row(current_row + 5, 30)
                    row_offset += 2

                base_row += 4  # Move to next shot's base position

        # Define department list once for reuse
        departments = ['Tracking', 'Roto', 'Paint', 'DMP', 'Comp', 'CG']

        # Create worksheet
        worksheet = workbook.add_worksheet(seq_name)
        worksheet.set_column('A:A', self.column_width * 1.3)  # Thumbnail column - 30% wider
        worksheet.set_column('B:B', 20)  # Shot info column - narrower
        worksheet.set_column('C:H', 25)  # Department columns

        # Define cell formats with new color scheme
        # Separate format for title row (dark gray)
        title_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': '#2C2C2C',  # Dark gray for title
            'font_color': 'white',
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_size': 14  # Slightly larger for title
        })

        # Format for column headers (darker purple)
        header_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': '#2C2C2C',  # Darker gray
            'font_color': 'white',
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })

        shot_name_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': "#505050",  # Dark gray
            'font_color': 'white',
            'bold': True,
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })

        task_header_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': '#2C2C2C',  # Darker gray
            'font_color': 'white',
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': True
        })

        dept_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': '#2C2C2C',  # Darker purple
            'font_color': 'white',
            'bold': True,
            'align': 'center',
            'valign': 'vcenter'
        })

        label_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': "#3A3A3A",  # Light purple
            'font_color': 'white',
            'align': 'left',
            'valign': 'vcenter',
            'text_wrap': True
        })

        empty_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': '#404040',  # Lighter gray
            'font_color': 'white',
            'valign': 'vcenter',
            'text_wrap': True
        })

        shot_name_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bold': True,
            'valign': 'vcenter',
            'text_wrap': True
        })

        task_format = workbook.add_format({
            'font_name': 'Helvetica',
            'valign': 'vcenter',
            'text_wrap': True
        })

        # Add basic column headers
        worksheet.write(0, 0, 'Thumbnail', header_format)  # A1
        worksheet.write(0, 1, 'Shot Info', header_format)  # B1
        worksheet.write(0, 2, 'Departments and Tasks', header_format)  # Merge remaining header cells
        worksheet.merge_range(0, 2, 0, 7, 'Departments and Tasks', header_format)

        # Process each shot starting from row 1
        current_row = 1

        # Process each shot
        for shot_name in self.shot_dict:
            # Set row heights for the 5-row block
            for i in range(5):
                worksheet.set_row(current_row + i, self.row_height / 5)

            # Thumbnail cell format matching shot name background
            thumbnail_format = workbook.add_format({
                'font_name': 'Helvetica',
                'bg_color': '#2C2C2C',  # Same dark gray as shot name
                'font_color': 'white',
                'valign': 'vcenter'
            })
            
            # Merge cells vertically for the thumbnail column
            worksheet.merge_range(current_row, 0, current_row + 4, 0, '', thumbnail_format)
            
            # Insert thumbnail with larger scale
            image_path = os.path.join(self.temp_image_path, shot_name) + '.jpg'
            # Re-apply the format after image insertion to ensure background color
            worksheet.merge_range(current_row, 0, current_row + 4, 0, '', thumbnail_format)
            worksheet.insert_image(current_row, 0, image_path, {
                'x_offset': self.x_offset,
                'y_offset': self.y_offset,
                'x_scale': 1.3,
                'y_scale': 1.3,
                # 'object_position': 3  # 3 = Move and size with cells, no background
            })

            # Set the top row height to double
            worksheet.set_row(current_row, self.row_height / 2)  # Double height for shot name row
            
            # Enhanced shot name format with larger font and dark gray background
            shot_name_format = workbook.add_format({
                'font_name': 'Helvetica',
                'bold': True,
                'font_size': 18,  # Increased to 18
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True,
                'bg_color': '#2C2C2C',  # Darker gray to match header style
                'font_color': 'white',   # White text
                'border': 1,             # Add subtle border
                'border_color': '#404040' # Slightly lighter border color
            })
            
            # Define the task format that will be used for both metadata and department labels
            dept_task_format = workbook.add_format({
                'font_name': 'Helvetica',
                'bg_color': '#2C2C2C',  # Darker purple
                'font_color': 'white',   # White text
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
            })

            # Write shot name merged across both name and buffer cell
            worksheet.merge_range(current_row, 1, current_row + 1, 2, shot_name, shot_name_format)  # Shot name merged across two columns

            # Write numbers 1-5 in the top row cells after the buffer
            for i in range(5):
                worksheet.write(current_row, i + 3, str(i + 1), shot_name_format)

            # Write metadata labels directly in top row cells
            # Define formats for metadata
            metadata_header_format = workbook.add_format({
                'font_name': 'Helvetica',
                'bg_color': '#2C2C2C',  # Dark gray background
                'font_color': 'white',   # White text
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True,
                'font_size': 10  # Smaller font for headers
            })

            # Format for source name (with text clipping)
            source_name_format = workbook.add_format({
                'font_name': 'Helvetica',
                'bg_color': '#404040',  # Lighter gray for values
                'font_color': 'white',
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': False,  # Disable text wrapping
                'font_size': 9  # Even smaller font for values
            })

            # Format for other metadata values (with text wrapping)
            metadata_value_format = workbook.add_format({
                'font_name': 'Helvetica',
                'bg_color': '#404040',  # Lighter gray for values
                'font_color': 'white',
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True,
                'font_size': 9  # Even smaller font for values
            })

            # Write metadata labels in top row with adjusted height
            metadata_labels = ['Source Name', 'Source TC I/O', 'Seq TC I/O', 'Length frames', '']
            worksheet.set_row(current_row, 30)  # Double the height for header row
            for i, label in enumerate(metadata_labels):
                worksheet.write(current_row, i + 3, label, metadata_header_format)

            # Set increased height for metadata values row to accommodate two lines of timecode
            worksheet.set_row(current_row + 1, self.row_height * 0.35)  # Increased height for two lines

            # Get the source name from the shot_dict (index 1 has the source name info)
            source_info = str(self.shot_dict[shot_name][1])
            source_name = source_info.split(': ', 1)[1] if ': ' in source_info else source_info

            # Get source timecode in and out points (indices 4 and 5 have the individual TC points)
            tc_in = str(self.shot_dict[shot_name][4]).split(': ', 1)[1]  # Source TC In
            tc_out = str(self.shot_dict[shot_name][5]).split(': ', 1)[1]  # Source TC Out
            source_tc = f"{tc_in}\n{tc_out}"  # Format with line break
            
            # Get sequence (record) timecode in and out points (indices 7 and 8 have the record TC points)
            seq_tc_in = str(self.shot_dict[shot_name][7]).split(': ', 1)[1]  # Record TC In
            seq_tc_out = str(self.shot_dict[shot_name][8]).split(': ', 1)[1]  # Record TC Out
            seq_tc = f"{seq_tc_in}\n{seq_tc_out}"  # Format with line break
            
            # Get frame count from Shot Length (index 11)
            length_info = str(self.shot_dict[shot_name][11])  # Format: "Shot Length: duration - X Frames"
            frame_count = length_info.split(' - ')[1].split(' ')[0]  # Extract just the number
            
            # Set specific width for source name column
            worksheet.set_column(3, 3, 20)  # Set column D (index 3) to width 20

            # Write the metadata values
            worksheet.write(current_row + 1, 3, source_name, source_name_format)  # Using clip format
            worksheet.write(current_row + 1, 4, source_tc, metadata_value_format)
            worksheet.write(current_row + 1, 5, seq_tc, metadata_value_format)
            worksheet.write(current_row + 1, 6, frame_count, metadata_value_format)
            worksheet.write(current_row + 1, 7, '', metadata_value_format)  # Empty cell with metadata formatting
            for col in range(3, 7):
                worksheet.write(current_row + 1, col, '', shot_name_format)
            
            # Write remaining info below the merged shot name
            worksheet.write(current_row + 2, 1, 'Task:', task_header_format)    # Task (darker purple)
            worksheet.write(current_row + 3, 1, 'Status:', label_format)        # Status (light purple)
            worksheet.write(current_row + 4, 1, 'Artist:', label_format)        # Artist (light purple)

            # Add department labels for task row (Bold and centered)
            dept_format = workbook.add_format({
                'font_name': 'Helvetica',
                'bold': True,
                'align': 'center',
                'valign': 'vcenter'
            })

            # Write department names in the task row (with darker purple background)
            # Write department names in task row using the pre-defined departments list
            for col, dept in enumerate(departments, start=2):  # Start from column C (index 2)
                worksheet.write(current_row + 2, col, dept, dept_task_format)  # Write in task row with dark purple

            # Keep cells empty for Status and Artist rows, but don't merge
            empty_format = workbook.add_format({
                'font_name': 'Helvetica',
                'valign': 'vcenter'
            })
            
            # Define status formats with colors
            status_formats = {
                'Not Started': workbook.add_format({
                    'font_name': 'Helvetica',
                    'valign': 'vcenter',
                    'bg_color': '#404040',  # Light gray for not started
                    'font_color': 'white'
                }),
                'In Progress': workbook.add_format({
                    'font_name': 'Helvetica',
                    'valign': 'vcenter',
                    'bg_color': '#1E90FF',  # Brighter blue
                    'font_color': 'white'
                }),
                'Review': workbook.add_format({
                    'font_name': 'Helvetica',
                    'valign': 'vcenter',
                    'bg_color': '#FF69B4',  # Brighter pink
                    'font_color': 'white'
                }),
                'Internally Approved': workbook.add_format({
                    'font_name': 'Helvetica',
                    'valign': 'vcenter',
                    'bg_color': '#3CB371',  # Medium sea green
                    'font_color': 'white'
                }),
                'Client Approved': workbook.add_format({
                    'font_name': 'Helvetica',
                    'valign': 'vcenter',
                    'bg_color': '#228B22',  # Forest green
                    'font_color': 'white'
                }),
                'On Hold': workbook.add_format({
                    'font_name': 'Helvetica',
                    'valign': 'vcenter',
                    'bg_color': '#DC143C',  # Crimson red
                    'font_color': 'white'
                })
            }            # Status options for dropdown
            status_options = list(status_formats.keys())
            
            # Add dropdowns and format cells
            for col in range(2, 8):  # Columns C through H
                # Empty cells in top rows should use shot name format for consistency
                # Keep the top row for metadata labels
                if current_row + 1 == current_row:  # Only clear second row
                    worksheet.write(current_row + 1, col, '', shot_name_format)
                
                # Add dropdown validation to status cell
                worksheet.data_validation(current_row + 3, col, current_row + 3, col, {
                    'validate': 'list',
                    'source': status_options
                })
                
                # Write default status and empty artist cell with lighter gray background
                status_cell_format = workbook.add_format({
                    'font_name': 'Helvetica',
                    'bg_color': '#404040',  # Lighter gray
                    'font_color': 'white',
                    'valign': 'vcenter',
                    'text_wrap': True
                })
                worksheet.write(current_row + 3, col, 'Not Started', status_cell_format)  # Status row with default value

                # Very light gray format for artist cells
                artist_format = workbook.add_format({
                    'font_name': 'Helvetica',
                    'bg_color': '#FFFFFF',  # Pure white
                    'valign': 'vcenter',
                    'text_wrap': True
                })
                # Write artist cell with white background
                worksheet.write(current_row + 4, col, '', artist_format)

            # Add conditional formatting for status cells
            for status, format_obj in status_formats.items():
                if status != 'Not Started':  # Skip default format
                    worksheet.conditional_format(current_row + 3, 2, current_row + 3, 7, {
                        'type': 'cell',
                        'criteria': 'equal to',
                        'value': f'"{status}"',
                        'format': format_obj
                    })
            
            # Force write the metadata values after all other operations
            source_info = str(self.shot_dict[shot_name][1])
            source_name = source_info.split(': ', 1)[1] if ': ' in source_info else source_info
            
            tc_in = str(self.shot_dict[shot_name][4]).split(': ', 1)[1]  # Source TC In
            tc_out = str(self.shot_dict[shot_name][5]).split(': ', 1)[1]  # Source TC Out
            source_tc = f"{tc_in}\n{tc_out}"  # Format with line break
            
            seq_tc_in = str(self.shot_dict[shot_name][7]).split(': ', 1)[1]  # Record TC In
            seq_tc_out = str(self.shot_dict[shot_name][8]).split(': ', 1)[1]  # Record TC Out
            seq_tc = f"{seq_tc_in}\n{seq_tc_out}"  # Format with line break
            
            # Get frame count
            length_info = str(self.shot_dict[shot_name][11])  # Format: "Shot Length: duration - X Frames"
            frame_count = length_info.split(' - ')[1].split(' ')[0]  # Extract just the number
            
            worksheet.write(current_row + 1, 3, source_name, source_name_format)  # Using clip format
            worksheet.write(current_row + 1, 4, source_tc, metadata_value_format)
            worksheet.write(current_row + 1, 5, seq_tc, metadata_value_format)
            worksheet.write(current_row + 1, 6, frame_count, metadata_value_format)

            # Add a black divider row after each shot
            divider_format = workbook.add_format({
                'bg_color': '#000000',  # Pure black
                'font_color': '#000000'  # Hide any potential content
            })
            
            # Set divider row height to be very small
            worksheet.set_row(current_row + 5, 15)  #second number is pixel height
            
            # Add black divider across all columns
            for col in range(8):  # Columns A through H
                worksheet.write(current_row + 5, col, '', divider_format)
            
            # Move to next group (5 rows for content + 1 for divider)
            current_row += 6

        # Add sequence name to header area using title_format
        worksheet.merge_range(0, 0, 0, 7, seq_name, title_format)

#-------------------------------------
# [Scopes]
#-------------------------------------

def scope_sequence(selection):

    for item in selection:
        if isinstance(item, (flame.PySequence, flame.PyClip)):
            return True
    return False

#-------------------------------------
# [Flame Menus]
#-------------------------------------

def get_media_panel_custom_ui_actions():

    return [
        {
           'hierarchy': [],
           'actions': [
               {
                    'name': 'Shot Sheet Maker',
                    'order': 1,
                    'separator': 'below',
                    'isVisible': scope_sequence,
                    'execute': ShotSheetMaker,
                    'minimumVersion': '2025'
               }
           ]
        }
    ]
