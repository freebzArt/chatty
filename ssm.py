"""
Description:
    Create shot sheets from selected sequence clips that can be loaded into Google Sheets.

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
import shutil
import xml.etree.ElementTree as ET
from collections import OrderedDict
from datetime import datetime
from pathlib import Path 
import platform
import flame
from lib.pyflame_lib_shot_sheet_maker import *
from xlsxwriter.utility import xl_rowcol_to_cell

#-------------------------------------
# [Constants]
#-------------------------------------

SCRIPT_NAME = 'Shot Sheet Maker - Joint'
SCRIPT_VERSION = 'v1.0.0'
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

        # Copy jpeg export preset to temp directory to modify for each export
        self.temp_export_preset = os.path.join(self.temp_path, 'Temp_Export_Preset.xml')
        shutil.copy(jpg_preset_path, self.temp_export_preset)

        # Make sure xlsxwriter is installed, if not, install it. Otherwise, open window.
        xlsxwriter_installed = self.xlsxwriter_check()

        if xlsxwriter_installed:
            return self.main_window()
        return self.install_xlsxwriter()

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

        def run_export():
            # Hide → run → close, unless error then user can fix or exit
            self.window.hide()
            ok = self.create_shot_sheets()
            if ok:
                self.window.close()
            else:
                try: 
                    self.window.show()
                except Exception:
                    pass

        self.window = PyFlameWindow(
            title=f'{SCRIPT_NAME} <small>{SCRIPT_VERSION}</small>',
            return_pressed=run_export,
            grid_layout_columns=5,
            grid_layout_rows=3,  
            parent=None,
        )

        self.export_path_label = PyFlameLabel(text='Export Path')
        default_root = self._suggest_export_root()
        self.export_path_entry = PyFlameEntry(text=str(default_root))

        self.export_path_browse_button = PyFlameButton(text='Browse', connect=export_path_browse)
        self.create_button = PyFlameButton(text='Create', connect=run_export, color=Color.BLUE)
        self.cancel_button = PyFlameButton(text='Cancel', connect=self.window.close)

        # Layout
        self.window.grid_layout.addWidget(self.export_path_label, 0, 0)
        self.window.grid_layout.addWidget(self.export_path_entry, 0, 1, 1, 3)
        self.window.grid_layout.addWidget(self.export_path_browse_button, 0, 4)

        self.window.grid_layout.addWidget(self.cancel_button, 2, 3)
        self.window.grid_layout.addWidget(self.create_button, 2, 4)

        self.export_path_entry.set_focus()

    def _add_roster_sheet(self, workbook, initial_artists=None):
        if initial_artists is None:
            initial_artists = ["Rajesh", "Track"]
        
        roster = workbook.add_worksheet("Roster")

        #Simple format
        header_format = workbook.add_format({
            'font_name': 'Helvetica', 
            'bold': True, 
            'bg_color': '#2C2C2C',
            'font_color': 'white', 
            'align': 'center', 
            'valign': 'vcenter',  
        })
        note_format = workbook.add_format({
            'font_name': 'Helvetica', 
            'align': 'left'
            })

        roster.set_column('A:A', 28)
        roster.write(0, 0, 'Artists (edit below)', header_format)
        roster.write(1, 0, 'Not Assigned', note_format) # Default choice

        # If initial artists is provided
        row = 2
        for name in initial_artists:
            roster.write(row, 0, name, note_format)
            row += 1

        workbook.define_name('Artists', f'=Roster!$A$2:$A$200')

    def create_shot_sheets(self):
        """
        Create Shot Sheets
        ==================

        Create shot sheets from selected sequences. Export to xlsx format. 
        """
        # Imported here so script does not fail if xlsxwriter is not installed and is needs installed by script
        import xlsxwriter   

        #### Create error dialogue methods
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
        ####

        # Sort selected sequences by name
        try:
            sorted_sequences = sorted(
                self.selection,
                key=lambda s: str(s.name)[1:-1]
            )
        except Exception as e:
            show_error(f'Failed to sort sequences: {e}')
            safe_log(f'Error sorting sequences: {e}')
            return

        # Assign contents of dialogue box path to chose_root
        chosen_root = Path(self.export_path_entry.text)

        # Try to create dir; if we can’t raise error
        try:
            chosen_root.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            PyFlameMessageWindow(
                message=(f"Unable to create xport folder :\n{chosen_root}\n\n{e}\n"
                         "Please choose a different location."),
                title=f"{SCRIPT_NAME}: Export Error",
                message_type=MessageType.ERROR,
                parent=self.window, 
            )
            return 
        
        # Try to create filer; if we can’t raise error
        testfile = chosen_root / ".write_test"
        try:
            with open(testfile, "w") as _f:
                _f.write("")
        except Exception as e:
            PyFlameMessageWindow(
                message=(f"No write permission in:\n{chosen_root}\n\n{e}\n"
                         "Please choose a different location."),
                title=f"{SCRIPT_NAME}: Export Error",
                message_type=MessageType.ERROR,
                parent=self.window,
            )
            return
        finally:
            try:
                testfile.unlink()
            except Exception:
                pass
        
        # If both tests pass, safe to use chosen_root
        export_root = chosen_root
        
        # Create one workbook per sequence
        for sequence in sorted_sequences:
            seq_name = str(sequence.name)[1:-1] # trims quotes
            xlsx_path = os.path.join(str(export_root), f'{seq_name}.xlsx')
            safe_log(f'Creating workbook at {xlsx_path}')

            try:
                workbook = xlsxwriter.Workbook(xlsx_path)
                self.get_shots(sequence)
                self.create_sequence_worksheet(workbook, seq_name)
                self._add_roster_sheet(workbook)
                workbook.close()
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
            def _open_export_location():
                try:
                    pyflame.open_in_finder(path=str(export_root))
                finally:
                    export_done_window.close()

            def _close_export_dialog():
                export_done_window.close()

            export_done_window = PyFlameWindow(
                title=f'{SCRIPT_NAME}: Export Complete',
                return_pressed=_close_export_dialog,   
                escape_pressed=_close_export_dialog,   
                grid_layout_columns=5,
                grid_layout_rows=3,
                parent=None,  
            )

            msg = PyFlameLabel(
                text=f'Shot sheet(s) exported to:\n{export_root}'
            )

            open_btn = PyFlameButton(
                text='Open Folder',
                connect=_open_export_location,
                color=Color.BLUE
            )

            ok_btn = PyFlameButton(
                text='OK',
                connect=_close_export_dialog
            )

            export_done_window.grid_layout.addWidget(msg,    0, 0, 1, 5)
            export_done_window.grid_layout.addWidget(ok_btn, 2, 3)
            export_done_window.grid_layout.addWidget(open_btn, 2, 4)

            ok_btn.set_focus()

        except Exception as e:
            safe_log(f'Error showing export complete message: {e}')

        safe_log('Done.')

    def get_shots(self, sequence):
        """
        Get Shots
        =========

        Export thumbnails and get shot info for all shots in selected sequence
        """
        # remember which sequence the sheet is for
        self.current_sequence = sequence

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

    def _suggest_export_root(self) -> Path:
        now = datetime.datetime.now()
        date_str = now.strftime("%Y%m%d")
        time_str = now.strftime("%H%M") 

        base = Path("/Volumes/vfx_1") if platform.system() == "Darwin" else Path("/home/jointadmin/mnt/vfx_1")
        job_root = base / self.flame_project_name

        if job_root.is_dir():
            return job_root / "pipeline" / "shot_tracker" / date_str / time_str
        
        desktop = Path.home() / "Desktop"
        if not desktop.exists():
            desktop = Path.home()
        return desktop / f"{self.flame_project_name}_shot_sheets" / date_str / time_str 
    
    def _sequence_meta(self, sequence) -> dict:
        """Collect and format useful sequence metadata from Flame, safely."""
        def _flt(x: object) -> str:
            s = str(x)
            if len(s) >= 2 and (s[0] == s[-1] == "'" or s[0] == s[-1] == '"'):
                return s[1:-1]
            return s

        # Project
        project = _flt(getattr(self, 'flame_project_name', ''))

        # Resolution
        width = getattr(sequence, 'width', None)
        height = getattr(sequence, 'height', None)
        resolution = f"{width} x {height}" if width and height else "Unknown"

        # Frame rate
        fps = (
            getattr(sequence, 'frame_rate', None)
            or getattr(sequence, 'fps', None)
            or getattr(sequence, 'rate', None)
        )
        # Coerce common PyFlame wrappers or numerics to a string
        fps_str = _flt(fps) if fps is not None else "Unknown"

        # Bit depth (varies by Flame version/schema)
        bit_depth = (
            getattr(sequence, 'bit_depth', None)
            or getattr(sequence, 'bitDepth', None)
            or getattr(sequence, 'depth', None)
        )
        bit_depth_str = _flt(bit_depth) if bit_depth is not None else "Unknown"

        # Duration (prefer a PyTime-like; fall back to compute from segments)
        duration = getattr(sequence, 'duration', None)
        if duration is not None:
            duration_str = _flt(duration)
        else:
            # Coarse fallback from segment bounds
            try:
                segs = [s for s in sequence.versions[0].tracks[0].segments if s.type == 'Video Segment']
                if segs:
                    start = segs[0].record_in
                    end = segs[-1].record_out
                    duration_str = f"{_flt(end - start)}"
                else:
                    duration_str = "Unknown"
            except Exception:
                duration_str = "Unknown"

        return {
            "project": project or "Unknown",
            "resolution": resolution,
            "fps": fps_str,
            "bit_depth": bit_depth_str,
            "duration": duration_str,
        }

    def create_sequence_worksheet(self, workbook, seq_name):
        # from xlsxwriter.utility import xl_rowcol_to_cell

        # Define department list for reuse
        departments = ['Tracking', 'Roto', 'Paint', 'DMP', 'Comp', 'CG']

        # Create worksheet
        worksheet = workbook.add_worksheet(seq_name)
        worksheet.set_column('A:A', self.column_width * 1.3)  # Thumbnail column - 30% wider
        # worksheet.set_column('B:B', 20)  # Shot info column - narrower
        worksheet.set_column('B:G', 25)  # Department columns

        #-------------------------------------
        # FORMATS   
        #-------------------------------------

        title_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': '#2C2C2C',  
            'font_color': 'white',
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_size': 14  
        })

        meta_title_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': '#2C2C2C',  
            'font_color': 'white',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_size': 10,
        })
                
        thumbnail_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': '#2C2C2C',  
            'font_color': 'white',
            'valign': 'vcenter'
        })

        shot_name_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bold': True,
            'font_size': 18,  
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'bg_color': '#2C2C2C',  
            'font_color': 'white',   
            'border': 1,             
            'border_color': '#404040' 
        })

        dept_label_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': '#2C2C2C',  
            'font_color': 'white',   
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
        })  

        metadata_label_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': '#2C2C2C',  
            'font_color': 'white',   
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'font_size': 10  
        })

        metadata_value_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': '#404040',  
            'font_color': 'white',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': False,
            'font_size': 9  
        })

        artist_default_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': "#6F6F6F",
            'font_color': 'white',
            'valign': 'vcenter',
            'text_wrap': True,
        })        

        artist_assigned_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': "#A79BAE",   # your pastel for any assigned artist
            'font_color': '#000000',
            'bold': True,
            'valign': 'vcenter',
            'text_wrap': True,
        })

        status_cell_format = workbook.add_format({
            'font_name': 'Helvetica',
            'bg_color': '#404040',  
            'font_color': 'white',
            'valign': 'vcenter',
            'text_wrap': True
        })

        divider_format = workbook.add_format({
            'bg_color': '#000000',  
            'font_color': '#000000'  
        })
        
        # Row indices for layout
        TITLE_ROW = 0
        META_ROW = 1
        FIRST_SHOT_ROW = META_ROW + 2 # leave one row for divider

        #Big Title Row
        worksheet.set_row(TITLE_ROW, 40) 
        worksheet.merge_range(TITLE_ROW, 0, TITLE_ROW, 6, seq_name, title_format)
        
        # Metadata row
        meta = self._sequence_meta(self.current_sequence)
        meta_text = (
            f"Project: {meta['project']}   •   "
            f"Resolution: {meta['resolution']}   •   "
            f"FPS: {meta['fps']}   •   "
            f"Bit Depth: {meta['bit_depth']}   •   "
            f"Duration: {meta['duration']}"
        )
        worksheet.set_row(META_ROW, 30)
        worksheet.merge_range(META_ROW, 0, META_ROW, 6, meta_text, meta_title_format)

        # Divider row after sequence metadata
        SPACER_ROW = META_ROW + 1
        worksheet.set_row(SPACER_ROW, 15)

        for col in range(7):
            worksheet.write_blank(SPACER_ROW, col, None, divider_format)

        current_row = FIRST_SHOT_ROW

        # Process each shot
        for shot_name in self.shot_dict:
            # Set row heights for the 5-row block
            for i in range(5):
                worksheet.set_row(current_row + i, self.row_height / 5)

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

            # Write shot name merged across both name and buffer cell
            worksheet.merge_range(current_row, 1, current_row + 1, 2, shot_name, shot_name_format)  # Shot name merged across two columns

            # Write metadata labels in top row with adjusted height
            metadata_labels = ['Source Name', 'Source TC I/O', 'Seq TC I/O', 'Length frames']
            worksheet.set_row(current_row, 30)  
            for i, label in enumerate(metadata_labels):
                worksheet.write(current_row, i + 3, label, metadata_label_format)

            # Set increased height for metadata values row to accommodate two lines of timecode
            worksheet.set_row(current_row + 1, self.row_height * 0.35)  

            # Get the source name from the shot_dict (index 1 has the source name info)
            source_info = str(self.shot_dict[shot_name][1])
            source_name = source_info.split(': ', 1)[1] if ': ' in source_info else source_info

            # Get source timecode in and out points (indices 4 and 5 have the individual TC points)
            tc_in = str(self.shot_dict[shot_name][4]).split(': ', 1)[1]  
            tc_out = str(self.shot_dict[shot_name][5]).split(': ', 1)[1]  
            source_tc = f"{tc_in}\n{tc_out}" 
            
            # Get sequence (record) timecode in and out points (indices 7 and 8 have the record TC points)
            seq_tc_in = str(self.shot_dict[shot_name][7]).split(': ', 1)[1]  
            seq_tc_out = str(self.shot_dict[shot_name][8]).split(': ', 1)[1]  
            seq_tc = f"{seq_tc_in}\n{seq_tc_out}"  
            
            # Get frame count from Shot Length (index 11)
            length_info = str(self.shot_dict[shot_name][11])  
            frame_count = length_info.split(' - ')[1].split(' ')[0]  
            
            # Set specific width for source name column
            worksheet.set_column(3, 3, 20)  

            # Write the metadata values
            worksheet.write(current_row + 1, 3, source_name, metadata_value_format) 
            worksheet.write(current_row + 1, 4, source_tc, metadata_value_format)
            worksheet.write(current_row + 1, 5, seq_tc, metadata_value_format)
            worksheet.write(current_row + 1, 6, frame_count, metadata_value_format)
            for col in range(3, 7):
                worksheet.write(current_row + 1, col, '', shot_name_format)

            # Write department names in task row using the pre-defined departments list
            for col, dept in enumerate(departments, start=1): 
                worksheet.write(current_row + 2, col, dept, dept_label_format)  
            
            # Define status formats with colors
            status_formats = {
                'Not Started': workbook.add_format({
                    'font_name': 'Helvetica',
                    'valign': 'vcenter',
                    'bg_color': '#404040',  
                    'font_color': 'white'
                }),
                'In Progress': workbook.add_format({
                    'font_name': 'Helvetica',
                    'valign': 'vcenter',
                    'bg_color': '#1E90FF',  
                    'font_color': 'white'
                }),
                'Review': workbook.add_format({
                    'font_name': 'Helvetica',
                    'valign': 'vcenter',
                    'bg_color': '#FF69B4',  
                    'font_color': 'white'
                }),
                'Internally Approved': workbook.add_format({
                    'font_name': 'Helvetica',
                    'valign': 'vcenter',
                    'bg_color': '#3CB371',  
                    'font_color': 'white'
                }),
                'Client Approved': workbook.add_format({
                    'font_name': 'Helvetica',
                    'valign': 'vcenter',
                    'bg_color': '#228B22',  
                    'font_color': 'white'
                }),
                'On Hold': workbook.add_format({
                    'font_name': 'Helvetica',
                    'valign': 'vcenter',
                    'bg_color': '#DC143C',  
                    'font_color': 'white'
                })
            }           
            status_options = list(status_formats.keys())
            
            # Add status dropdowns 
            for col in range(1, 7):  

                # Add dropdown validation to status cell
                worksheet.data_validation(
                    current_row + 3, col, current_row + 3, col, 
                    {
                    'validate': 'list',
                    'source': status_options
                    }
                )
                
                # Write default status 
                worksheet.write(current_row + 3, col, 'Not Started', status_cell_format)  

            # Add conditional formatting for status cells
            for status, format_obj in status_formats.items():
                if status != 'Not Started':  # Skip default format
                    worksheet.conditional_format(current_row + 3, 1, current_row + 3, 6, {
                        'type': 'cell',
                        'criteria': 'equal to',
                        'value': f'"{status}"',
                        'format': format_obj
                    })

            # Add artist dropdowns 
            for col in range(1, 7):
                worksheet.data_validation(
                    current_row + 4, col, current_row + 4, col,
                    {
                    'validate': 'list', 
                    'source': '=Artists'
                    }
                )

                # Write default artist
                worksheet.write(current_row + 4, col, 'Not Assigned', artist_default_format) 

            # Add conditional formatting for artist cells
            artist_row = current_row + 4
            top_left = f'B{artist_row + 1}'  # A1 ref for the top-left of this CF range (Excel rows are 1-based)

            worksheet.conditional_format(
                artist_row, 1,    
                artist_row, 6,    
                    {   
                    'type': 'formula',
                    'criteria': f'=AND({top_left}<>"", {top_left}<>"Not Assigned")',
                    'format': artist_assigned_format,
                    }
                )
            
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
            length_info = str(self.shot_dict[shot_name][11])  
            frame_count = length_info.split(' - ')[1].split(' ')[0]  
            
            worksheet.write(current_row + 1, 3, source_name, metadata_value_format)  
            worksheet.write(current_row + 1, 4, source_tc, metadata_value_format)
            worksheet.write(current_row + 1, 5, seq_tc, metadata_value_format)
            worksheet.write(current_row + 1, 6, frame_count, metadata_value_format)
            
            # Set divider row height to be very small
            worksheet.set_row(current_row + 5, 15)  #second number is pixel height
            
            # Add black divider across all columns
            for col in range(7):  # Columns A through H
                worksheet.write(current_row + 5, col, '', divider_format)
            
            # Move to next group (5 rows for content + 1 for divider)
            current_row += 6

        #### End of shots loop ####

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
