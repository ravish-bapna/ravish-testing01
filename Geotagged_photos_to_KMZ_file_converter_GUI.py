import os,sys,time,zipfile,shutil,cv2,wx

from PIL import Image
from PIL.ExifTags import TAGS, GPSTAGS

class ProgressDialog(wx.Dialog):
    def __init__(self, parent, title):
        super(ProgressDialog, self).__init__(parent, title=title, size=(300, 130))
        self.gauge = wx.Gauge(self, range=100, size=(250, 25), pos=(20, 20))
        self.status_text = wx.StaticText(self, label='', pos=(80, 60))
        self.Show()

class PhotoProcessor(wx.Frame):
    def __init__(self, parent, title):
        style = wx.DEFAULT_FRAME_STYLE & ~wx.MAXIMIZE_BOX  # Remove the maximize button
        super(PhotoProcessor, self).__init__(parent, title=title, size=(615, 312), style=style)
        self.panel = wx.Panel(self)
        self.create_widgets()

    def create_widgets(self):
        
        # Input folder
        self.input_folder_label = wx.StaticText(self.panel, label="Input folder of geotagged photos :", pos=(10, 12))
        self.input_folder_text = wx.TextCtrl(self.panel, style=wx.TE_PROCESS_ENTER, pos=(200, 10), size=(300, -1))
        self.input_folder_text.SetBackgroundColour(wx.Colour(255, 255, 255))  # Set background color to white

        self.input_folder_button = wx.Button(self.panel, label="Browse", pos=(510, 10), size=(80, -1))
        self.input_folder_button.Bind(wx.EVT_BUTTON, self.on_browse_input_folder)

        # Thumbnail size dropdown menu
        self.thumbnail_size_label = wx.StaticText(self.panel, label="Thumbnail size (in pixels) :", pos=(10, 42))
        self.thumbnail_size_choices = ["400", "500", "600", "700", "800", "900", "1000"]
        self.thumbnail_size_dropdown = wx.Choice(self.panel, choices=self.thumbnail_size_choices, pos=(200, 40))
        self.thumbnail_size_dropdown.SetSelection(2)
        self.thumbnail_size_dropdown.Bind(wx.EVT_CHOICE, self.on_select_thumbnail_size)

        # Thumbnail quality dropdown menu
        self.thumbnail_quality_label = wx.StaticText(self.panel, label="Thumbnail quality :", pos=(10, 72))
        self.thumbnail_quality_choices = ["10", "20", "30", "40", "50", "60", "70", "80", "90"]
        self.thumbnail_quality_dropdown = wx.Choice(self.panel, choices=self.thumbnail_quality_choices, pos=(200, 70))
        self.thumbnail_quality_dropdown.SetSelection(4)
        self.thumbnail_quality_dropdown.Bind(wx.EVT_CHOICE, self.on_select_thumbnail_quality)


        # Output KMZ file
        self.output_kmz_label = wx.StaticText(self.panel, label="Output KMZ file :", pos=(10, 102))
        self.output_kmz_text = wx.TextCtrl(self.panel, style=wx.TE_PROCESS_ENTER, pos=(200, 100), size=(300, -1))
        self.output_kmz_text.SetBackgroundColour(wx.Colour(255, 255, 255))  # Set background color to white
        self.output_kmz_button = wx.Button(self.panel, label="Browse", pos=(510, 100), size=(80, -1))
        self.output_kmz_button.Bind(wx.EVT_BUTTON, self.on_browse_output_kmz)
        
        # Output text file
        self.input_text_file_label = wx.StaticText(self.panel, label="Output log text file :", pos=(10, 132))
        self.input_text_file_text = wx.TextCtrl(self.panel, style=wx.TE_PROCESS_ENTER, pos=(200, 130), size=(300, -1))
        self.input_text_file_text.SetBackgroundColour(wx.Colour(255, 255, 255))  # Set background color to white
        self.input_text_file_button = wx.Button(self.panel, label="Browse", pos=(510, 130), size=(80, -1))
        self.input_text_file_button.Bind(wx.EVT_BUTTON, self.on_browse_input_text_file)

        # Checkbox to save non-geotagged photos filename
        self.save_checkbox = wx.CheckBox(self.panel, label="Create log for filenames of non-geotagged photos", pos=(10, 162))
        self.save_checkbox.Bind(wx.EVT_CHECKBOX, self.on_checkbox)

        # Button to process photos
        self.process_button = wx.Button(self.panel, label="Generate KMZ file", pos=(207, 200), size=(200, 30))
        self.process_button.Bind(wx.EVT_BUTTON, self.process_photos)

        # Divide line
        line = wx.StaticLine(self.panel, pos=(0, 240), size=(615, 2), style=wx.LI_HORIZONTAL)
        
        # Developer information
        developer_info = wx.StaticText(self.panel, label="Developed by : Ravish Bapna (Dar Pune). Software version : 1.0.",
                                       pos=(150, 250))
        developer_info.SetForegroundColour(wx.Colour(0, 0, 255))  # Set font color to blue

        # Disable input text file elements at start
        self.input_text_file_label.Disable()
        self.input_text_file_text.Disable()
        self.input_text_file_button.Disable()

        self.Center()
        self.Show()

    def on_browse_input_folder(self, event):
        dlg = wx.DirDialog(self, "Choose the input folder", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            self.input_folder_text.SetValue(dlg.GetPath())
        dlg.Destroy()

    def on_select_thumbnail_size(self, event):
        selected_size_index = self.thumbnail_size_dropdown.GetSelection()
        selected_size = self.thumbnail_size_choices[selected_size_index]
        self.thumbnail_size_dropdown.SetStringSelection(selected_size)
    
    def on_select_thumbnail_quality(self, event):
        selected_quality_index = self.thumbnail_quality_dropdown.GetSelection()
        selected_quality = self.thumbnail_quality_choices[selected_quality_index]
        self.thumbnail_quality_dropdown.SetStringSelection(selected_quality)

    def on_browse_input_text_file(self, event):
        dlg = wx.FileDialog(self, "Choose the input text file", wildcard="Text files (*.txt)|*.txt",
                            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        if dlg.ShowModal() == wx.ID_OK:
            self.input_text_file_text.SetValue(dlg.GetPath())
        dlg.Destroy()

    def on_browse_output_kmz(self, event):
        dlg = wx.FileDialog(self, "Choose the output KMZ file", wildcard="KMZ files (*.kmz)|*.kmz",
                            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT)
        if dlg.ShowModal() == wx.ID_OK:
            self.output_kmz_text.SetValue(dlg.GetPath())
        dlg.Destroy()

    def on_checkbox(self, event):
        state = self.save_checkbox.GetValue()
        self.input_text_file_label.Enable(state)
        self.input_text_file_text.Enable(state)
        self.input_text_file_button.Enable(state)

    def process_photos(self, event):
        photo_folder_path = self.input_folder_text.GetValue()
        thumbnail_image_size = float(self.thumbnail_size_choices[self.thumbnail_size_dropdown.GetSelection()])
        jpeg_quality = int(self.thumbnail_quality_choices[self.thumbnail_quality_dropdown.GetSelection()])

        if self.save_checkbox.GetValue():
            txt_file = self.input_text_file_text.GetValue()

        kmz_file = self.output_kmz_text.GetValue()
        save_txt = self.input_text_file_text.GetValue()

        # Factor to reduce the dimensions of image for icon
        icon_scale_factor = 0.03

        # Make lists if geotagged and non-geotagged photos
        geotagged_photos_list = []
        non_geotagged_photos_list = []
        image_coords = {}

        for filename in os.listdir(photo_folder_path):
            if filename.lower().endswith(".jpg") or filename.lower().endswith(".jpeg"):
                photo_path = os.path.join(photo_folder_path, filename)

                # check if image is geotagged
                with Image.open(photo_path) as image:
                    exif_data = image._getexif()
                    lat, lon = None, None
                    if exif_data is not None:
                        for tag, value in exif_data.items():
                            tag_name = TAGS.get(tag, tag)
                            if tag_name == 'GPSInfo':
                                # Extract latitude and longitude
                                if len(value):
                                    lat = value[2][0] + value[2][1] / 60.0 + value[2][2] / 3600.0
                                    if value[1] == 'S':
                                        lat = -lat
                                    lon = value[4][0] + value[4][1] / 60.0 + value[4][2] / 3600.0
                                    if value[3] == 'W':
                                        lon = -lon

                                    break

                    if lat is None or lon is None:
                        non_geotagged_photos_list.append(filename)
                    else:
                        geotagged_photos_list.append(filename)
                        image_coords[filename] = {'Latitude': lat, 'Longitude': lon}

        if len(geotagged_photos_list) == 0:
            wx.MessageBox('No geotagged photo in input folder.',
                          'Error', wx.OK | wx.ICON_ERROR)
            return

        dir_name, file_name = os.path.split(kmz_file)
        kmz_folder_tmp = os.path.join(dir_name, 'KMZ_{}'.format(time.strftime("%y%m%d%H%M%S")))
        images_folder_tmp = os.path.join(kmz_folder_tmp, 'images')
        icons_folder_tmp = os.path.join(kmz_folder_tmp, 'icons')
        kml_file_tmp = os.path.join(kmz_folder_tmp, 'doc.kml')

        os.makedirs(kmz_folder_tmp)
        os.makedirs(images_folder_tmp)
        os.makedirs(icons_folder_tmp)

        kml_file = open(kml_file_tmp, "w")
        kml_file.write(
            """<?xml version="1.0" encoding="UTF-8"?>
        <kml xmlns="http://earth.google.com/kml/2.1">
        <Document>
        <name>Geotagged photos</name>""")

        # Hide the main window
        self.Hide()

        # Create progress bar dialog
        progress_dialog = ProgressDialog(self, "Processing photos")
        progress_dialog.Center()

        try:
            total_files = len(geotagged_photos_list)
            progress_dialog.gauge.SetRange(total_files)

            for index, filename in enumerate(geotagged_photos_list, start=1):
                
                # Update progress bar
                wx.Yield()
                progress_dialog.gauge.SetValue(index)                
                progress_dialog.status_text.SetLabel(f'Processing file {index} of {total_files}')
                
                photo_path = os.path.join(photo_folder_path, filename)

                # Load the image
                img = cv2.imread(photo_path)

                # Get the image dimensions
                height, width, _ = img.shape

                # Calculate the new dimensions for the icon image
                new_height = int(height * icon_scale_factor)
                new_width = int(width * icon_scale_factor)

                # Resize the image to make the icon image
                resized_img = cv2.resize(img, (new_width, new_height), interpolation=cv2.INTER_AREA)

                icon_filename = '{}_{}'.format(time.strftime("%y%m%d%H%M%S"), filename)
                icon_file = os.path.join(icons_folder_tmp, icon_filename)

                # Save the icon image
                cv2.imwrite(icon_file, resized_img, [cv2.IMWRITE_JPEG_QUALITY, jpeg_quality])

                # Calculate the new dimensions
                if height>=width:
                    image_scale_factor=thumbnail_image_size/float(height)
                else:
                    image_scale_factor=thumbnail_image_size/float(width)
                    
                new_height = int(height * image_scale_factor)
                new_width = int(width * image_scale_factor)

                # Resize the image
                resized_img = cv2.resize(img, (new_width, new_height), interpolation=cv2.INTER_AREA)

                # Save the resized image
                cv2.imwrite(os.path.join(images_folder_tmp, filename), resized_img,
                            [cv2.IMWRITE_JPEG_QUALITY, jpeg_quality])

                kml_file.write("""
            <Placemark>
            <Style><IconStyle><Icon><href>icons/{0}</href></Icon></IconStyle></Style>
            <name>{1}</name>
            <Snippet></Snippet>
            <description>
            <![CDATA[<img src="images/{2}" width="{3}" height="{4}">
            ]]>
            </description>
            <Point><coordinates>{5},{6}</coordinates></Point>
            </Placemark>""".format(icon_filename, os.path.splitext(filename)[0], filename, new_width, new_height,
                                   image_coords[filename]['Longitude'], image_coords[filename]['Latitude']))



            wx.Yield()  # Ensure that the progress bar updates before continuing

        finally:
            progress_dialog.Destroy()

        kml_file.write("""
        </Document>
        </kml>""")
        kml_file.close()

        with zipfile.ZipFile(kmz_file, 'w', zipfile.ZIP_DEFLATED) as zip:
            for root, dirs, files in os.walk(kmz_folder_tmp):
                for file in files:
                    file_path = os.path.join(root, file)
                    zip.write(file_path, os.path.relpath(file_path, kmz_folder_tmp))

        shutil.rmtree(kmz_folder_tmp)

        if len(non_geotagged_photos_list) > 0:

            if self.save_checkbox.GetValue():
                with open(txt_file, 'w') as f:
                    for current_photo in non_geotagged_photos_list:
                        f.write(current_photo)
                        f.write('\n')

        # Show the main window again
        self.Show()
        

if __name__ == '__main__':
    app = wx.App(False)
    frame = PhotoProcessor(None, "Geotagged photos to KMZ file converter")
    frame.SetSizeHints(615, 312, 615, 312)  # Set the constant size here
    frame.Show()
    app.MainLoop()  