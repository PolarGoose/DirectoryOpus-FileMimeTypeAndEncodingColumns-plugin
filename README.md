# DirectoryOpus-FileMimeTypeAndEncodingColumns-plugin
This plugin adds two columns, `MIME type` and `Encoding` using `FileMimeTypeDetector` from [DOpus-Scripting-Extensions](https://github.com/PolarGoose/DOpus-Scripting-Extensions).<br>
The plugin uses the content of a file and not its extension.<br>
The screenshot below shows how the encoding and MIME type are correctly determined, even for files with incorrect extensions.
![Example](doc/screenshot.png)

# Extra Features
* Mime-type aliases can be specified by modifying `mimeTypeAliases` dictionary at the beginning of the script

# Note
* The encoding doesn't show if a file has BOM or not. You can use [DirectoryOpus-TextFileEncoding-plugin](https://github.com/PolarGoose/DirectoryOpus-TextFileEncoding-plugin) if you need this functionality.

# Prerequisites
* You need to have DOpus-Scripting-Extensions installed: download it from the [release page](https://github.com/PolarGoose/DOpus-Scripting-Extensions/releases)

# Limitations
* The script doesn't support DOpus portable mode because it requires a PC to have `DOpus-Scripting-Extensions` installed.
* The plugin is relatively fast but not instantaneous. Thus, for a folder with many files, you might need to wait a little bit for all files to get its MIME type.

# How to use
* Make sure [DOpus-Scripting-Extensions](https://github.com/PolarGoose/DOpus-Scripting-Extensions) are installed
* Download the `js` file from the [latest release](https://github.com/PolarGoose/DirectoryOpus-FileMimeTypeColumn-plugin/releases)
* Copy the `js` file to the `%AppData%\GPSoftware\Directory Opus\Script AddIns` folder
* The extra columns will become available in the `Settings`->`File Display Columns`->`Appearance`->`Columns:`.
* The script uses aliases for common MIME types. You can change them by modifying the `mimeTypeAliases` variable at the beginning of the script.

# References
* Discussion of this plugin on DOpus forum: [File MIME type column](https://resource.dopus.com/t/file-mime-type-column/53983)
