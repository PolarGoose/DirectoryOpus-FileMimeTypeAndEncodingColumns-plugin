var mimeTypeAliases = {
  "application/vnd.microsoft.portable-executable": "exe/dll",
  "application/x-wine-extension-ini": "ini",
  "application/octet-stream": "binary",
  "inode/x-empty": "empty",
  "text/plain": "text",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation": "pptx",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
}

var fileMimeTypeDetector = new ActiveXObject("DOpusScriptingExtensions.FileMimeTypeDetector")

function OnInit(/* ScriptInitData */ data) {
  data.name = "File_MIME_type_and_Encoding_columns_using_file.exe"
  data.desc = "Shows the MIME type and Encoding using FileMimeTypeDetector from https://github.com/PolarGoose/DOpus-Scripting-Extensions"
  data.default_enable = true
  data.config_desc = DOpus.NewMap()
  data.config_desc("debug") = "Print debug messages to the script log"
  data.config.debug = false
  data.config_desc("ignoreBinaryEncoding") = "Don't show 'binary' encoding. Leave the Encoding column empty instead."
  data.config.ignoreBinaryEncoding = false
  data.version = "0.0-dev";
  data.url = "https://github.com/PolarGoose/DirectoryOpus-FileMimeTypeAndEncodingColumns-plugin";

  var col = data.AddColumn()
  col.name = "MIME type"
  col.method = "OnColumnDataRequested"
  col.autorefresh = true
  col.justify = "right"
  col.multicol = true

  var col = data.AddColumn()
  col.name = "Encoding"
  col.method = "OnColumnDataRequested"
  col.autorefresh = true
  col.justify = "right"
  col.multicol = true
}

function OnColumnDataRequested(/* ScriptColumnData */ data) {
  var filePath = data.item.realpath
  debug("OnColumnDataRequested: filePath=" + filePath)

  if(isUncOrFtpPath(filePath) || data.item.is_dir) {
    debug("Skip UNC, FTP or directory path")
    data.columns("MIME type").value = "";
    data.columns("Encoding").value = "";
    return
  }

  var mimeType = ""
  var encoding = ""

  try {
    var mimeTypeAndEncoding = fileMimeTypeDetector.DetectMimeType(filePath)
    mimeType = mimeTypeAndEncoding.MimeType
    encoding = mimeTypeAndEncoding.Encoding
    debug("MimeType=" + mimeType + "; Encoding=" + encoding)

    mimeTypeAlias = mimeTypeAliases[mimeType]
    if(mimeTypeAlias) {
      mimeType = mimeTypeAlias
    }

    if(encoding === "binary" && Script.config.ignoreBinaryEncoding) {
      encoding = ""
    }
  } catch (e) {
    debug("Exception: " + e)
    mimeType = "?"
    encoding = "?"
  }

  data.columns("MIME type").value = mimeType;
  data.columns("Encoding").value = encoding;
}

function isUncOrFtpPath(/* Path */ filePath) {
  var fileFullName = String(filePath)
  return fileFullName.substr(0, 2) === "\\\\" || fileFullName.substr(0, 3) === "ftp"
}

function debug(text) {
  if (Script.config.debug) {
    DOpus.Output(text)
  }
}
