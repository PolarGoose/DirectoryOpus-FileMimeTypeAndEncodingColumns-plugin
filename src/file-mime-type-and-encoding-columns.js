var mimeTypeAliases = {
  "application/vnd.microsoft.portable-executable": "exe",
  "application/x-wine-extension-ini": "ini",
  "application/octet-stream": "binary",
  "inode/x-empty": "empty",
  "text/plain": "text",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation": "pptx",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
}

var shell = new ActiveXObject("WScript.shell")
var fso = new ActiveXObject("Scripting.FileSystemObject")
var fsu = DOpus.FSUtil()
var stt = DOpus.Create().StringTools()

function OnInit(/* ScriptInitData */ data) {
  data.name = "File_MIME_type_and_Encoding_columns_using_file.exe"
  data.desc = "Shows the MIME type using the file.exe utility"
  data.default_enable = true
  data.config_desc = DOpus.NewMap()
  data.config_desc("debug") = "Print debug messages to the script log"
  data.config.debug = false
  data.config_desc("fileExeFullName") = "Path to the 'file.exe' utility. You can get this utility by installing 'Git for Windows': https://git-scm.com/download/win"
  data.config.fileExeFullName = "%ProgramFiles%/Git/usr/bin/file.exe"
  data.config_desc("fileExeMagicFileFullName") = "Specify the path to the 'magic.mgc' file. This is only needed if you use a portable version of 'file.exe'. Keep this empty if you use 'Git for Windows'."
  data.config.fileExeMagicFileFullName = ""
  data.config_desc("ignoreBinaryEncoding") = "Don't show 'binary' encoding. Leave the Encoding column empty instead."
  data.config.ignoreBinaryEncoding = false
  data.version = "0.0-dev";
  data.url = "https://github.com/PolarGoose/DirectoryOpus-TabLabelizer-plugin";

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
    var output = getFileMimeTypeAndEncoding(filePath)
    debug("output from file.exe: " + output)

    // If file.exe fails it might print "cannot open ..." to the output.
    if (output.indexOf("cannot open") === 0) {
      throw "File.exe failed. Output of File.exe: " + output
    }

    // at this point output should look like: "text/plain; charset=us-ascii\n"

    var re = /^(.+);\scharset=(.+)\n$/.exec(output);
    if(!re) {
      throw "Failed to parse the output"
    }

    debug("MimeType=" + re[1] + "; Encoding=" + re[2])

    mimeType = re[1]
    encoding = re[2]

    mimeTypeAlias = mimeTypeAliases[mimeType]
    if(mimeTypeAlias) {
      mimeType = mimeTypeAlias
    }

    if(re[1] === "application/vnd.microsoft.portable-executable") {
      var arch = getExeOrDllArchitecture(filePath)
      if(arch !== null) {
        mimeType += " " + arch
      }
      if(isManagedExeOrDll(filePath)) {
        mimeType += " .Net"
      }
    }

    if(encoding === "utf-8") {
      var handle = fsu.OpenFile(filePath)
      var bom = handle.Read(3)
      handle.Close()
      if(bom(0) === 0xEF && bom(1) === 0xBB && bom(2) === 0xBF) {
        encoding = "utf-8 bom"
      }
    }

    if(encoding === "binary" && Script.config.ignoreBinaryEncoding) {
      encoding = ""
    }
  } catch (e) {
    debug("Exception: " + e)
  }

  data.columns("MIME type").value = mimeType;
  data.columns("Encoding").value = encoding;
}

function getFileMimeTypeAndEncoding(/* Path */ filePath) {
  var command = '"' + Script.config.fileExeFullName + '" ' + "--mime --brief " + '"' + filePath + '"'
  if(Script.config.fileExeMagicFileFullName !== "") {
    command += ' --magic-file "' + Script.config.fileExeMagicFileFullName + '"'
  }
  return runCommandAndReturnOutput(command)
}

function runCommandAndReturnOutput(/* string */ command) {
  var tempFileFullName = fsu.GetTempFilePath()
  var cmdLine = 'cmd.exe /c "' + command + ' > "' + tempFileFullName + '""' + " 2>&1"
  debug("shell.run " + cmdLine)

  try {
    var exitCode = shell.run(cmdLine, 0, true)
    var output = readAllTextIfFileExists(tempFileFullName)
    if (exitCode !== 0) {
      throw "Failed to execute the command. ExitCode=" + exitCode + ". Output: " + output
    }
    return output
  }
  finally {
    fso.DeleteFile(tempFileFullName)
  }
}

function readAllTextIfFileExists(/* Path */ filePath) {
  if(!fso.FileExists(filePath)) {
    return ""
  }

  var handle = fsu.OpenFile(filePath)
  var content = stt.Decode(handle.Read(), "utf8")
  handle.Close()
  return content
}

function isUncOrFtpPath(/* Path */ filePath) {
  var fileFullName = String(filePath)
  return fileFullName.substr(0, 2) === "\\\\" || fileFullName.substr(0, 3) === "ftp"
}

function isManagedExeOrDll(/* Path */ exeOrDllPath) {
  var command = "powershell.exe -ExecutionPolicy Bypass -Command [System.Reflection.AssemblyName]::GetAssemblyName(\\\"" + exeOrDllPath + "\\\")"
  debug("excecute: " + command)
  var res = shell.run(command, 0, true);
  return res === 0;
}

function getExeOrDllArchitecture(exeOrDllPath) {
  var ado = new ActiveXObject("ADODB.Stream");
  ado.Type = 1; // binary
  ado.Open();
  ado.LoadFromFile(exeOrDllPath);

  // Read the DOS Header to find the offset to the PE header (e_lfanew)
  ado.Position = 0x3C;
  var eLfanewOffsetBytes = (new VBArray(ado.Read(4))).toArray();
  var eLfanewOffset = eLfanewOffsetBytes[0] + (eLfanewOffsetBytes[1] << 8) + (eLfanewOffsetBytes[2] << 16) + (eLfanewOffsetBytes[3] << 24);

  // Go to the COFF File Header (PE Signature is 4 bytes, so skip that)
  ado.Position = eLfanewOffset + 4;

  // Read the Machine field
  var machineBytes = (new VBArray(ado.Read(2))).toArray();
  var machine = machineBytes[0] + (machineBytes[1] << 8);

  switch (machine) {
    case 0x014c: return "x32";
    case 0x8664: return "x64";
    case 0x01c4: return "x32-ARM";
    case 0xaa64: return "x64-ARM64";
    default: return null;
  }
}

function debug(text) {
  if (Script.config.debug) {
    DOpus.Output(text)
  }
}
