var shell = new ActiveXObject("WScript.shell")
var fso = new ActiveXObject("Scripting.FileSystemObject")
var fsu = DOpus.FSUtil()
var stt = DOpus.Create().StringTools()

function OnInit(/* ScriptInitData */ data) {
  data.name = "File MIME type column using file.exe"
  data.desc = "Shows the MIME type using the file.exe utility"
  data.default_enable = true
  data.config_desc = DOpus.NewMap()
  data.config_desc("debug") = "Print debug messages to the script log"
  data.config.debug = false
  data.config_desc("fileExeFullName") = "Path to the 'file.exe' utility. You can get this utility by installing 'Git for Windows': https://git-scm.com/download/win"
  data.config.fileExeFullName = "%ProgramFiles%/Git/usr/bin/file.exe"
  data.version = "0.0-dev";
  data.url = "https://github.com/PolarGoose/DirectoryOpus-TabLabelizer-plugin";

  var cmd = data.AddColumn()
  cmd.name = "File MIME type column using file.exe"
  cmd.method = "OnMimeTypeColumnDataRequested"
  cmd.label = "MIME type"
  cmd.autorefresh = true
  cmd.justify = "right"

  var cmd = data.AddColumn()
  cmd.name = "File encoding column using file.exe"
  cmd.method = "OnEncodingColumnDataRequested"
  cmd.label = "Enconding"
  cmd.autorefresh = true
  cmd.justify = "right"
}

function OnMimeTypeColumnDataRequested(/* ScriptColumnData */ data) {
  var fileFullName = data.item.realpath
  debug("OnFileType: fileFullName=" + fileFullName)

  try {
    var fileType = getFileType(fileFullName)
    debug("fileType=" + fileType)
    data.value = fileType
  } catch (e) {
    debug("Exception:" + e)
    data.value = "<error>"
  }
}

function OnEncodingColumnDataRequested(/* ScriptColumnData */ data) {
  var fileFullName = data.item.realpath
  debug("OnFileType: fileFullName=" + fileFullName)

  try {
    var fileType = getEncoding(fileFullName)
    debug("fileType=" + fileType)
    data.value = fileType
  } catch (e) {
    debug("Exception:" + e)
    data.value = "<error>"
  }
}

function getFileType(/* Path */ fileFullName) {
  var fileCommandLineArguments = "--mime-type --brief "
  return runFileExeAndReturnOutput(fileFullName, fileCommandLineArguments)
}

function getEncoding(/* Path */ fileFullName) {
  var fileCommandLineArguments = "--mime --brief "
  var output = runFileExeAndReturnOutput(fileFullName, fileCommandLineArguments)
  if(output === "") {
    return ""
  }

  // The output look like: "text/plain; charset=us-ascii". We need to get "us-ascii"
  var match = /charset=([^;\r\n]+)/.exec(output)
  if (match && match.length > 1) {
    return match[1]
  }

  throw "Failed to get the text encoding. Output of file.exe: " + output
}

function runFileExeAndReturnOutput(/* Path */ fileFullName, /* string */ fileCommandLineArguments) {
  // file.exe tool doesn't work for ftp and UNC paths
  if (fileFullName.pathpart.substr(0, 2) === "\\\\" || fileFullName.pathpart.substr(0, 3) === "ftp") {
    debug("Skip UNC or FTP path")
    return "";
  }

  var command = '"' + Script.config.fileExeFullName + '" ' + fileCommandLineArguments + '"' + fileFullName + '"'
  return runCommandAndReturnOutput(command)
}

function runCommandAndReturnOutput(/* string */ command) {
  var tempFileFullName = fsu.GetTempFilePath()
  var cmdLine = 'cmd.exe /c "' + command + ' > "' + tempFileFullName + '""'
  debug("shell.run " + cmdLine)

  try {
    var exitCode = shell.run(cmdLine, 0, true)
    if (exitCode !== 0) {
      throw "Failed to execute the command. ExitCode=" + exitCode
    }

    var content = readAllText(tempFileFullName)
    if (content.indexOf("cannot open") === 0) {
      throw "File.exe failed. Output of File.exe: " + content
    }

    return content
  }
  finally {
    fso.DeleteFile(tempFileFullName)
  }
}

function readAllText(/* string */ fileFullName) {
  var handle = fsu.OpenFile(fileFullName)
  var content = stt.Decode(handle.Read(), "utf8")
  handle.Close()
  return content
}

function debug(text) {
  if (Script.config.debug) {
    DOpus.Output(text)
  }
}
