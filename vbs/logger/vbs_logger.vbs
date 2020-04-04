class Logger
'------------------------------------------------------
'@ description : Logger
'@ auther : t.n
'@ version : 1.0
'@ since : 2020.4.4
'@ update : 2020.4.4
'------------------------------------------------------

    private fso
    Private f
    private logName
    private logPath
    private scriptPath

    'Event procedure
    public sub class_Initialize()
        set fso = CreateObject("Scripting.FileSystemObject")
        logName = ""
        logPath = ""
        scriptPath = fso.getParentFolderName(WScript.ScriptFullName)
    end sub

    public sub class_Terminate()
        set fso = Nothing
    end sub

    'property
    public function setName(byval name)
        logName = name
        set setName = Me
    end function

    public function setLogpath(byval path)
        logPath = path
        set setLogpath = Me
    end function

    'method
    public sub error(byval message)
        writeLog "error", message
    end sub

    public sub warning(byval message)
        writeLog "warning", message
    end sub

    public sub info(byval message)
        writeLog "info", message
    end sub

    private sub writeLog(byval level, byval message)

        if (logPath = "") Then
            logPath = scriptPath & "\log"
            if fso.FolderExists(logPath) = False Then
                fso.CreateFolder logPath
            End If
        Else
            If Right(logPath, 1) = "\" Then
                logPath = Left(logPath,len(logPath)-1)
            End If
            if fso.FolderExists(logPath) = False Then
                logPath = scriptPath & "\log"
                if fso.FolderExists(logPath) = False Then
                    fso.CreateFolder logPath
                End If
            End If
        End If

        logPath2 = logPath
        logName2 = logName & Replace(Left(Now(),10), "/", "") & ".log"

        Set f = fso.OpenTextFile(logPath2 & "\" & logName2 ,8 , True)
        f.WriteLine """" & now() & """,""" & level & """,""" & message & """"
        f.Close
    end sub
end class

'test'
dim log
set log = new Logger
log.setName("")
log.setLogpath("")

log.error "test error message"
log.warning "test warning message"
log.info "test info message"
