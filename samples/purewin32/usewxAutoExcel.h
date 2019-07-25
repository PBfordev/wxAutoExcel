#ifndef USE_WXAUTOEXCEL_H_DEFINED
#define USE_WXAUTOEXCEL_H_DEFINED

/**
    A class that will accept messages redirected
    from wxLog*() calls. See RichedLogger in purewin32.cpp
    for an example of actual implementation.
**/
class MyLogger
{
public:
    enum Level //
    {
        FatalError, // wxLOG_FatalError: program can't continue, abort immediately
        Error,      // wxLOG_Error:      a serious error, user must be informed about it
        Warning,    // wxLOG_Warning:    user is normally informed about it but may be ignored
        Message,    // wxLOG_Message:    normal message (i.e. normal output of a non GUI app)
        Status,     // wxLOG_Status:     informational: might go to the status line of GUI app
        Info,       // wxLOG_Info:       informational message (a.k.a. 'Verbose')
        Debug,      // wxLOG_Debug:      never shown to the user, disabled in release mode
        Trace       // wxLOG_Trace:      trace messages are also only enabled in debug mode
    };

    MyLogger() {}
    virtual ~MyLogger() {}

    virtual void Log(int level, LPCWSTR message) = 0;
};

// see the source code for description
bool UsewxAutoExcel(MyLogger* logger, bool enableLogTimeStamp = true, bool enableTrace = true);

#endif // #ifndef USE_WXAUTOEXCEL_H_DEFINED