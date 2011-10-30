
  ' CPUsage.bas. A utility to measure and monitor the CPU usage
  ' simple and without any frustration!

  ' Created : 05/08/2011
  ' Last Modified : 06/08/2011

'----------------------------------------------------------------------------------------------------------------------------------------------

  nomainwin

  WindowWidth = 300
  WindowHeight = 120

  stylebits #cpu.pr, _SS_CENTERIMAGE or _
                     _SS_CENTER or _
                     _WS_THICKFRAME, _
                     0, _                                                                           ' Set main screen style
                     _WS_EX_DLGMODALFRAME or _WS_EX_CLIENTEDGE, _
                     0
  stylebits #cpu, _WS_SYSMENU, _WS_MAXIMIZEBOX, _WS_EX_CLIENTEDGE, 0

  menu #cpu, "&Help", "Help Topics", [help]
  menu #cpu, "&About", "About CPUsage", [about]

  statictext #cpu.pr, "", 8, 8, 275, 50

  open "CPUsage" for window_nf as #cpu

  #cpu, "trapclose [quit]"
  #cpu.pr, "!font arial 5 15"

  struct IdleTime, _
    highValue as ulong, _
    lowValue as ulong

  struct KernelTime, _
    highValue as ulong, _
    lowValue as ulong

  struct UserTime, _
    highValue as ulong, _
    lowValue as ulong

  global oldUsage

  open "kernel32" for dll as #k

  processorCount = val(GetEnvironmentVariable$("NUMBER_OF_PROCESSORS"))

'----------------------------------------------------------------------------------------------------------------------------------------------

  ' Test loop

  [cpuMeasure]

  select case processorCount
    case 1
        usageType$ = "Normal "
    case else
        usageType$ = "Average "
  end select

  #cpu.pr, usageType$ + "CPU Usage: " ; cpuLoad(processorCount) ; "    Processor count: " ; processorCount

  timer 250, [cpuMeasure]
  wait

'----------------------------------------------------------------------------------------------------------------------------------------------

  [quit]

  close #k
  close #cpu
  end

'----------------------------------------------------------------------------------------------------------------------------------------------

  function cpuLoad(NumberOfProcessors)

    calldll #k, "GetSystemTimes", _
        IdleTime as struct, _
        KernelTime as struct, _
        UserTime as struct, _
        result as boolean

    idle = GetLargeIntegerValue(IdleTime.highValue.struct, IdleTime.lowValue.struct)
    kernel = GetLargeIntegerValue(KernelTime.highValue.struct, KernelTime.lowValue.struct)
    user = GetLargeIntegerValue(UserTime.highValue.struct, UserTime.lowValue.struct)

    total = idle + kernel + user
    usage = int(((kernel + user) / total) * 100)
    usage = int(usage / NumberOfProcessors)                                                       ' # of processors
    cpuLoad = int((usage + oldUsage) / 2)                                                         ' Average Load
    oldUsage = usage

  end function

'----------------------------------------------------------------------------------------------------------------------------------------------

  function GetLargeIntegerValue(high, low)

    highHex$ = right$("0000"; dechex$(high), 4)
    lowHex$ = right$("0000"; dechex$(low), 4)
    GetLargeIntegerValue = hexdec(highHex$ ; lowHex$)

  end function

'----------------------------------------------------------------------------------------------------------------------------------------------

  function GetEnvironmentVariable$(lpName$)

    ' Get the value of an environment variable

    nSize = 1024
    lpBuffer$ = space$(nSize)

    calldll #kernel32, "GetEnvironmentVariableA", _
        lpName$   as ptr, _
        lpBuffer$ as ptr, _
        nSize     as ulong, _
        result    as ulong

    if result > 0 then GetEnvironmentVariable$ = left$(lpBuffer$, result)

  end function

'----------------------------------------------------------------------------------------------------------------------------------------------

  [help]

  run "NOTEPAD CPUsage.txt"

  wait

'----------------------------------------------------------------------------------------------------------------------------------------------

  [about]

  WindowWidth = 270
  WindowHeight = 215

  stylebits #info, 0, 0, _WS_EX_TOPMOST or _WS_EX_TOOLWINDOW or _WS_EX_CLIENTEDGE, 0

  statictext #info.ver, "CPUsage is a small and simple program for monitoring the usage "_
                        + "of the CPU (normal for 1 processor - average for 2 or more)."_
                        + chr$(13) + chr$(13) + "Version 1.0, 2011"_
                        + chr$(13) + chr$(13) + "Final Release: 06/08/2011"_
                        + chr$(13) + chr$(13) + "For more information about the program, "_
                        + "see the CPUsage help topics.", 10, 10, 230, 180

  open "About CPUsage" for dialog_nf_modal as #info

  #info, "trapclose [exitInfo]"
  #info.ver, "!font arial 5 15"

  wait

'----------------------------------------------------------------------------------------------------------------------------------------------

  [exitInfo]

  close #info
  timer 125, [cpuMeasure]
  wait
