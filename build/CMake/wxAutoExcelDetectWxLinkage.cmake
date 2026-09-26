###############################################################################
# Purpose:     Detect whether the selected wxWidgets build uses shared libraries
# Author:      OpenAI Codex, under the direction of PB
# Generated:   AI-generated using GPT-5.6 Sol with High reasoning effort
# Copyright:   (c) 2026 PB <pbfordev@gmail.com>
# License:     MIT license
###############################################################################

function(wxAutoExcel_detect_wx_shared outputVariable)
  # Config-package mode provides component targets whose type directly
  # identifies the selected wxWidgets linkage.
  foreach(wxTarget wxWidgets::core wx::core)
    if(TARGET ${wxTarget})
      get_target_property(wxTargetType ${wxTarget} TYPE)
      if(wxTargetType STREQUAL "SHARED_LIBRARY")
        set(${outputVariable} ON PARENT_SCOPE)
        return()
      elseif(wxTargetType STREQUAL "STATIC_LIBRARY")
        set(${outputVariable} OFF PARENT_SCOPE)
        return()
      endif()
    endif()
  endforeach()

  # FindwxWidgets provides one aggregate interface target and adds WXUSINGDLL
  # when it selected a DLL build on Windows.
  if(TARGET wxWidgets::wxWidgets)
    get_target_property(wxDefinitions
      wxWidgets::wxWidgets INTERFACE_COMPILE_DEFINITIONS)
    if("WXUSINGDLL" IN_LIST wxDefinitions)
      set(${outputVariable} ON PARENT_SCOPE)
    else()
      set(${outputVariable} OFF PARENT_SCOPE)
    endif()
    return()
  endif()

  message(FATAL_ERROR
    "Cannot determine whether the selected wxWidgets build is static or shared")
endfunction()
