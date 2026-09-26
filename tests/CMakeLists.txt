find_package(Catch2 3.15 CONFIG QUIET)

if(NOT Catch2_FOUND)
  include(FetchContent)

  set(CATCH_BUILD_TESTING OFF CACHE BOOL "" FORCE)
  set(CATCH_INSTALL_DOCS OFF CACHE BOOL "" FORCE)
  set(CATCH_INSTALL_EXTRAS OFF CACHE BOOL "" FORCE)

  FetchContent_Declare(
    Catch2
    GIT_REPOSITORY https://github.com/catchorg/Catch2.git
    GIT_TAG 8b08d4d79514f45f7e4ce2a607ac9c94e920d1bb
  )
  FetchContent_MakeAvailable(Catch2)
endif()

if(MSVC AND wxAutoExcel_BUILD_USE_STATIC_RUNTIME AND TARGET Catch2)
  set_property(TARGET Catch2 PROPERTY
    MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>"
  )
endif()

add_executable(wxAutoExcel_tests
  test_main.cpp
  catch_wx_string_makers.h
  excel_test_fixture.h
  wxExcelRange_access.cpp
  wxExcelRange_formula.cpp
  wxExcelRange_format.cpp
  wxExcelRange_merge.cpp
  wxExcelRange_value.cpp
  wxExcelWorksheet.cpp
)

target_compile_features(wxAutoExcel_tests PRIVATE cxx_std_14)
target_link_libraries(wxAutoExcel_tests PRIVATE
  wxAutoExcel::wxAutoExcel
  Catch2::Catch2
)

set_target_properties(wxAutoExcel_tests PROPERTIES
  FOLDER "wxAutoExcel Tests"
  RUNTIME_OUTPUT_DIRECTORY "${wxAutoExcel_RUNTIME_DIR}"
)

if(wxAutoExcel_BUILD_TESTS)
  function(wxAutoExcel_copy_runtime_target target)
    get_target_property(targetType ${target} TYPE)
    if(targetType STREQUAL "SHARED_LIBRARY")
      add_custom_command(
        TARGET wxAutoExcel_tests
        POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy_if_different
                "$<TARGET_FILE:${target}>"
                "$<TARGET_FILE_DIR:wxAutoExcel_tests>"
        VERBATIM
      )
    endif()
  endfunction()

  wxAutoExcel_copy_runtime_target(Catch2::Catch2)

  if(wxAutoExcel_WX_SHARED)
    if(TARGET wxWidgets::core AND TARGET wxWidgets::base)
      wxAutoExcel_copy_runtime_target(wxWidgets::core)
      wxAutoExcel_copy_runtime_target(wxWidgets::base)
    elseif(TARGET wx::core AND TARGET wx::base)
      wxAutoExcel_copy_runtime_target(wx::core)
      wxAutoExcel_copy_runtime_target(wx::base)
    elseif(wxWidgets_LIB_DIR)
      file(GLOB wxAutoExcel_WX_DLL_CANDIDATES CONFIGURE_DEPENDS
        "${wxWidgets_LIB_DIR}/wxbase*.dll"
        "${wxWidgets_LIB_DIR}/wxmsw*ud_core_*.dll"
        "${wxWidgets_LIB_DIR}/wxmsw*u_core_*.dll"
      )
      set(wxAutoExcel_WX_RELEASE_DLLS)
      set(wxAutoExcel_WX_DEBUG_DLLS)
      foreach(dll IN LISTS wxAutoExcel_WX_DLL_CANDIDATES)
        get_filename_component(dllName "${dll}" NAME)
        if(dllName MATCHES "^wxbase[0-9]+u(d)?_(net|xml)_")
          continue()
        elseif(dllName MATCHES "^wx(base|msw)[0-9]+ud(_core)?_.*\\.dll$")
          list(APPEND wxAutoExcel_WX_DEBUG_DLLS "${dll}")
        elseif(dllName MATCHES "^wx(base|msw)[0-9]+u(_core)?_.*\\.dll$")
          list(APPEND wxAutoExcel_WX_RELEASE_DLLS "${dll}")
        endif()
      endforeach()

      if(wxAutoExcel_WX_RELEASE_DLLS AND wxAutoExcel_WX_DEBUG_DLLS)
        add_custom_command(
          TARGET wxAutoExcel_tests
          POST_BUILD
          COMMAND ${CMAKE_COMMAND} -E copy_if_different
                  "$<$<CONFIG:Debug>:${wxAutoExcel_WX_DEBUG_DLLS}>"
                  "$<$<NOT:$<CONFIG:Debug>>:${wxAutoExcel_WX_RELEASE_DLLS}>"
                  "$<TARGET_FILE_DIR:wxAutoExcel_tests>"
          COMMAND_EXPAND_LISTS
          VERBATIM
        )
      else()
        message(WARNING
          "Could not find the wxWidgets core and base runtime DLLs in "
          "${wxWidgets_LIB_DIR}; wxAutoExcel_tests may not be runnable"
        )
      endif()
    endif()
  endif()
endif()

wxAutoExcel_adjust_target_flags(wxAutoExcel_tests)

add_test(
  NAME wxAutoExcel.Tests
  COMMAND wxAutoExcel_tests
)

set_tests_properties(wxAutoExcel.Tests PROPERTIES RUN_SERIAL TRUE)
