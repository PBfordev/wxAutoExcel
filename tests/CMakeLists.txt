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
  getter_lists.h
  wxExcel_getters.cpp
  wxExcel_setters.cpp
  wxExcelApplication.cpp
  wxExcelChart.cpp
  wxExcelFormatConditions.cpp
  wxExcelListObject.cpp
  wxExcelNames.cpp
  wxExcelObject_errors.cpp
  wxExcelPageSetup.cpp
  wxExcelRange_access.cpp
  wxExcelRange_find.cpp
  wxExcelRange_formula.cpp
  wxExcelRange_format.cpp
  wxExcelRange_merge.cpp
  wxExcelRange_objects.cpp
  wxExcelRange_operations.cpp
  wxExcelRangeData.cpp
  wxExcelRange_value.cpp
  wxExcelShape.cpp
  wxExcelSheets.cpp
  wxExcelSortFilter.cpp
  wxExcelWorkbook.cpp
  wxExcelWorksheet.cpp
  tests.rc
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

get_target_property(wxAutoExcel_CATCH2_TYPE Catch2::Catch2 TYPE)
if(wxAutoExcel_CATCH2_TYPE STREQUAL "SHARED_LIBRARY")
  add_custom_command(
    TARGET wxAutoExcel_tests
    POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
            "$<TARGET_FILE:Catch2::Catch2>"
            "$<TARGET_FILE_DIR:wxAutoExcel_tests>"
    VERBATIM
  )
endif()

if(TARGET wxAutoExcel_wx_dlls)
  add_dependencies(wxAutoExcel_tests wxAutoExcel_wx_dlls)
endif()

wxAutoExcel_adjust_target_flags(wxAutoExcel_tests)

add_test(
  NAME wxAutoExcel.Tests
  COMMAND wxAutoExcel_tests
)

set_tests_properties(wxAutoExcel.Tests PROPERTIES RUN_SERIAL TRUE)
