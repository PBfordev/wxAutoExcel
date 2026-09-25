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
  ${wxWidgets_LIBRARIES}
)

set_target_properties(wxAutoExcel_tests PROPERTIES
  FOLDER "wxAutoExcel Tests"
  RUNTIME_OUTPUT_DIRECTORY "${wxAutoExcel_RUNTIME_DIR}"
)

wxAutoExcel_adjust_target_flags(wxAutoExcel_tests)

add_test(
  NAME wxAutoExcel.Tests
  COMMAND wxAutoExcel_tests
)

set_tests_properties(wxAutoExcel.Tests PROPERTIES RUN_SERIAL TRUE)
