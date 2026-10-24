vcpkg_from_github(
    OUT_SOURCE_PATH SOURCE_PATH
    REPO PBfordev/wxAutoExcel
    REF ac723d56220405bd4c439860286809c9e2ead059
    SHA512 d7ad75e9aa1e9a2285210fe16aa801116f4e9da59495d40b095e2c344ef1182afb504fb8ef4667b99b264b444c4c80446ca71bc43e1061e5417a319922abb67e
    HEAD_REF master
)

set(OPTIONS "")

if(VCPKG_CRT_LINKAGE STREQUAL "static")
    list(APPEND OPTIONS -DwxAutoExcel_BUILD_USE_STATIC_RUNTIME=ON)
endif()

vcpkg_cmake_configure(
    SOURCE_PATH "${SOURCE_PATH}"
    OPTIONS
        ${OPTIONS}
        -DwxAutoExcel_BUILD_SAMPLES=OFF
        -DwxAutoExcel_BUILD_TESTS=OFF
        -DwxAutoExcel_BUILD_INSTALL=ON
)

vcpkg_cmake_install()
vcpkg_cmake_config_fixup(CONFIG_PATH lib/cmake/wxAutoExcel)
vcpkg_copy_pdbs()

file(REMOVE_RECURSE
        ${CURRENT_PACKAGES_DIR}/lib/cmake
        ${CURRENT_PACKAGES_DIR}/debug/lib/cmake
        ${CURRENT_PACKAGES_DIR}/debug/include
        ${CURRENT_PACKAGES_DIR}/debug/share
)

file(INSTALL "${CMAKE_CURRENT_LIST_DIR}/usage" DESTINATION "${CURRENT_PACKAGES_DIR}/share/${PORT}")
vcpkg_install_copyright(FILE_LIST "${SOURCE_PATH}/LICENSE.txt")