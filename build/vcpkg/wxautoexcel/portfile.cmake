vcpkg_from_github(
    OUT_SOURCE_PATH SOURCE_PATH
    REPO PBfordev/wxAutoExcel
    REF 018636f86d3c5ff6d908cc9554e40710fe5cbcde
    SHA512 cc0ceaa95ea3e79cffe16e054b155aab758821c93fd2aa47b718e641adfb05af8647f01d49a4052cdb82a368d699dd039569c932214ef5d50a013df15cc4742a
    HEAD_REF 018636f86d3c5ff6d908cc9554e40710fe5cbcde
)

set(OPTIONS "")

if(VCPKG_CRT_LINKAGE STREQUAL "static")
    list(APPEND OPTIONS -DwxAutoExcel_BUILD_USE_STATIC_RUNTIME=ON)
endif()

vcpkg_cmake_configure(
    SOURCE_PATH "${SOURCE_PATH}"
    OPTIONS
        -DwxAutoExcel_BUILD_SAMPLES=OFF
        -DwxAutoExcel_BUILD_INSTALL=ON
)

vcpkg_cmake_install()
vcpkg_cmake_config_fixup(CONFIG_PATH lib/cmake/wxAutoExcel)
vcpkg_copy_pdbs()

file(REMOVE_RECURSE
        ${CURRENT_PACKAGES_DIR}/lib/cmake
        ${CURRENT_PACKAGES_DIR}/debug/lib/cmake
        ${CURRENT_PACKAGES_DIR}/debug/include
)

file(INSTALL "${CMAKE_CURRENT_LIST_DIR}/usage" DESTINATION "${CURRENT_PACKAGES_DIR}/share/${PORT}")
vcpkg_install_copyright(FILE_LIST "${SOURCE_PATH}/LICENSE.txt")


