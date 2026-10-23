vcpkg_from_github(
    OUT_SOURCE_PATH SOURCE_PATH
    REPO PBfordev/wxAutoExcel
    REF 2152aab761c0af1aa543da4485cfb5a3384ca17a
    SHA512 9dde22f9ea5aa8459e6b76fe88faa47db045a4bb470c47e7962dd7855ccf41281a6e955393792fb8d6a46345c705bcaa6e7cc1e8e72a9c1375f7a7eacd6abc04
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
)

file(INSTALL "${CMAKE_CURRENT_LIST_DIR}/usage" DESTINATION "${CURRENT_PACKAGES_DIR}/share/${PORT}")
vcpkg_install_copyright(FILE_LIST "${SOURCE_PATH}/LICENSE.txt")