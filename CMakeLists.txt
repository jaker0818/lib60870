cmake_minimum_required(VERSION 3.10)
project(cs101_master_gui C)

# 设置 C 标准
set(CMAKE_C_STANDARD 99)
set(CMAKE_C_STANDARD_REQUIRED ON)

# lib60870 库路径
set(LIB60870_DIR "${CMAKE_CURRENT_SOURCE_DIR}/lib60870-C")

# 添加 lib60870 库
add_subdirectory(${LIB60870_DIR})

# GUI 程序
set(GUI_SRCS
    cs101_master_gui.c
)

# 创建可执行文件
add_executable(cs101_master_gui WIN32 ${GUI_SRCS})

# 包含头文件目录
target_include_directories(cs101_master_gui PRIVATE
    ${LIB60870_DIR}/config
    ${LIB60870_DIR}/src/inc/api
    ${LIB60870_DIR}/src/inc/internal
    ${LIB60870_DIR}/src/hal/inc
    ${LIB60870_DIR}/src/common/inc
)

# 链接 lib60870 库
target_link_libraries(cs101_master_gui PRIVATE
    lib60870
    ws2_32
)

# 链接 Windows 库
if(WIN32)
    target_link_libraries(cs101_master_gui PRIVATE
        user32
        gdi32
        comctl32
        setupapi
    )
endif()

# 设置输出目录
set_target_properties(cs101_master_gui PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}"
)

# 编译选项
if(MSVC)
    target_compile_options(cs101_master_gui PRIVATE /W3 /utf-8)
    # 禁用特定警告
    target_compile_definitions(cs101_master_gui PRIVATE _CRT_SECURE_NO_WARNINGS UNICODE _UNICODE)
else()
    target_compile_options(cs101_master_gui PRIVATE -Wall -Wextra)
    target_compile_definitions(cs101_master_gui PRIVATE UNICODE _UNICODE)
endif()
