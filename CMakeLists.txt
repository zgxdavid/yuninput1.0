cmake_minimum_required(VERSION 3.21)
project(yuninput LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

if(MSVC)
    # Bundle MSVC runtime into binaries to avoid target-machine VC redist dependency.
    set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")
endif()

add_library(yuninput SHARED
    src/DllMain.cpp
    src/Globals.cpp
    src/TextService.cpp
    src/CompositionEngine.cpp
    src/CandidateWindow.cpp
    yuninput.rc
    yuninput.def
)

target_include_directories(yuninput PRIVATE include)

target_compile_definitions(yuninput PRIVATE
    UNICODE
    _UNICODE
    NOMINMAX
    WIN32_LEAN_AND_MEAN
)

target_link_libraries(yuninput PRIVATE
    ole32
    oleaut32
    advapi32
    uuid
)

add_executable(yuninput_sort_probe
    src/CompositionEngine.cpp
    tools/sort_probe.cpp
)

target_include_directories(yuninput_sort_probe PRIVATE include)

target_compile_definitions(yuninput_sort_probe PRIVATE
    UNICODE
    _UNICODE
    NOMINMAX
    WIN32_LEAN_AND_MEAN
)

