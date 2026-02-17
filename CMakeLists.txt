cmake_minimum_required(VERSION 3.20)
project(eloquence_wrapper LANGUAGES C CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# --- MinHook static lib ---
add_library(minhook STATIC
  src/minhook/src/buffer.c
  src/minhook/src/hook.c
  src/minhook/src/trampoline.c
)

if(CMAKE_SIZEOF_VOID_P EQUAL 4)
  target_sources(minhook PRIVATE src/minhook/src/hde/hde32.c)
else()
  target_sources(minhook PRIVATE src/minhook/src/hde/hde64.c)
endif()

target_include_directories(minhook PUBLIC
  src/minhook/include
  src/minhook/src
)

# --- Sonic (time-stretch for rate boost) ---
add_library(sonic STATIC
  src/sonic/sonic.c
)
target_include_directories(sonic PUBLIC src/sonic)

# --- Wrapper DLL ---
add_library(eloquence_wrapper SHARED
  src/eloquence_wrapper.cpp
)

target_include_directories(eloquence_wrapper PRIVATE
  src/minhook/include
  src/sonic
)

target_link_libraries(eloquence_wrapper PRIVATE
  minhook
  sonic
  winmm
  user32
)

target_compile_definitions(eloquence_wrapper PRIVATE ELOQUENCE_WRAPPER_EXPORTS)
set_target_properties(eloquence_wrapper PROPERTIES OUTPUT_NAME "eloquence_wrapper")
