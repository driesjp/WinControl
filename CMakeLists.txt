cmake_minimum_required(VERSION 3.28)
project(WinControl C)

set(CMAKE_C_STANDARD 11)

add_executable(WinControl main.c
        wincontrol.h
        wincontrol.c)
