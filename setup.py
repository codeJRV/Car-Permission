
from cx_Freeze import setup, Executable


setup(
    name = "CarPermissionBot" ,
    version = "0.1" ,
    description = " Car Permission Bot " ,
    executables = [Executable("CarPermissionBot.py")]  ,
)

