from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but they might need fine-tuning.
build_exe_options = {
    "excludes": ["tkinter", "unittest"],
    "zip_include_packages": ["encodings", "PySide6", "shiboken6"],
}

setup(
    name="guifoo",
    version="0.1",
    description="SNR Merge",
    options={"build_exe": build_exe_options},
    executables=[Executable("snr_merge.py", base="gui")],
)