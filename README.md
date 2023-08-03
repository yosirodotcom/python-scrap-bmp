# Create Conda Env:
conda create --prefix C:\Users\ASUS\Documents\pythonenv\scrap_bmp_env python=3.9 --no-default-packages

# To activate this environment, use cmd prompt     
conda activate G:\myrepos\python-scrap-bmp\myenv
conda activate C:\Users\ASUS\Documents\pythonenv\scrap_bmp_env

# To deactivate an active environment, use
conda deactivate

# Install library
pip install selenium
pip install requests
pip install pyautogui
pip install img2pdf
pip install PIL