dependencies
  python -m pip install PyPDF2
  python -m pip install move
  python -m pip install datetime
  python -m pip install requests
  python -m pip install ctypes

generate .exe:
  # install dependencies
  python -m pip install pyinstaller
  
  # open powershell, goto script.py location and execute:
  pyinstaller --onefile -c '.\script.py'

