start "" /B "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"

taskkill /im outlook.exe /f

"C:\Program Files\Python310\python.exe" -m pip install -r c:\threat\requirements.txt