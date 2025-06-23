# 意识DT舰长抢车位工具
pyinstaller --onefile --clean --windowed --name Yishi_DT --icon=dt.ico --add-data "dt.ico;." --add-data "bgm.mp3;." --exclude-module numpy --exclude-module pandas --exclude-module matplotlib vx.py
