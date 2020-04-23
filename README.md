# MonitorWork
Monitor if you are working or slacking of, saves  what you are doing and classifies it acording to a given list


remove -example to see example info.json file.

A csv is generated where in each column the current action is classified, the date as well as if the computer is being used.

only works in windows


.\env\Scripts\pyinstaller --onefile .\monitorer.py

pip install --upgrade 'setuptools<45.0.0'


https://github.com/pypa/setuptools/issues/1963

Commnad to build
.\env\Scripts\pyinstaller -y --onefile --noconsole  .\monitorer.py
 .\env\Scripts\pyinstaller -y --onefile  .\monitorer.py  --icon=eye.ico
 pip freeze > requirements.txt