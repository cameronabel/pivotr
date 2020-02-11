import app
import sys

if sys.platform == 'win32' and sys.executable.split('\\')[-1] == 'pythonw.exe':
    sys.stdout = open('log.txt', 'w')
    sys.stderr = open('err.txt', 'w')

if __name__ == '__main__':
    app.Pivotr().run()
