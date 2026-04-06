import os
import sys

# Ensure numpy doesn't think it's running from a source directory
# by setting __file__ and cwd properly for PyInstaller bundles
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))

    # QtWebEngine needs writable cache/data dirs — set them to TEMP
    # This also prevents issues on network/read-only drives
    tmp = os.path.join(os.environ.get('LOCALAPPDATA', os.environ.get('TEMP', '.')), 'PainelGestao')
    os.makedirs(tmp, exist_ok=True)
    os.environ.setdefault('QTWEBENGINE_CHROMIUM_FLAGS', '--disable-gpu --no-sandbox')
    os.environ['QTWEBENGINE_DISK_CACHE_DIR'] = os.path.join(tmp, 'cache')
    os.environ['QTWEBENGINE_DATA_DIR'] = os.path.join(tmp, 'data')
