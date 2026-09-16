import openpyxl
import json
import os
import subprocess
import sys

def main():
    script_dir = os.path.dirname(__file__)
    build_script = os.path.join(script_dir, 'build_site_from_excel.py')
    
    print("Menjalankan sinkronisasi website dari Excel...")
    subprocess.run([sys.executable, build_script], check=True)

if __name__ == '__main__':
    main()
