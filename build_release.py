import os
import shutil
import subprocess
import zipfile

APP_NAME = "ClassroomArchiver"

dist_dir = "dist"
release_dir = "release"

exe_path = os.path.join(dist_dir, "main.exe")
cred_path = "credentials.json"

# 1. PyInstaller 実行
subprocess.run([
    "pyinstaller",
    "main.py",
    "--onefile",
    "--add-data", "materials;materials"
], check=True)

# 2. release フォルダ作り直し
if os.path.exists(release_dir):
    shutil.rmtree(release_dir)

os.makedirs(release_dir)

# 3. exeコピー
shutil.copy(exe_path, os.path.join(release_dir, "main.exe"))

# 4. credentialsコピー
shutil.copy(cred_path, os.path.join(release_dir, "credentials.json"))

# 5. zip作成
zip_name = f"{APP_NAME}.zip"

with zipfile.ZipFile(zip_name, "w", zipfile.ZIP_DEFLATED) as z:
    for file in os.listdir(release_dir):
        z.write(os.path.join(release_dir, file), file)

print("Release zip created:", zip_name)