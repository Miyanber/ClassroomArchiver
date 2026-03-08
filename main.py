from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.http import MediaIoBaseDownload, HttpError
from googleapiclient.discovery import build
from jinja2 import Environment, FileSystemLoader
from datetime import datetime, timezone, timedelta
from tqdm import tqdm
import io, requests, os, sys
import signal
import shutil
import mimetypes
from concurrent.futures import ThreadPoolExecutor, as_completed

signal.signal(signal.SIGINT, lambda sig, frame: stop_event.set())


jst_today = datetime.now().astimezone(timezone(timedelta(hours=9)))
jst_today_str = jst_today.strftime("%Y%m%d%H%M%S")

base_dir = f"classroomArchive/archive_{jst_today_str}"

# ロギングの設定
import logging
from colorama import init, Fore, Style

# coloramaの初期化（WindowsのANSIエスケープシーケンス対応）
init(autoreset=True)

log_dir = os.path.join(base_dir, "logs")
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f"archive_{jst_today_str}.log")

logger = logging.getLogger("ClassroomArchiver")
logger.setLevel(logging.DEBUG)

file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
file_handler.setFormatter(file_formatter)

console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

logger.addHandler(file_handler)
logger.addHandler(console_handler)

def log_info(msg):
    logger.info(msg)

def log_error(msg):
    logger.error(f"{Fore.RED}{msg}{Style.RESET_ALL}", exc_info=True)

def log_warning(msg):
    logger.warning(f"{Fore.YELLOW}{msg}{Style.RESET_ALL}")

def log_debug(msg, exc_info=False):
    logger.debug(msg, exc_info=exc_info)


log_info(f"保存先: {base_dir}")

def resource_path(path):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, path)
    return os.path.join(os.path.abspath("."), path)

materials_dir = resource_path("materials")

os.makedirs(f"{base_dir}", exist_ok=True)
os.makedirs(f"{base_dir}/driveFiles", exist_ok=True)
os.makedirs(f"{base_dir}/css", exist_ok=True)
os.makedirs(f"{base_dir}/img", exist_ok=True)
os.makedirs(f"{base_dir}/img/icons", exist_ok=True)
shutil.copy(os.path.join(materials_dir, "style.css"), f"{base_dir}/css/style.css")
shutil.copy(os.path.join(materials_dir, "assignment.svg"), f"{base_dir}/img/assignment.svg")
shutil.copy(os.path.join(materials_dir, "book.svg"), f"{base_dir}/img/book.svg")
shutil.copy(os.path.join(materials_dir, "user.svg"), f"{base_dir}/img/user.svg")


SCOPES = [
    "https://www.googleapis.com/auth/classroom.courses.readonly",
    "https://www.googleapis.com/auth/classroom.announcements.readonly",
    "https://www.googleapis.com/auth/classroom.coursework.me",
    "https://www.googleapis.com/auth/classroom.courseworkmaterials.readonly",
    "https://www.googleapis.com/auth/classroom.student-submissions.students.readonly",
    "https://www.googleapis.com/auth/classroom.rosters.readonly",
    "https://www.googleapis.com/auth/classroom.profile.photos",
    "https://www.googleapis.com/auth/classroom.topics.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/drive.file",
]

flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
creds = flow.run_local_server(port=0)
service = build("classroom", "v1", credentials=creds)

env = Environment(loader=FileSystemLoader(materials_dir))
template = env.get_template("course.html")

archive_folder_id = None

try:
    file_name = "archive_folder_id.txt"
    drive_service = build("drive", "v3", credentials=creds)

    folder_name = "Classroom Archive"

    query = (
        f"name = '{folder_name}' "
        f"and 'root' in parents "
        f"and mimeType = 'application/vnd.google-apps.folder' "
        f"and trashed = false"
    )

    results = drive_service.files().list(
        q=query,
        spaces='drive',
        fields='files(id, name)', # 必要なフィールドだけ取得
        pageSize=10
    ).execute()
    
    items = results.get('files', [])
    root_folder_id = None

    if not items:
        file_metadata = {
            "name": "Classroom Archive",
            "mimeType": "application/vnd.google-apps.folder",
        }
        file = drive_service.files().create(body=file_metadata, fields="id").execute()
        root_folder_id = file.get("id")
    else:
        # 複数ヒットする可能性があるため、最初の一つを返す
        folder = items[0]
        root_folder_id = folder["id"]

    # 個別フォルダ作成
    file_metadata = {
        "name": jst_today_str,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [root_folder_id],
    }
    file = drive_service.files().create(body=file_metadata, fields="id").execute()
    archive_folder_id = file.get("id")
    
except HttpError as error:
    log_error(f"An error occurred: {error}")
    log_error("プログラムを終了します。詳細はログファイルを確認してください。")
    sys.exit(1)


def format_size(size):
    size = int(size)
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size < 1024:
            return f"{size:.2f} {unit}"
        size /= 1024


def list_all(method, key):
    items = []
    page_token = None
    
    while True:
        result = method(pageToken=page_token).execute()
        items.extend(result.get(key, []))
        page_token = result.get("nextPageToken")

        if not page_token:
            break

    return items


courses = list_all(
    lambda **kwargs: service.courses().list(**kwargs),
    "courses"
)

import threading
thread_local = threading.local()
import re

stop_event = threading.Event()

user_profiles = {}
pictures_to_download = set()
all_drive_files_to_download = set()
all_drive_files_to_copy = set()
all_files_to_download_size = 0
file_cache = {}

# 1GBの閾値 (バイト単位)
THRESHOLD_GB = 1 * 1024 * 1024 * 1024


def get_jst_str(iso_str):
    utc_dt = datetime.fromisoformat(iso_str.replace('Z', '+00:00'))
    jst_timezone = timezone(timedelta(hours=9))
    jst_dt = utc_dt.astimezone(jst_timezone)
    jst_dt_str = f"{jst_dt.year}年{jst_dt.month}月{jst_dt.day}日 {jst_dt.hour}時{jst_dt.minute}分"
    return jst_dt_str

def get_file_type_name(mime_type):
    if mime_type == None:
        return None
    elif mime_type == "application/vnd.google-apps.document":
        return "Google ドキュメント"
    elif mime_type == "application/vnd.google-apps.presentation":
        return "Google スライド"
    elif mime_type == "application/vnd.google-apps.spreadsheet":
        return "Google スプレッドシート"
    else:
        return None

def get_download_file_path(id, name):
    return f"{base_dir}/driveFiles/id_{id}_name_{name}"

# (dict | None) を返す。
# None の場合はダウンロード・コピー共に不可
# ダウンロード可能なファイルのみリストに追加し、それ以外はドライブにコピーする。
def fetch_drive_file_details(drive_file):
    file_name = drive_file["title"]
    file_id = drive_file["id"]

    # 強制終了用
    if stop_event.is_set():
        log_warning(f"Cancelled: {file_name}")
        return None

    if not hasattr(thread_local, "drive_service"):
        thread_local.drive_service = build("drive", "v3", credentials=creds)

    drive_service = thread_local.drive_service

    # 利用不可の文字を消す
    file_name = re.sub(r'[\\/*?:"<>|]', "_", file_name)
    file_name = file_name.replace("\n", " ")
    file_name = file_name.rstrip(" .")

    file_type = None
    path = get_download_file_path(file_id, file_name)

    if os.path.exists(path):
        log_info(f"Skip (already exists): {file_name}")
        mime_type = mimetypes.guess_file_type(file_name)[0]
        if mime_type:
            drive_extension = mimetypes.guess_extension(mime_type)
            file_type = drive_extension.upper()[1:]
        return {
            "file_name": file_name,
            "file_type": file_type,
            "save_type": "download",
            "size": 0,
        }
    
    if file_id in file_cache:
        file = file_cache[file_id]
    else:
        try:
            # 仮に404ならここでエラーが出る
            file = drive_service.files().get(
                fileId=file_id,
                fields="name,mimeType,size,capabilities",
                supportsAllDrives=True,
            ).execute()
            file_cache[file_id] = file
            if not "size" in file:
                file["size"] = 0
        except HttpError as e:
            try:
                # 課題等で稀にClassroomが返してるIDとDriveの実ファイルIDが別物になっている場合がある
                m = re.search(r'/d/([a-zA-Z0-9_-]+)', drive_file["alternateLink"])
                file_id = m.group(1) if m else None
                if file_id:
                    file = drive_service.files().get(
                        fileId=file_id,
                        fields="name,mimeType,size,capabilities",
                        supportsAllDrives=True,
                    ).execute()
                    file_cache[file_id] = file
                    if not "size" in file:
                        file["size"] = 0
                else:
                    log_warning(f"ファイル（{file_name}）の情報が取得できなかったためスキップします。ステータスコード: {e.status_code}")
                    log_debug(f"Failed to get file information; drive_file: {drive_file}; error: {e}", exc_info=True)
                    return None
            except HttpError as e:
                log_warning(f"ファイル（{file_name}）の情報が取得できなかったためスキップします。ステータスコード: {e.status_code}")
                log_debug(f"Failed to get file information; drive_file: {drive_file}; error: {e}", exc_info=True)
                return None
        
    mime_type = file["mimeType"]
    size = int(file["size"])

    drive_extension = mimetypes.guess_extension(mime_type)
    if drive_extension:
        file_type = f"{drive_extension.upper()[1:]} ファイル"
    else:
        file_type = get_file_type_name(mime_type)

    _, name_extension = os.path.splitext(file_name)

    # 拡張子が必要な場合は付与
    if drive_extension and name_extension != drive_extension and not mime_type.startswith("application/vnd.google-apps"):
        file_name += drive_extension

    if mime_type == "application/vnd.google-apps.folder":
        log_info(f"フォルダはダウンロード・コピーできません。フォルダ名: {file_name}, リンク: {drive_file['alternateLink']}")
        return None

    elif mime_type.startswith("application/vnd.google-apps.") and file["capabilities"]["canCopy"]:
        return {
            "file_name": file_name,
            "file_type": file_type,
            "save_type": "copy",
            "size": 0,
        }
    
    elif file["capabilities"]["canDownload"]:
        return {
            "file_name": file_name,
            "file_type": file_type,
            "save_type": "download",
            "size": size,
        }
    else:
        if not file["capabilities"]["canDownload"]:
            log_info(f"ファイルのダウンロードが許可されていません。ファイル名: {file_name}, リンク: {drive_file['alternateLink']}")
        elif not file["capabilities"]["canCopy"]:
            log_info(f"ファイルのコピーが許可されていません。ファイル名: {file_name}, リンク: {drive_file['alternateLink']}")
        else:
            log_info(f"ダウンロード・コピー両方できないファイル形式です。ファイル名: {file_name}, リンク: {drive_file['alternateLink']}")
            log_debug(f"Failed to get file information; drive_file: {drive_file};", exc_info=True)
        return None


def download_file(url, path):
    # 強制終了用
    if stop_event.is_set():
        log_warning(f"Cancelled: {path}")
        return 
    
    r = requests.get(url, )
    if r.status_code == 200:
        with open(path, "wb") as f:
            f.write(r.content)
    else:
        log_warning(f"Failed to save {path}.png; status_code: {r.status_code};")


# Google ファイル以外のダウンロード
def download_drive_file(file_id, file_name):
    # 強制終了用
    if stop_event.is_set():
        log_warning(f"Cancelled: {file_name}")
        return None

    if not hasattr(thread_local, "drive_service"):
        thread_local.drive_service = build("drive", "v3", credentials=creds)

    drive_service = thread_local.drive_service
    
    path = get_download_file_path(file_id, file_name)

    try:
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.FileIO(path, "wb")
        downloader = MediaIoBaseDownload(fh, request)

        done = False
        while not done:
            status, done = downloader.next_chunk()
            # log_info(f"filename: {file_name}; file_id: {file_id}; progress: {int(status.progress() * 100)}%; done: {done}")
    except HttpError as error:
        log_warning(f"Failed to download file; filename: {file_name}; file_id: {file_id}; error: {error}")
        done = True


def copy_drive_file(file_id, file_name):
    # 強制終了用
    if stop_event.is_set():
        log_warning(f"Cancelled: {file_name}")
        return None

    if not hasattr(thread_local, "drive_service"):
        thread_local.drive_service = build("drive", "v3", credentials=creds)

    drive_service = thread_local.drive_service

    try:
        copied_file = drive_service.files().copy(
            fileId=file_id,
            body={
                "name": file_name,
                "parents": [archive_folder_id] # Apps Script (.gs) は親フォルダ指定無視でドライブ直下に保存される
            },
            fields="id,name,webViewLink,mimeType"
        ).execute()
        log_info(f"ダウンロードできないファイル形式のため、マイドライブにコピーが作成されました。ファイル名: {copied_file["name"]}, リンク: {copied_file['webViewLink']}")
    except HttpError as error:
        log_warning(f"Failed to copy file; filename: {file_name}; file_id: {file_id}; error: {error}")
        done = True

courses_to_archive = []

# コース別の処理
for course in courses:
    # 強制終了用
    if stop_event.is_set():
        log_warning(f"Cancelled: {course}")
        exit()
    
    drive_files_to_copy = set()
    drive_files_to_download = set()
    files_to_download_size = 0

    announcements = list_all(
        lambda **kwargs: service.courses().announcements().list(courseId=course["id"], **kwargs),
        "announcements"
    )
    course_works = list_all(
        lambda **kwargs: service.courses().courseWork().list(courseId=course["id"], **kwargs),
        "courseWork"
    )
    course_work_materials = list_all(
        lambda **kwargs: service.courses().courseWorkMaterials().list(courseId=course["id"], **kwargs),
        "courseWorkMaterial"
    )
    teachers = list_all(
        lambda **kwargs: service.courses().teachers().list(courseId=course["id"], **kwargs),
        "teachers"
    )
    students = list_all(
        lambda **kwargs: service.courses().students().list(courseId=course["id"], **kwargs),
        "students"
    )
    topics = list_all(
        lambda **kwargs: service.courses().topics().list(courseId=course["id"], **kwargs),
        "topic"
    )
    submissions = list_all(
        lambda **kwargs: service.courses().courseWork().studentSubmissions().list(
            courseId=course["id"],
            courseWorkId="-",
            userId="me",
            **kwargs
        ),
        "studentSubmissions"
    )

    # 授業のトピック
    topic_map = {
        topic["topicId"]: topic
        for topic in topics
    }
        
    # 提出物（課題の添付ファイル）
    submission_map = {
        s["courseWorkId"]: s
        for s in submissions
    }

    # CourseWork の個別提出物・返却物の取得
    def get_course_work_attachments(course_work):
        global files_to_download_size, drive_files_to_copy, drive_files_to_download

        submission = submission_map.get(course_work["id"])
        if not submission:
            return

        assignmentSubmission = submission.get("assignmentSubmission")
        if not assignmentSubmission:
            return

        attachments = assignmentSubmission.get("attachments", [])

        course_work["attachments"] = attachments
        for attachment in attachments:
            if "driveFile" in attachment:
                # Materialとのズレを修正するため
                # テンプレートではMaterialと同様に扱う
                attachment["driveFile"]["driveFile"] = attachment["driveFile"]
                drive_file = attachment["driveFile"]["driveFile"]
                if "title" in drive_file:
                    drive_file = attachment["driveFile"]["driveFile"]
                    file_detail = fetch_drive_file_details(drive_file)
                    if file_detail:
                        drive_file["title"] = file_detail["file_name"] # 拡張子補完等のため必要
                        drive_file["file_type"] = file_detail["file_type"]
                        drive_file["save_type"] = file_detail["save_type"]
                        drive_file["size"] = file_detail["size"]
                        files_to_download_size += file_detail["size"]
                        if drive_file["save_type"] == "copy":
                            drive_files_to_copy.add((drive_file["id"], drive_file["title"]))
                        elif drive_file["save_type"] == "download":
                            drive_files_to_download.add((drive_file["id"], drive_file["title"]))
                        else:
                            log_warning(f"Unsupported save type. DriveFile: {drive_file}")

    # 投稿のObjectにフィールドを追加する
    def clean_item(item):
        if item["creatorUserId"] in user_profiles:
            item["creatorUserProfile"] = user_profiles[item["creatorUserId"]]
        else:
            # エラー防止でダミーオブジェクト挿入
            item["creatorUserProfile"] = {
                "id": item["creatorUserId"],
                "name": {
                    "givenName": "不明ユーザー",
                    "familyName": "",
                    "fullName": "不明ユーザー"
                },
                "photoUrl": None,
            }
        item["creationTime"] = get_jst_str(item["creationTime"])
        item["updateTime"] = get_jst_str(item["updateTime"])
        item["was_updated"] = item["creationTime"] != item["updateTime"]
        if "topicId" in item:
            item["topicName"] = topic_map[item["topicId"]]["name"]

    # 投稿の添付資料の取得
    def get_all_materials(item):
        global files_to_download_size, drive_files_to_copy, drive_files_to_download

        if "materials" in item:
            for material in item["materials"]:
                if "driveFile" in material and "title" in material["driveFile"]["driveFile"]:
                    drive_file = material["driveFile"]["driveFile"]
                    file_detail = fetch_drive_file_details(drive_file)
                    if file_detail:
                        drive_file["title"] = file_detail["file_name"] # 拡張子補完等のため必要
                        drive_file["file_type"] = file_detail["file_type"]
                        drive_file["save_type"] = file_detail["save_type"]
                        drive_file["size"] = file_detail["size"]
                        files_to_download_size += file_detail["size"]
                        if drive_file["save_type"] == "copy":
                            drive_files_to_copy.add((drive_file["id"], drive_file["title"]))
                        elif drive_file["save_type"] == "download":
                            drive_files_to_download.add((drive_file["id"], drive_file["title"]))
                        else:
                            log_warning(f"Unsupported save type. DriveFile: {drive_file}")

    for item in announcements:
        item["item_type"] = "Announcement"
    for item in course_works:
        item["item_type"] = "CourseWork"
    for item in course_work_materials:
        item["item_type"] = "CourseWorkMaterial"


    all_items = announcements + course_works + course_work_materials
    # get_jst_str(item["creationTime"]) で変換する前に実行する必要あり
    all_items.sort(key=lambda item: item['updateTime'], reverse=True)

    try:
        with ThreadPoolExecutor(max_workers=8) as executor:
            # 先に全タスクをsubmitし、futureオブジェクトをリスト化する
            futures = [executor.submit(get_course_work_attachments, item) for item in course_works]
            futures += [executor.submit(get_all_materials, item) for item in all_items]
            futures += [executor.submit(clean_item, item) for item in all_items]
            
            # as_completedで終わったものから取り出し、tqdmでラップする
            results = []
            for future in tqdm(as_completed(futures), total=len(futures), desc=f"{course["name"]}の情報を取得中"):
                results.append(future.result())

    except KeyboardInterrupt:
        stop_event.set()

    log_info("\n==============================")
    log_info(f"クラス名: {course["name"]}")
    log_info(f"投稿（お知らせ・課題・資料）の合計数: {len(all_items)}")
    log_info(f"ドライブへコピー対象のファイル数: {len(drive_files_to_copy)}")
    log_info(f"ダウンロード対象ファイルの数: {len(drive_files_to_download)}")
    log_info(f"合計サイズ（ダウンロード対象のみ）: {format_size(files_to_download_size)}")
    log_info("==============================\n")

    if files_to_download_size > THRESHOLD_GB:
        log_warning(f"注意: ダウンロード対象ファイルの合計サイズが1GBを超えています ({format_size(files_to_download_size)})")
        confirm = input("このクラスをアーカイブ対象に含めますか？ (y/N): ").strip().lower()
        if confirm != "y":
            log_info(f"クラス「{course["name"]}」をアーカイブ対象から除外します。")
            continue
        else:
            log_info(f"クラス「{course["name"]}」をアーカイブ対象に登録します。")
            courses_to_archive.append(course)
    else:
        log_info("ダウンロード対象ファイルの合計サイズが1GB未満のため、自動的にアーカイブ対象に登録します。")
        courses_to_archive.append(course)
    
    for user in list(teachers + students):
        if user["userId"] in user_profiles:
            continue
        profile = user["profile"]
        user_profiles[user["userId"]] = profile
        if "photoUrl" in profile:
            path = f"{base_dir}/img/icons/{profile["id"]}.png"
            if os.path.exists(path):
                log_info(f"Skip (already exists): {path};")
            else:
                pictures_to_download.add((f"https:{profile["photoUrl"]}", path))

    all_drive_files_to_copy |= drive_files_to_copy
    all_drive_files_to_download |= drive_files_to_download
    all_files_to_download_size += files_to_download_size

    html = template.render(
        course=course,
        announcements=announcements,
        course_work=course_works,
        course_work_materials=course_work_materials,
        teachers=teachers,
        students=students,
        students_count=(len(students) + 1),
        all_items=all_items
    )

    with open(f"{base_dir}/クラス_{course["name"]}.html", "w", encoding="utf-8") as f:
        f.write(html)

log_info("アーカイブ対象のクラスが確定しました。")

log_info("\n==============================")
log_info("アーカイブ対象のクラス: ")
for course in courses_to_archive:
    log_info(f"- {course["name"]}")
log_info(f"計 {len(courses_to_archive)} クラス")
log_info(f"ドライブへコピー対象のファイルの合計数: {len(all_drive_files_to_copy)}")
log_info(f"ダウンロード対象ファイルの合計数: {len(all_drive_files_to_download)}")
log_info(f"ダウンロード対象ファイルの合計容量: {format_size(all_files_to_download_size)}")
log_info("==============================")

confirm = input("ファイルのダウンロードを開始しますか？ (y/N): ").strip().lower()

if confirm != "y":
    log_warning("処理を中止しました。")
    sys.exit(1)


try:
    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = [executor.submit(download_file, picture[0], picture[1]) for picture in pictures_to_download]
        futures += [executor.submit(copy_drive_file, file[0], file[1]) for file in drive_files_to_copy]
        futures += [executor.submit(download_drive_file, file[0], file[1]) for file in drive_files_to_download]
        
        results = []
        for future in tqdm(as_completed(futures), total=len(futures), desc=f"ファイルを保存中"):
            results.append(future.result())

except KeyboardInterrupt:
    stop_event.set()

log_info(f"完了しました。アーカイブは {base_dir} に出力されています。")