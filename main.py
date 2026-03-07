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
print(f"保存先: {base_dir}")

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
        print(f"フォルダ '{folder_name}' は見つかりませんでした。")
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
    print(f"An error occurred: {error}")
    exit(0)


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
drive_files_to_download = set()
file_cache = {}
files_to_download_size = 0


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
# ダウンロード可能なファイルのみリストに追加し、それ以外はドライブにコピーする。
def add_drive_file(file_id, file_name):
    global files_to_download_size
    # 強制終了用
    if stop_event.is_set():
        print(f"Cancelled: {file_name}")
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
        print(f"Skip (already exists): {file_name}")
        mime_type = mimetypes.guess_file_type(file_name)[0]
        if mime_type:
            drive_extension = mimetypes.guess_extension(mime_type)
            file_type = drive_extension.upper()[1:]
        return {
            "file_name": file_name,
            "file_type": file_type,
            "is_saved": True,
        }
    
    if file_id in file_cache:
        file = file_cache[file_id]
    else:
        try:
            # 仮に404ならここでエラーが出る
            file = drive_service.files().get(
                fileId=file_id,
                fields="name,mimeType,size,capabilities"
            ).execute()
            file_cache[file_id] = file
            if not "size" in file:
                file["size"] = 0
        except HttpError as e:
            print(f"Failed to get file information; filename: {file_name}; file_id: {file_id};")
            print(e)
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
        print(f"フォルダはダウンロードできません。ファイル名: {file_name}")
        return None

    elif mime_type.startswith("application/vnd.google-apps.") and file["capabilities"]["canCopy"]:
        copied_file = drive_service.files().copy(
            fileId=file_id,
            body={
                "name": file_name,
                "parents": [archive_folder_id] # Apps Script (.gs) は親フォルダ指定無視でドライブ直下に保存される
            },
            fields="id,name,webViewLink,mimeType"
        ).execute()
        print(f"ダウンロードできないファイル形式のため、マイドライブにコピーが作成されました。ファイル名: {copied_file["name"]}, リンク: {copied_file['webViewLink']}")

        return {
            "file_name": file_name,
            "file_type": file_type,
            "is_saved": False,
            "web_view_link": copied_file["webViewLink"],
        }
    
    elif file["capabilities"]["canDownload"]:
        drive_files_to_download.add((file_id, file_name))
        files_to_download_size += size

        return {
            "file_name": file_name,
            "file_type": file_type,
            "is_saved": True,
        }
    else:
        if not file["capabilities"]["canDownload"]:
            print(f"ファイルのダウンロードが許可されていません。ファイル名: {file_name}; ファイルID: {file_id};")
        elif not file["capabilities"]["canCopy"]:
            print(f"ファイルのコピーが許可されていません。ファイル名: {file_name}; ファイルID: {file_id};")
        else:
            print(f"ファイルがダウンロードできません。ファイル名: {file_name}; ファイルID: {file_id};")
        return None


def download_file(url, path):
    # 強制終了用
    if stop_event.is_set():
        print(f"Cancelled: {path}")
        return 
    
    r = requests.get(url, )
    if r.status_code == 200:
        with open(path, "wb") as f:
            f.write(r.content)
    else:
        print(f"Failed to save {path}.png; status_code: {r.status_code};")


# Google ファイル以外のダウンロード
def download_drive_file(file_id, file_name):
    # 強制終了用
    if stop_event.is_set():
        print(f"Cancelled: {file_name}")
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
            # print(f"filename: {file_name}; file_id: {file_id}; progress: {int(status.progress() * 100)}%; done: {done}")
    except HttpError as error:
        print(f"Failed to download file; filename: {file_name}; file_id: {file_id};")
        print(error)
        done = True


# コース別の処理
for course in courses:
    # 強制終了用
    if stop_event.is_set():
        print(f"Cancelled: {course}")
        exit()

    print(f"クラス名: {course["name"]}")

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

    for user in list(teachers + students):
        if user["userId"] in user_profiles:
            continue
        profile = user["profile"]
        user_profiles[user["userId"]] = profile
        if "photoUrl" in profile:
            path = f"{base_dir}/img/icons/{profile["id"]}.png"
            if os.path.exists(path):
                print(f"Skip (already exists): {path};")
            else:
                pictures_to_download.add((f"https:{profile["photoUrl"]}", path))

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

    def get_course_work_attachments(course_work):
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
                    file_dict = add_drive_file(drive_file["id"], drive_file["title"])
                    if file_dict:
                        drive_file["title"] = file_dict["file_name"] # 拡張子補完等のため必要
                        drive_file["file_type"] = file_dict["file_type"]
                        drive_file["is_saved"] = file_dict["is_saved"]
                        if "web_view_link" in file_dict:
                            drive_file["web_view_link"] = file_dict["web_view_link"]


    def get_all_materials(item):
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

        if "materials" in item:
            for material in item["materials"]:
                if "driveFile" in material and "title" in material["driveFile"]["driveFile"]:
                    drive_file = material["driveFile"]["driveFile"]
                    file_dict = add_drive_file(drive_file["id"], drive_file["title"])
                    if file_dict:
                        drive_file["title"] = file_dict["file_name"] # 拡張子補完等のため必要
                        drive_file["file_type"] = file_dict["file_type"]
                        drive_file["is_saved"] = file_dict["is_saved"]
                        if "web_view_link" in file_dict:
                            drive_file["web_view_link"] = file_dict["web_view_link"]


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
            
            # as_completedで終わったものから取り出し、tqdmでラップする
            results = []
            for future in tqdm(as_completed(futures), total=len(futures), desc=f"{course["name"]}のファイルを整理中"):
                results.append(future.result())

    except KeyboardInterrupt:
        stop_event.set()

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


def format_size(size):
    size = int(size)
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size < 1024:
            return f"{size:.2f} {unit}"
        size /= 1024

print("\n==============================")
print(f"ダウンロード予定ファイル数: {len(drive_files_to_download)}")
print(f"合計容量: {format_size(files_to_download_size)}")
print("==============================")

confirm = input("ダウンロードを開始しますか？ (y/N): ").strip().lower()

if confirm != "y":
    print("処理を中止しました。")
    exit()


try:
    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = [executor.submit(download_file, picture[0], picture[1]) for picture in pictures_to_download]
        futures += [executor.submit(download_drive_file, file[0], file[1]) for file in drive_files_to_download]
        
        results = []
        for future in tqdm(as_completed(futures), total=len(futures), desc=f"ファイルを保存中"):
            results.append(future.result())

except KeyboardInterrupt:
    stop_event.set()

print(f"完了しました。アーカイブは {base_dir} に出力されています。")