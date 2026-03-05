from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from jinja2 import Environment, FileSystemLoader

from googleapiclient.http import MediaIoBaseDownload, HttpError
import io
import requests, os
from datetime import datetime, timezone, timedelta

import shutil
from tqdm import tqdm


jst_today = datetime.now().astimezone(timezone(timedelta(hours=9)))
jst_today_str = jst_today.strftime("%Y%m%d%H%M%S")

# base_dir = f"classroomArchive/archive_{jst_today_str}"
base_dir = f"classroomArchive/archive_20260305180358"
print(f"保存先: {base_dir}")

os.makedirs(f"{base_dir}", exist_ok=True)
os.makedirs(f"{base_dir}/driveFiles", exist_ok=True)
os.makedirs(f"{base_dir}/css", exist_ok=True)
os.makedirs(f"{base_dir}/img", exist_ok=True)
os.makedirs(f"{base_dir}/img/icons", exist_ok=True)
shutil.copy('materials/style.css', f"{base_dir}/css/style.css")
shutil.copy('materials/assignment.svg', f"{base_dir}/img/assignment.svg")
shutil.copy('materials/book.svg', f"{base_dir}/img/book.svg")


SCOPES = [
    "https://www.googleapis.com/auth/classroom.courses.readonly",
    "https://www.googleapis.com/auth/classroom.announcements.readonly",
    "https://www.googleapis.com/auth/classroom.coursework.me",
    "https://www.googleapis.com/auth/classroom.courseworkmaterials.readonly",
    "https://www.googleapis.com/auth/classroom.student-submissions.students.readonly",
    "https://www.googleapis.com/auth/classroom.rosters.readonly",
    "https://www.googleapis.com/auth/classroom.profile.photos",
    "https://www.googleapis.com/auth/classroom.addons.student",
    "https://www.googleapis.com/auth/classroom.topics.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
creds = flow.run_local_server(port=0)
service = build("classroom", "v1", credentials=creds)

env = Environment(loader=FileSystemLoader("materials"))
template = env.get_template("course.html")

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

# for course in courses.get("courses", []):
course = courses[4]
print(f"コース情報: {course}")

announcements = list_all(
    lambda **kwargs: service.courses().announcements().list(courseId=course["id"], **kwargs),
    "announcements"
)
course_work = list_all(
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

user_profiles = {}

for teacher in tqdm(teachers, desc="教師アイコンを取得中"):
    if teacher["userId"] in user_profiles:
        continue
    profile = teacher["profile"]
    user_profiles[teacher["userId"]] = profile
    if "photoUrl" in profile:
        path = f"{base_dir}/img/icons/{profile["id"]}.png"
        if os.path.exists(path):
            print(f"Skip (already exists): {path}.png;")
        else:
            r = requests.get(f"https:{profile["photoUrl"]}", )
            if r.status_code == 200:
                with open(path, "wb") as f:
                    f.write(r.content)
                print(f"Saved teacher icon: {path}.png;")
            else:
                print(f"Failed to save teacher icon: {path}.png; status_code: {r.status_code};")

drive_files_to_download = []

def get_jst_str(iso_str):
    utc_dt = datetime.fromisoformat(iso_str.replace('Z', '+00:00'))
    jst_timezone = timezone(timedelta(hours=9))
    jst_dt = utc_dt.astimezone(jst_timezone)
    jst_dt_str = f"{jst_dt.year}年{jst_dt.month}月{jst_dt.day}日 {jst_dt.hour}時{jst_dt.minute}分"
    return jst_dt_str


# driveFile download
drive_service = build("drive", "v3", credentials=creds)

def download_drive_file(file_id, filename):
    path = f"{base_dir}/driveFiles/id_{file_id}_name_{filename}"

    if os.path.exists(path):
        print(f"Skip (already exists): {filename}")
        return
    
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.FileIO(path, "wb")
    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while not done:
        try:
            status, done = downloader.next_chunk()
            print(f"filename: {filename}; file_id: {file_id}; progress: {int(status.progress() * 100)}%; done: {done}")
        except HttpError as error:
            print(f"Failed to download file; filename: {filename}; file_id: {file_id};")
            print(error)
            done = True

# 授業のトピック
topic_map = {
    topic["topicId"]: topic
    for topic in topics
}

for item in list(announcements + course_work + course_work_materials):
    item["creatorUserProfile"] = user_profiles[item["creatorUserId"]]
    item["creationTime"] = get_jst_str(item["creationTime"])
    item["updateTime"] = get_jst_str(item["updateTime"])
    item["was_updated"] = item["creationTime"] != item["updateTime"]
    if "topicId" in item:
        item["topicName"] = topic_map[item["topicId"]]["name"]

    if "materials" in item:
        for material in item["materials"]:
            if "driveFile" in material and "title" in material["driveFile"]["driveFile"]:
                file_id = material["driveFile"]["driveFile"]["id"]
                file_name = material["driveFile"]["driveFile"]["title"]
                drive_files_to_download.append((file_id, file_name))

# 提出物（課題の添付ファイル）
submission_map = {
    s["courseWorkId"]: s
    for s in submissions
}

for item in course_work:
    submission = submission_map.get(item["id"])
    if not submission:
        continue

    assignmentSubmission = submission.get("assignmentSubmission")
    if not assignmentSubmission:
        continue

    attachments = assignmentSubmission.get("attachments", [])

    item["attachments"] = attachments
    for attachment in attachments:
        if "driveFile" in attachment:
            # Materialとのズレを修正するため
            # テンプレートではMaterialと同様に扱う
            attachment["driveFile"]["driveFile"] = attachment["driveFile"]
            if "title" in attachment["driveFile"]:
                file_id = attachment["driveFile"]["id"]
                file_name = attachment["driveFile"]["title"]
                drive_files_to_download.append((file_id, file_name))

for item in announcements:
    item["item_type"] = "Announcement"
for item in course_work:
    item["item_type"] = "CourseWork"
for item in course_work_materials:
    item["item_type"] = "CourseWorkMaterial"

all_items = announcements + course_work + course_work_materials
all_items.sort(key=lambda item: item['updateTime'], reverse=True)

html = template.render(
    name=course["name"],
    section=course.get("section", ""),
    announcements=announcements,
    course_work=course_work,
    course_work_materials=course_work_materials,
    all_items=all_items
)

with open(f"{base_dir}/クラス_{course["name"]}.html", "w", encoding="utf-8") as f:
    f.write(html)

from concurrent.futures import ThreadPoolExecutor

for item in tqdm(drive_files_to_download, desc="ドライブファイルをダウンロード中"):
    with ThreadPoolExecutor(max_workers=8) as ex:
        ex.map(lambda file: download_drive_file(file[0], file[1]), drive_files_to_download)

print(f"完了しました。アーカイブは {base_dir} に出力されています。")