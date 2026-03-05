from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from jinja2 import Environment, FileSystemLoader

from googleapiclient.http import MediaIoBaseDownload
import io
import requests
from datetime import datetime, timezone, timedelta

SCOPES = [
    "https://www.googleapis.com/auth/classroom.courses.readonly",
    "https://www.googleapis.com/auth/classroom.announcements.readonly",
    "https://www.googleapis.com/auth/classroom.coursework.me",
    "https://www.googleapis.com/auth/classroom.courseworkmaterials.readonly",
    "https://www.googleapis.com/auth/classroom.student-submissions.students.readonly",
    "https://www.googleapis.com/auth/classroom.rosters.readonly",
    "https://www.googleapis.com/auth/classroom.profile.photos",
    "https://www.googleapis.com/auth/drive.readonly",
]

flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
creds = flow.run_local_server(port=0)
service = build("classroom", "v1", credentials=creds)
courses = service.courses().list().execute()

env = Environment(loader=FileSystemLoader("templates"))
template = env.get_template("course.html")

user_profiles = {}

# for course in courses.get("courses", []):
course = courses.get("courses", [])[4]
print(f"コース情報: {course}")


announcements = service.courses().announcements().list(courseId=course["id"]).execute().get("announcements", [])
course_work = service.courses().courseWork().list(courseId=course["id"]).execute().get("courseWork", [])
course_work_materials = service.courses().courseWorkMaterials().list(courseId=course["id"]).execute().get("courseWorkMaterial", [])
teachers = service.courses().teachers().list(courseId=course["id"]).execute().get("teachers", [])

for teacher in teachers:
    if teacher["userId"] in user_profiles:
        continue
    profile = teacher["profile"]
    user_profiles[teacher["userId"]] = profile
    # if "photoUrl" in profile:
    #     r = requests.get(f"https:{profile["photoUrl"]}", )
    #     if r.status_code == 200:
    #         with open(f"output/icons/{profile["id"]}.png", "wb") as f:
    #             f.write(r.content)
    #         print("Saved teacher icon:", profile["id"])
    #     else:
    #         print("Failed to save teacher icon:", profile["id"], "status_code:", r.status_code)

# with open(f"output/example.json", "w", encoding="utf-8") as f:
#     f.write(announcements.__str__())

def get_jst_str(iso_str):
    utc_dt = datetime.fromisoformat(iso_str.replace('Z', '+00:00'))
    jst_timezone = timezone(timedelta(hours=9))
    jst_dt = utc_dt.astimezone(jst_timezone)
    jst_dt_str = f"{jst_dt.year}年{jst_dt.month}月{jst_dt.day}日 {jst_dt.hour}時{jst_dt.minute}分"
    return jst_dt_str


for item in list(announcements + course_work + course_work_materials):
    item["creatorUserProfile"] = user_profiles[item["creatorUserId"]]
    item["creationTime"] = get_jst_str(item["creationTime"])
    item["updateTime"] = get_jst_str(item["updateTime"])
    if item["creationTime"] == item["updateTime"]:
        item["updateTime"] = None
    

html = template.render(
    name=course["name"],
    section=course.get("section", ""),
    announcements=announcements,
    course_work=course_work,
    course_work_materials=course_work_materials,
)

# driveFile download
drive_service = build("drive", "v3", credentials=creds)

def download_file(file_id, filename):
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.FileIO(f"output/driveFiles/id_{file_id}_name_{filename}", "wb")
    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()
        print(f"filename: {filename}; file_id: {file_id}; progress: {int(status.progress() * 100)}%; done: {done}")

# for annoucement in announcements.get("announcements", []):
#     if "materials" in annoucement:
#         for material in annoucement["materials"]:
#             if "driveFile" in material:
#                 file_id = material["driveFile"]["driveFile"]["id"]
#                 file_name = material["driveFile"]["driveFile"]["title"]
#                 download_file(file_id, file_name)

with open(f"output/クラス_{course["name"]}.html", "w", encoding="utf-8") as f:
    f.write(html)