import os
import shutil
import datetime
from tkinter import Tk, filedialog, Button, Label, Text, END, BOTH, messagebox, Listbox, SINGLE, Toplevel
from zipfile import ZipFile
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
import io
from google.auth.transport.requests import Request  # これを追加

# Google Drive APIスコープ
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# credentials.json ファイルのフルパスを指定
CREDENTIALS_PATH = r"C:\Users\hp\OneDrive\ドキュメント\4年次データ\ソフトウェア工学\PythonApplication1\PythonApplication1\credentials.json"

# パスが正しいか確認するためのデバッグ情報を追加
print(f"Looking for credentials file at: {CREDENTIALS_PATH}")
print(f"File exists: {os.path.exists(CREDENTIALS_PATH)}")

def authenticate_google_drive():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            print(f"Using credentials file: {CREDENTIALS_PATH}")
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def upload_to_drive(file_path, creds, backup_name):
    service = build('drive', 'v3', credentials=creds)
    file_metadata = {'name': backup_name}
    media = MediaFileUpload(file_path, mimetype='application/zip')
    file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    print(f'File ID: {file.get("id")}')
    return file.get("id")

def list_drive_files(creds):
    service = build('drive', 'v3', credentials=creds)
    results = service.files().list(pageSize=10, fields="files(id, name)").execute()
    items = results.get('files', [])
    return items

def download_from_drive(file_id, file_name, creds):
    service = build('drive', 'v3', credentials=creds)
    request = service.files().get_media(fileId=file_id)
    fh = io.FileIO(file_name, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print(f"Download {int(status.progress() * 100)}%.")
    fh.close()

# ファイル拡張子と対応するフォルダ名の辞書
extensions_folders = {
    'doc': 'Word Documents',
    'docx': 'Word Documents',
    'xls': 'Excel Files',
    'xlsx': 'Excel Files',
    'ppt': 'PowerPoint Presentations',
    'pptx': 'PowerPoint Presentations',
    'pdf': 'PDF Files',
    'jpg': 'Images',
    'jpeg': 'Images',
    'png': 'Images',
    'gif': 'Images',
    'mp4': 'Videos',
    'avi': 'Videos',
    'mov': 'Videos',
    'txt': 'Text Files',
    'csv': 'CSV Files',
}

def organize_and_zip_files(source_folder, zip_output_path):
    organized_paths = []
    with ZipFile(zip_output_path, 'w') as zipf:
        for file_name in os.listdir(source_folder):
            file_path = os.path.join(source_folder, file_name)
            if os.path.isfile(file_path):
                file_extension = file_name.split('.')[-1].lower()
                if file_extension in extensions_folders:
                    folder_name = extensions_folders[file_extension]
                    arcname = os.path.join(folder_name, file_name)
                    zipf.write(file_path, arcname)
                    organized_paths.append(f'{file_path} を {folder_name} にコピーしました')
    return organized_paths

def gui():
    root = Tk()
    root.title("ファイル整理アプリ")
    root.geometry("600x400")

    creds = authenticate_google_drive()

    def folder_organize():
        source_folder = filedialog.askdirectory(title="フォルダを選択してください")
        if source_folder:
            backup_name = filedialog.asksaveasfilename(title="名前を付けて保存", defaultextension=".zip", filetypes=[("ZIP files", "*.zip")])
            if backup_name:
                date_str = datetime.datetime.now().strftime("%Y%m%d")
                zip_output_path = f'{backup_name}_{date_str}.zip'
                
                organized_paths = organize_and_zip_files(source_folder, zip_output_path)
                
                output_message = "\n".join(organized_paths)
                output_text.delete('1.0', END)
                output_text.insert(END, output_message)
                messagebox.showinfo("完了", "フォルダを整理して圧縮しました")

    def backup():
        source_folder = filedialog.askdirectory(title="バックアップするフォルダを選択してください")
        if source_folder:
            backup_name = filedialog.asksaveasfilename(title="バックアップの名前を作成", defaultextension=".zip", filetypes=[("ZIP files", "*.zip")])
            if backup_name:
                date_str = datetime.datetime.now().strftime("%Y%m%d")
                zip_output_path = f'{backup_name}_{date_str}.zip'
                
                organized_paths = organize_and_zip_files(source_folder, zip_output_path)
                
                drive_backup_name = f'{backup_name}_{date_str}.zip'
                file_id = upload_to_drive(zip_output_path, creds, drive_backup_name)
                
                output_text.delete('1.0', END)
                output_text.insert(END, f'\n'.join(organized_paths) + f'\nFile ID: {file_id}')
                messagebox.showinfo("完了", "バックアップが完了しました")

    def download_backup():
        items = list_drive_files(creds)
        if not items:
            messagebox.showinfo("情報", "Google Driveにバックアップファイルがありません")
            return

        def on_select(evt):
            w = evt.widget
            index = int(w.curselection()[0])
            value = w.get(index)
            file_id = value.split(' ')[-1]
            download_name = filedialog.asksaveasfilename(title="ダウンロード先を選択してください", defaultextension=".zip", filetypes=[("ZIP files", "*.zip")])
            if download_name:
                download_from_drive(file_id, download_name, creds)
                messagebox.showinfo("完了", "ダウンロードが完了しました")

        top = Toplevel(root)
        top.title("バックアップファイルを選択")
        lb = Listbox(top, selectmode=SINGLE, width=80, height=20)
        lb.pack(pady=20)

        for item in items:
            lb.insert(END, f'{item["name"]} {item["id"]}')

        lb.bind('<<ListboxSelect>>', on_select)

    Label(root, text="ファイル整理アプリ", font=("Arial", 20)).pack(pady=20)
    Button(root, text="フォルダ整理", command=folder_organize, width=20, height=2).pack(pady=10)
    Button(root, text="バックアップ", command=backup, width=20, height=2).pack(pady=10)
    Button(root, text="バックアップフォルダのダウンロード", command=download_backup, width=30, height=2).pack(pady=10)
    
    output_text = Text(root, wrap='word', height=10)
    output_text.pack(expand=True, fill=BOTH)
    
    root.mainloop()

if __name__ == '__main__':
    gui()
