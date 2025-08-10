import zipfile
import os
import json
import tempfile

def save_project_zip(project_name, file_dict):
    zip_path = tempfile.NamedTemporaryFile(delete=False, suffix=".zip").name
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        meta = {
            "project_name": project_name,
            "files": list(file_dict.keys())
        }
        zipf.writestr("project.json", json.dumps(meta, indent=2))
        for fname, content in file_dict.items():
            zipf.writestr(fname, content)
    return zip_path

def load_project_zip(zip_file):
    files = {}
    meta = {}
    with zipfile.ZipFile(zip_file, 'r') as zipf:
        for name in zipf.namelist():
            with zipf.open(name) as f:
                content = f.read().decode("utf-8")
                if name == "project.json":
                    meta = json.loads(content)
                else:
                    files[name] = content
    return meta, files