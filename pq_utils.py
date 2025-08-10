import zipfile

def extract_m_code_from_pq(file_path_or_obj):
    extracted = []
    with zipfile.ZipFile(file_path_or_obj, 'r') as archive:
        for name in archive.namelist():
            if name.endswith(".m"):
                with archive.open(name) as f:
                    m_script = f.read().decode("utf-8")
                    extracted.append((name.split("/")[-1].replace(".m", ""), m_script))
    return extracted