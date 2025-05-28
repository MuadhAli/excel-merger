import base64

def encode_file_to_base64(file_path):
    with open(file_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

print(encode_file_to_base64("D:\Excel Merger\Excel_File_2.xlsx"))
