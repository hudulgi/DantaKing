import os
import shutil
from dt_config import *
import datetime
import time

# 일간 타겟리스트 공유 받은 폴더
receive_path = "C:\\Users\\jeong\\OneDrive\\단타킹\\dt_king\\daily_target"

now = datetime.datetime.now()
file_name = "target_list_%s.csv" % now.strftime("%y%m%d")
print(file_name)

target_file = os.path.join(receive_path, file_name)
dest_file = os.path.join(target_path, file_name)

while True:
    # 해당 일의 타켓 파일이 존재할 경우 공유폴더 -> 설정파일의 target_path로 파일 복사
    if os.path.exists(target_file):
        print("파일 확인")
        shutil.copy(target_file, dest_file)
        print(f"파일 복사: {target_file} -> {dest_file}")
        break
    time.sleep(1)
