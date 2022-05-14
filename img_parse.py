import os
from openpyxl import load_workbook
import shutil
from config import LOCAL_ROOT


workbook = load_workbook("data_new_2.xlsx", data_only=True)

char_sheet = workbook["캐릭터"]

tag = {
    "4.5": "",
    "NS": "_rank5",
    "AS": "_s2_rank5",
    "ES": "_s3_rank5"
}

if not os.path.exists("character"):
    os.makedirs("character")

for row in char_sheet.rows:
    filename = "{}{}_command.png".format(row[3].value, tag.get(str(row[2].value), "jsalkjdakljdlajdalkd"))
    
    # 테일즈 캐릭터는 5성 이미지만 있어서 보정
    if row[1].value in ["클레스", "유리", "미라", "벨벳"] and row[2].value=="4.5":
        filename = filename.replace("command", "rank5_command")
    
    # 아이디대로 이미지를 복사
    for file in os.listdir(LOCAL_ROOT):
        if filename in file:
            shutil.copyfile("rawimage/" + file, "character/{}.png".format(row[0].value))

# 퍼스널리티 이미지 저장
per_sheet = workbook["퍼스널리티"]

if not os.path.exists("personality"):
    os.makedirs("personality")

for row in per_sheet.rows:    
    # 코드대로 이미지를 복사
    for file in os.listdir(LOCAL_ROOT):
        if str(row[4].value) in file and ".png" in file:
            if row[1].value and "s3" in file:
                shutil.copyfile("rawimage/" + file, "personality/{}(ES).png".format(row[0].value))
                break
            elif not row[1].value and "s3" not in file:
                shutil.copyfile("rawimage/" + file, "personality/{}.png".format(row[0].value))