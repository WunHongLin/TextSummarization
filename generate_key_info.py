import os
from docx import Document
from openpyxl import load_workbook
import argparse

# 1. Get the date of the folder
def get_folder_name():
    # construct file_path
    folder_path = opt.folder_name
    # vertify data
    print("-----------------")
    print(f"包含資料夾: {os.listdir(folder_path)}")
    # call follwing function
    get_case_number(os.listdir(folder_path))

# 2. Follwing the date, to construct list (case_number, date)
def get_case_number(include_folder):
    # construct list (case_number, dates) to store key info
    case_number, dates = [], []
    # loop for paremeter include_folder to get the file_path
    for folder in include_folder:
        # construct variant (file_path)
        file_path = f"{opt.folder_name}/{folder}"
        # construct variable (date)
        date = folder.split("-")[2]
        # loop for file_path to get date and case_number
        for file in os.listdir(file_path):
            dates.append(f"{date[0:4]}年{date[4:6]}月{date[6:8]}日")
            case_number.append(file.split("_")[0])

    # vertify data
    print("\n-----------------")
    print(f"doc_content: {case_number[0]}, len: {len(case_number)}, date: {dates[0]}, len: {len(dates)}")

    # call the follwing function
    get_speakers(case_number, dates)

# 3. Follwing the above list to determine variant (file_path) and get the speakers
def get_speakers(case_numbers, dates):
    # construct list (speakers) to store info
    speakers = []
    try:
        # loop for case_numbers
        for case in case_numbers:
            # set variant (date)
            date = case[1:9]
            # follwing paremeter (case_numbers, dates) to set file_path
            file_path = f"{opt.folder_name}/available-text-{date}-all/{case}_DIA.txt"
            # cpnstruct list (participate)
            participate = [] 
            # read txt content
            with open(file_path, "r", encoding="utf-8") as f:
                for line in f.readlines():
                    if "SPK" in line:
                        participate.append(line.split(":")[1][:-1])

            # append the speaker content to list
            speakers.append(",".join(list(set(participate))))

        # vertify data 
        print("\n-----------------")
        print(f"speaker: {speakers[0]}, len: {len(speakers)}")

        #call the follwing function   
        update_excel(dates, case_numbers, speakers)   

    except FileNotFoundError:
        print(f"找不到目前的文件_{file_path}")
    except Exception as e:
        print(f"您遇到的問題為_{e}")

# 4. write down the info to excel
def update_excel(dates, case_numbers, speakers):
    # construct file path
    file_path = opt.excel_name

    try:
        # construct variable (wb) to load the page
        wb = load_workbook(file_path, data_only=True)
        # active wb
        ws = wb.active

        # store the data to excel
        ws.append(dates)
        ws.append(case_numbers)
        ws.append(speakers)

        # save file
        wb.save(file_path)
        print("\n-----------------")
        print("結果已經更新至檔案中")

    except FileNotFoundError:
        print(f"找不到目前的文件_{file_path}")
    except Exception as e:
        print(f"您遇到的問題為_{e}")

if __name__ == "__main__":
    # construct variant (parser) to set argparser lib
    parser = argparse.ArgumentParser()
    # add argurment
    parser.add_argument("--folder_name", type=str, default="250221_第二批逐字稿")
    parser.add_argument("--excel_name", type=str, default="250221_逐字稿摘要.xlsx")
    # construct variant (opt) to present paremeter
    opt = parser.parse_args()
    print(opt)
    get_folder_name()