import argparse
from openpyxl import load_workbook
import pandas as pd

# 1. loading the excel according the paremeter (excel_name)
def load_excel():
    file_path = opt.excel_name
    try:
        # load excel by file path.
        file  = pd.read_excel(file_path, usecols=f'A:{opt.range}', header=0)
        # get extractive summary
        row = file.iloc[3]
        # drop out the unname and nan
        cleaned_row = row[~row.index.str.contains('Unnamed', na=False)]
        cleaned_row = cleaned_row.dropna()
        # construct list (extractive_summary) to store summary
        extractive_summaries = cleaned_row.tolist()
        #vertify data
        print("-----------------")
        print(f"extractive summary: {extractive_summaries[0]}, len: {len(extractive_summaries)}")

        #call the follwing function
        split_sentences(extractive_summaries)

    except FileNotFoundError:
        print(f"找不到目前的文件_{file_path}")
    except Exception as e:
        print(f"您遇到的問題為_{e}")

# 2. split the sentence to two list
def split_sentences(extractive_summaries):
    # construct two list (worker_contents, masses_content) to store the info
    worker_summaries = []
    masses_summaries = []
    # loop for paremeter (extractive_summaries)
    for summary in extractive_summaries:
        # construct two list (worker_sentences, masses_sentences)
        worker_sentences = []
        masses_sentences = []
        # split the summary to list (sentences)
        sentences = summary.split("\n")
        # loop for sentences
        for sentence in sentences:
            # check whether the sentence is worker or masses
            if "工作人員" in sentence: 
                worker_sentences.append(sentence)
            elif "民眾" in sentence: 
                masses_sentences.append(sentence)

        # join the list to a string
        worker_content = "\n".join(worker_sentences)
        masses_content = "\n".join(masses_sentences)

        # append data to list
        worker_summaries.append(" " if len(worker_content) == 0 else worker_content)
        masses_summaries.append(" " if len(masses_content) == 0 else masses_content)

    #vertify data
    print("-----------------")
    print(f"worker_summaries: {worker_summaries[0]}, len: {len(worker_summaries)}")
    print("-----------------")
    print(f"masses_summaries: {masses_summaries[0]}, len: {len(masses_summaries)}")

    # call the follwing function
    update_excel(worker_summaries, masses_summaries)

# 3. update the excel 
def update_excel(worker_summaries, masses_summaries):
    # construct file path
    file_path = opt.excel_name

    try:
        # construct variable (wb) to load the page
        wb = load_workbook(file_path, data_only=True)
        # active wb
        ws = wb.active
        ws.append(masses_summaries)
        ws.append(worker_summaries)
        # save file
        wb.save(file_path)
        print("\n-----------------")
        print("結果已經更新至檔案中")

    except Exception as e:
        print(f"您遇到的問題為_{e}")

if __name__ == "__main__":
    # construct variant (parser) to set argparser lib
    parser = argparse.ArgumentParser()
    # add argurment
    parser.add_argument("--folder_name", type=str, default="250221_第二批逐字稿")
    parser.add_argument("--excel_name", type=str, default="250221_逐字稿摘要.xlsx")
    parser.add_argument("--range", type=str, default="AFY")
    # construct variant (opt) to present paremeter
    opt = parser.parse_args()
    print(opt)
    load_excel()