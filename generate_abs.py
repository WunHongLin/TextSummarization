import os
from openai import AzureOpenAI
import pandas as pd
from docx import Document
import time
from openpyxl import load_workbook
import argparse

#1. load excel and read speaker、case number.
def load_excel():
    file_path = opt.excel_name
    try:
        # load excel by file path.
        file  = pd.read_excel(file_path, usecols=f'A:{opt.range}', header=0)
        # loop for load the content
        for i in range(2):
            row = file.iloc[i]
            # drop out the unname and nan
            cleaned_row = row[~row.index.str.contains('Unnamed', na=False)]
            cleaned_row = cleaned_row.dropna()
            # get case number、 read speaker and load them in list
            if i == 0:
                case_numbers = cleaned_row.tolist()
            elif i == 1:
                participates = cleaned_row.tolist()

        #vertify data
        print("-----------------")
        print(f"case_number: {case_numbers[:2]}, len: {len(case_numbers)}")
        print("-----------------")
        print(f"speakers: {participates[:2]}, len: {len(participates)}")

        #call the follwing function
        load_document(case_numbers[:5], participates[:5])

    except FileNotFoundError:
        print(f"找不到目前的文件_{file_path}")
    except Exception as e:
        print(f"您遇到的問題為_{e}")

#2. according the case number, read the content of doc file.
def load_document(case_numbers, participates):
    #construct list (doc_content) for store info
    doc_content = []
    try:
        #loop for case number
        for number in case_numbers:
            # construct variable date to determine file path
            date = number[1:9]
            # construct file path
            doc_file_path = f"{opt.folder_name}/available-text-{date}-all/{number}_DIA.txt"
            # construct list (lines)
            lines = []
            # read doc content
            with open(doc_file_path, "r", encoding="utf-8") as f:
                for index, line in enumerate(f.readlines()):
                    if index % 4 == 0 or index % 4 == 2: lines.append(line)

            doc_content.append("".join(lines))

        #vertify data
        print("\n-----------------")
        print(f"doc_content: {doc_content[0]}, len: {len(doc_content)}\n")

        #call the follwing function
        identify_speaker(doc_content, participates)

    except FileNotFoundError:
        print(f"找不到目前的文件_{doc_file_path}")
    except Exception as e:
        print(f"您遇到的問題為_{e}")

# 3 the function is additional tool which is to help the below function to identify speakers
def identify_speaker(doc_content, participates):
    # construct list (masses) to store the result of which speaker is people
    masses = []

    # construct variant (end_point, deployment, subscription, api_version, client, example)
    endpoint = "https://aicodeassistant.openai.azure.com/"
    deployment = "gpt-4o-mini"

    subscription = ""
    api_version = "2024-12-01-preview"

    # follwing the above parameters to construct client
    client = AzureOpenAI(
        api_version=api_version,
        azure_endpoint=endpoint,
        api_key=subscription,
    )

    for indices in range(len(doc_content)):
        # construct prompt to give gpt
        prompt = f"文件內容: \n\"\"\"{doc_content[indices]}\"\"\"\n \
                   參與人員: \n\"\"\"{participates[indices]}\"\"\"\n \
                   文件內容為參考的原始文件內容, 參與人員為原始文件中出現的人員(可能為工作人員或是民眾)。 \
                   請你根據文件中出現的語氣,判斷哪一些\"語者\"是民眾,並進行回傳。 \
                   回傳規範: 不需要新增額外的邏輯推倒或是分析, 只需要回傳答案(例如: 語者1)就可以。 \
                   以下為補充資訊: \
                   地點: 裁決中心 \
                   工作人員: 裁決中心的接洽人員 \
                   民眾: 有特定事情需要詢問裁決中心的民眾 \
                 "
        
        try:
            # give the request (includes role、content) to gpt and get the respond from gpt
            response = client.chat.completions.create(
                messages=[{
                    "role": "system",
                    "content": "You are a helpful assistant."
                },{
                    "role": "user",
                    "content": prompt,
                }],
                max_tokens=4096,
                temperature=1.0,
                top_p=1.0,
                model=deployment
            )

            # store the result to list (masses)
            masses.append(response.choices[0].message.content)
        except Exception as e:
            masses.append("##")

        print(f"\r辨識語者-執行進度:[{indices}/{len(doc_content)-1}]", end="")
        time.sleep(1)

    # vertifi data
    print("\n-----------------")
    print(f"masses: {masses[0]}, len: {len(masses)}")

    # call the follwing data
    process_doc_content(doc_content, masses)

# 4. follwing the file, pre-process the content.
def process_doc_content(doc_content, masses):
    # construct list (processed_content) for store processed content.
    processed_content = []

    # loop for doc_content
    for indices, content in enumerate(doc_content):
        # drop out space sentence in the content.
        lines = [line.strip() for line in content.split("\n") if line.strip()]
        # construct list (processed_lines) for store processed current line
        processed_lines = []
        # construct variable (current_apeaker) for control who say the sentence
        current_speaker = None

        for line in lines:
            if "SPK" in line:
                current_speaker = "民眾" if line.split(":")[1] in masses[indices] else "工作人員"
            else:
                # drop out the sentence space
                line = line.strip()
                # pop the last strange token of a sentence
                if line[-1] in "，！#$%^&*()": line = line[:len(line)-1]
                # concatenate the current speaker and line
                processed_line = f"{current_speaker}: {line}。"
                processed_lines.append(processed_line)

        processed_content.append("\n".join(processed_lines))

    #vertify data
    print("\n-----------------")
    print(f"processed_content: {processed_content[0]}, len: {len(processed_content)}\n")

    #call the follwing function
    get_gpt_result(processed_content)

# 5. send the source document to gpt and get the result
def get_gpt_result(processed_content):
    # construct list to store output result
    output_result = [] 
    # prepare relavant parameters
    endpoint = "https://aicodeassistant.openai.azure.com/"
    deployment = "gpt-4o-mini"

    subscription = ""
    api_version = "2024-12-01-preview"

    # follwing the above parameters to construct client
    client = AzureOpenAI(
        api_version=api_version,
        azure_endpoint=endpoint,
        api_key=subscription,
    )

    #example used in prompt
    example = "一名車主撥打電話到裁決中心，詢問自己的一個交通違規案件的相關問題。車主表示自己之前有投報錯車號，但是現在面臨著吊銷牌照的問題。車主詢問是否可以在裁決中心申訴，工作人員表示這不是他們的職責範圍，建議車主聯繫監理機關。工作人員提供了監理機關的電話號碼，車主表示感謝並結束了通話。"

    #loop for the content which is already processed
    for index, content in enumerate(processed_content):
        # construct prompt for gpt
        prompt = f"段落一: \n\"\"\"{content}\"\"\"\n \
                   段落二: \n\"\"\"{example}\"\"\"\n \
                   段落一為需要轉換的原始文件，段落二為生成範例。 \
                   請你根據段落一的內容，生成類似段落二格式的抽象式摘要(abstractive summarization)。 \
                   以下為補充資訊: \
                   地點: 裁決中心 \
                   工作人員: 裁決中心的接洽人員\
                 "

        # store the result from gpt
        try:
            # give the request (includes role、content) to gpt and get the respond from gpt
            response = client.chat.completions.create(
                messages=[{
                    "role": "system",
                    "content": "You are a helpful assistant."
                },{
                    "role": "user",
                    "content": prompt,
                }],
                max_tokens=4096,
                temperature=1.0,
                top_p=1.0,
                model=deployment
            )

            output_result.append(response.choices[0].message.content)
        except Exception as e:
            # the result is denied because of the limitary of openai rule, so append an anchor to list.
            # can delete it in excel mutually
            output_result.append("##")

        print(f"\r生成抽象式摘要-執行進度:[{index}/{len(processed_content)-1}]", end="")
        time.sleep(1)
        
    # vertify data
    print("\n-----------------")
    print(f"output result: {output_result[0]}, len: {len(output_result)}")

    #call the follwing function
    update_excel(output_result)

# 6. update excel
def update_excel(output_result):
    # construct file path
    file_path = opt.excel_name

    try:
        # construct variable (wb) to load the page
        wb = load_workbook(file_path, data_only=True)
        # active wb
        ws = wb.active
        ws.append(output_result)
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