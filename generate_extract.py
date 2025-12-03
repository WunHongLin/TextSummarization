import pandas as pd
import numpy as np
from docx import Document
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import jieba
from openpyxl import load_workbook
from openai import AzureOpenAI
import argparse
import time

# 1. loading excel and read abs_summarization、 case_number、 speakers.
def process_excel():
    file_path = opt.excel_name
    try:
        # load excel by file path.
        file  = pd.read_excel(file_path, usecols=f'A:{opt.range}', header=0)
        # loop for load the content
        for i in range(3):
            row = file.iloc[i]
            # drop out the unname and nan
            cleaned_row = row[~row.index.str.contains('Unnamed', na=False)]
            cleaned_row = cleaned_row.dropna()
            # get case number、 read speaker and load them in list
            if i == 0:
                case_numbers = cleaned_row.tolist()
            elif i == 1:
                participates = cleaned_row.tolist()
            elif i == 2:
                abs_summarizations = cleaned_row.to_list()

        #vertify data
        print("-----------------")
        print(f"case_number: {case_numbers[:2]}, len: {len(case_numbers)}")
        print("-----------------")
        print(f"speakers: {participates[:2]}, len: {len(participates)}")
        print("-----------------")
        print(f"abs_summarizations: {abs_summarizations[:2]}, len: {len(abs_summarizations)}")

        #call the follwing function
        load_document(case_numbers, participates, abs_summarizations)

    except FileNotFoundError:
        print(f"找不到目前的文件_{file_path}")
    except Exception as e:
        print(f"您遇到的問題為_{e}")

# 2. according the case number, read the content of doc file.
def load_document(case_numbers, participates, abs_summarizations):
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
        identify_speaker(doc_content[:5], participates[:5], abs_summarizations[:5])

    except FileNotFoundError:
        print(f"找不到目前的文件_{doc_file_path}")
    except Exception as e:
        print(f"您遇到的問題為_{e}")

# 3 the function is additional tool which is to help the below function to identify speakers
def identify_speaker(doc_content, participates, abs_summarizations):
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
                   文件概述: \n\"\"\"{abs_summarizations[indices]}\"\"\"\n\
                   參與人員: \n\"\"\"{participates[indices]}\"\"\"\n \
                   (文件內容)為參考的原始文件內容, (文件概述)為根據文件內容所產生的抽象是摘要, (參與人員)為文件內容中出現的人員(有工作人員, 民眾, 開場白三種可能)。 \
                   請你根據(文件內容)以及(文件概述)的內容, 判斷文件內容中的那些speaker為民眾, 其中民眾可能不只一個。 \
                   1. 如果\"語者\"為民眾的話, 其特徵包含(在文件中提供個人的身分證或是車牌資訊), 或是(有疑問需要詢問工作人員。) \
                   2. 如果\"語者\"為工作人員的話, 其特徵為(說話語氣較直接), (提供具體方向讓民眾可以操作)。 \
                   3. 如果\"語者\"為開場白的話, 其特徵為(不參與直接對話), (在文件中一開始時說明簡易事項)。 \
                   以下為回傳規範說明以及範例。 \
                   回傳規範說明: 需要針對出現的各個\"語者\"進行標籤, 說明其代表的腳色, 並在完成之後簡單說明可能的民眾為哪一個\"語者\"。 \
                   回傳範例: **語者1**: 這位參與者提供了身分證號碼，並詢問如何處理交通違規罰款，因此可以被判定為民眾。 **語者**: 這位參與者提供了具體的指導，回答了民眾的問題，並且語氣較直接，因此可以被判定為工作人員。 結果: 語者1\
                   以下為補充資訊: \
                   地點: 裁決中心 \
                   工作人員: 裁決中心的接洽人員 \
                   民眾: 有特定事情需要請裁決中心協助的人 \
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
            masses.append(response.choices[0].message.content.split(": ")[-1])
        except Exception as e:
            masses.append("##")

        print(f"\r辨識語者-執行進度:[{indices}/{len(doc_content)-1}]", end="")

    # vertifiy data
    print("\n-----------------")
    print(f"masses: {masses[0]}, len: {len(masses)}")

    # call the follwing data
    process_doc_content(doc_content, masses, abs_summarizations)

# 4. follwing the file, pre-process the content.
def process_doc_content(doc_content, masses, abs_summarizations):
    # construct list (processed_content) for store processed content.
    processed_content = []
    # loop for doc_content
    for count, content in enumerate(doc_content):
        # drop out space sentence in the content.
        lines = [line.strip() for line in content.split("\n") if line.strip()]
        # construct list (processed_lines) for store processed current line
        processed_lines = []
        # construct variable (current_apeaker) for control who say the sentence
        current_speaker = None

        for line in lines:
            if "SPK" in line:
                current_speaker = "民眾" if line.split(":")[1] in masses[count] else "工作人員"
            else:
                # drop out the sentence space
                line = line.strip()
                # pop the last strange token of a sentence
                if line[-1] in "，！#$%^&*()": line = line[:len(line)-1]
                # concatenate the current speaker and line
                processed_line = f"{current_speaker}:{line}。"
                processed_lines.append(processed_line)

        processed_content.append("\n".join(processed_lines))

    #vertify data
    print("\n-----------------")
    print(f"processed_content: {processed_content[0]}, len: {len(processed_content)}")

    # call the follwing fumction
    extract_summary(processed_content, masses, abs_summarizations)

# 5. follwing the processed_content to extract summary
def extract_summary(processed_content, masses, abs_summarization):
    # construct list (extractive_summaries) to stored result
    extractive_summaries = []
    # construct a TF-IDF vector translater
    vectorizer = TfidfVectorizer(
        tokenizer=lambda x: jieba.cut(x, cut_all=False),
        token_pattern=None 
        # use personal tokenizer
    )

    # construct the variant (preposition)
    preposition = "有、是、與、的、在、從、到、自、由、往、朝、於、對、跟、與、給、為、因、因為、由於、由、因應、靠、以、用、替、為了、關於、對於、就、有關、自從、直到、至、依、據、憑、循、藉由、藉、趁、比、如、按照、根據、依照、照、沿著、沿、順著、隨、隨著、除了、除了以外、不如"
    # translate the above variant to list
    prepositions = preposition.split("、")
    # replace the preposition to space in abs_summarization
    for index, abs_summary in enumerate(abs_summarization):
        current_summary = abs_summary
        for prop in prepositions:
            current_summary = current_summary.replace(prop, "")
        abs_summarization[index] = current_summary

    # follwing processed_content, to drop out the sentence which is too short
    for index in range(len(processed_content)):
        current_content = processed_content[index].split("\n")
        result_content = []
        for content in current_content:
            if len(content.split(":")[1]) > 6:
                result_content.append(content)
        processed_content[index] = "\n".join(result_content)

    # follwing the processed_content above, drop out something might influence result
    # factors = ["權益", "服務", "分機", "忙碌"]
    # for index in range(len(processed_content)):
    #     current_content = processed_content[index].split("\n")
    #     result_content = []
    #     for content in current_content:
    #         flag = False
    #         for factor in factors:
    #             if factor in content: 
    #                 flag = True
    #                 break
    #         if not flag: result_content.append(content)

    #     # restored content to list (processed_content)
    #     processed_content[index] = "\n".join(result_content)
    
    # according each abs_summarization to caculate the similarity value 
    for i in range(len(processed_content)):
        # split the \n of the element in processed_content
        processed_sentences = processed_content[i].split('\n')
        # get abs_summary
        abs_summary = abs_summarization[i]
        # concatenate the sentences and abs_summary
        all_texts = processed_sentences + [abs_summary]
        # send all_text to TF-IDF vector transklater
        tfidf_matrix = vectorizer.fit_transform(all_texts)

        # construct two variant (sentence_vectors, summary_vectors) to get sentence and summary vectors
        sentence_vectors, summary_vector = tfidf_matrix[:-1], tfidf_matrix[-1:]
        
        # caculate cosine similarity between sentence and summary
        sentence_similarities = cosine_similarity(sentence_vectors, summary_vector)
        
        # follwing the length of processed_content to dertimine the length of result
        content_length = len(processed_content[i].split("\n"))

        # construct the variant (ext_len) to set final length
        ext_len = 0
        if content_length > 0 and content_length < 10:
            ext_len = -6
        elif content_length >= 10 and content_length <= 20:
            ext_len = -8
        else:
            ext_len = -10

        #get top ext_len-th of sentences
        top_sentence_indices = np.argsort(sentence_similarities.flatten())[ext_len:]
        # construct a list (selected_sentences) to stored result
        selected_sentences = [processed_sentences[indice] for indice in top_sentence_indices]
        # because of the list (selected_sentences) is not sorted, might interrupt the result, so below code should sort the list
        selected_sentences.sort(key=lambda x: processed_sentences.index(x))
        extractive_summaries.append('\n'.join(selected_sentences))

    # vertify data
    print("\n-----------------")
    print(f"extractive_summaries: {extractive_summaries[0]}, len: {len(extractive_summaries)}")
    # call the follwing fumction
    process_extract_summary(extractive_summaries)

# 6. according the result of above function, to let each sentence which contains more than one sentence split.
def process_extract_summary(extractive_summaries):
    # construct a list (processd_extracts) to store the result
    processd_extracts = []
    # loop for paremeter (extractive_summaries)
    for s_index, summary in enumerate(extractive_summaries):
        try:
            # construct list (processed_sentence) to store the processed content
            processed_sentence = []
            # construct list (sentences) to store sentence in each summaries
            sentences = summary.split("\n")

            # loop for list (sentences) to process each sentence
            for sentence in sentences:
                speaker, content = sentence.split(":")[0], sentence.split(":")[1]
                # translate content from str to list
                content = list(content)
                # construct two variant (left, right) to represent pointer
                left, right = 0, 0
                # loop for content to split 。!?
                for index, word in enumerate(content):
                    # word is !。?
                    if word in "!。?！？":
                        right = index
                        # avoid the content only has one character
                        clause = content[left:right]
                        if len(clause) > 1:
                            processed_sentence.append(f"{speaker}:{"".join(clause)}{content[right]}")
                        left = right + 1

            processd_extracts.append("\n".join(processed_sentence))

        except Exception as e:
            print(s_index)
            print(summary)

    # vertify data
    print("\n-----------------")
    print(f"processed extractive summary: {processd_extracts[0]}, len: {len(processd_extracts)}")
    # call the follwing function
    update_excel(processd_extracts)

# 7. update the extractive summarization to excel
def update_excel(processd_extracts):
    # construct file path
    file_path = opt.excel_name

    try:
        # construct variable (wb) to load the page
        wb = load_workbook(file_path, data_only=True)
        # active wb
        ws = wb.active
        ws.append(processd_extracts)
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
    process_excel()