# Top-3 Annotation Suggestion Task

This project consists of five main Python scripts designed to automate the extraction of key information and generation of annotation suggestions. Below is a description of each script:

## 📁 Script Descriptions

### 1. `generate_key_info.py`
- **Function**: Extracts key information from the raw data, including:
  - Date
  - Case number
  - List of participants
- **Output**: Saves the extracted information into corresponding files for later use.

---

### 2. `generate_abstract.py`
- **Function**: Generates an **abstractive summary** based on meeting transcripts or other source data.
- **Output**: Abstractive summary.

---

### 3. `generate_extract.py`
- **Function**: Produces an **extractive summary** by comparing and selecting relevant segments from the abstractive summary.
- **Output**: Extractive summary.
---

### 4. `split_speaker.py`
- **Function**: Since the extractive summary includes content from both mass and workers, this script splits and labels the segments according to speaker.
- **Output**: Separates and tags the speech segments of different speakers accordingly.

---

### 5. `top_k_indices.py`
- **Function**: According to each sentence in **mass summary content** and selects the most representative FAQ question.
- **Output**: Generates the top-3 annotation suggestions (Top-K, where K = 3).

---

## 🚀 Workflow

Follow the steps below to execute the full annotation suggestion pipeline:

1. Place the dataset (e.g., `250703_逐字稿摘要`) inside the `summarization/` directory.
2. Create an Excel file (e.g., `250703_逐字稿摘要.xlsx`) and place it in the same `summarization/` directory.
3. Execute the following scripts in order:
   - `generate_key_info.py`
   - `generate_abstract.py`

   After running `generate_abstract.py`, manually check the output files for any occurrences of the `"##"` symbol. If found, remove correspond column manually.

4. After cleaning the summaries, continue running:
   - `generate_extract.py`
   - `split_speaker.py`

5. Create an Excel file to store the top-3 annotation suggestions (e.g., `FAQ_Result.xlsx`) then place file (`裁決中心常見問題`) inside the `summarization/` directory.
6. Finally, run `top_k_indices.py` to generate the top-3 suggestions.

(Note: In case of any changes to the "逐字稿" file name or path, please update the corresponding references at line 48 in generate_key_info.py, line 51 in generate_abs.py, and line 58 in generate_extract.py.)

## 🖥️ Command Line Usage

Below are the command-line instructions for executing each script:

### 1. `generate_key_info.py`
```
python generate_key_info.py --folder_name "數據集位置" --excel_name "excel檔案位置(負責存放程式1到程式4的資訊)"
```

### 2. `generate_abstract.py`
```
python generate_abstract.py --folder_name "數據集位置" --excel_name "excel檔案位置(負責存放程式1到程式4的資訊)" --range "檔案最後欄位位置"
```

### 3. `generate_extract.py`
```
python generate_extract.py --folder_name "數據集位置" --excel_name "excel檔案位置(負責存放程式1到程式4的資訊)" --range "檔案最後欄位位置"
```

### 4. `split_speaker.py`
```
python split_speaker.py --excel_name "excel檔案位置(負責存放程式1到程式4的資訊)" --range "檔案最後欄位位置"
```

### 5. `top_k_indices.py`
```
python top_k_indices.py --excel_name "excel檔案位置(負責存放程式1到程式4的資訊)" --FAQ_Result "前三標註建議excel檔案位置" --excel_range "檔案最後欄位位置"
```
