def check_category(sentence, questions, categories, answers):
    # loop for questions to check whether the sentence belongs to it
    for index in range(len(questions)):
        if sentence in questions[index]:
            return categories[index], answers[index]