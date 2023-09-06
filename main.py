import spacy
import openpyxl
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans
from collections import defaultdict
from langdetect import detect

# Файл с ключами, один ключ - одна строка
seo = 'seo.txt'


# Чтение ключевых слов из файла
def read_keywords(file_path):
    with open(file_path, 'r') as file:
        keywords = file.read().splitlines()
    return keywords


# Очистка текста от стоп-слов
def remove_stopwords(text):
    doc = nlp(text)
    filtered_words = [token.text for token in doc if not token.is_stop]
    return ' '.join(filtered_words)


# Кластеризация семантического ядра
def cluster_keywords(keywords, num_clusters):
    words = []
    result_keys = []
    for keyword in keywords:
        doc = nlp(keyword.lower())
        result_key = " ".join([token.lemma_ for token in doc if token and not token.is_stop])
        result_keys.append(result_key)

    cleaned_keywords = list([remove_stopwords(keyword) for keyword in result_keys])
    for i in cleaned_keywords:
        words.extend(i.split())

    vectorizer = TfidfVectorizer()
    X = vectorizer.fit_transform(cleaned_keywords)
    # print(X)
    kmeans = KMeans(n_clusters=num_clusters)
    kmeans.fit(X)

    return kmeans.labels_


# def write_in_xl():

def main(file_path):
    # Укажите путь к файлу с ключевыми словами
    keywords = read_keywords(file_path)
    num_clusters = round(len(keywords)/10)  # Укажите количество желаемых кластеров
    labels = cluster_keywords(keywords, num_clusters)

    clusters = defaultdict(list)
    for keyword, label in zip(keywords, labels):
        clusters[label].append(keyword)
    print(len(clusters))
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet[f'A{1}'] = 'Cluster ID'
    sheet[f'B{1}'] = 'Group'
    sheet[f'C{1}'] = 'Key'
    # for index, value in enumerate(column1_data, start=1):
    #     sheet.cell(row=index, column=1, value=value)

    index = 2
    for label, cluster_keywords_list in clusters.items():
        result = f"Cluster {label + 1}: {', '.join(cluster_keywords_list)}\n"
        print(result)
        # with open('result.txt', 'a') as file:
        #     file.write(result + '\n')
        for key in cluster_keywords_list:
            sheet[f'A{index}'] = label + 1
            sheet[f'B{index}'] = cluster_keywords_list[0]
            sheet[f'C{index}'] = key
            index += 1

    # Сохраняем файл
    workbook.save('example.xlsx')
    workbook.close()


if __name__ == "__main__":
    text = ' '.join(read_keywords(seo))
    lang = detect(text)
    # full_lang = to_name(lang).lower()
    if lang == 'en':
        spacy_lang_model = f'{lang.lower()}_core_web_sm'  # en
    else:
        spacy_lang_model = f'{lang.lower()}_core_news_sm'  # ru
    try:
        nlp = spacy.load(spacy_lang_model)
    except IOError:
        spacy.cli.download(spacy_lang_model)
        try:
            nlp = spacy.load(spacy_lang_model)
        except IOError:
            print('Unsupported language')
    main(seo)

