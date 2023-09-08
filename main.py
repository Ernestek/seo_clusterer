import spacy
import openpyxl
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans
from collections import defaultdict
from langdetect import detect

# Файл с ключами, один ключ - одна строка
seo = 'seo_en.txt'


# Чтение ключевых слов из файла
def read_keywords(file_path):
    with open(file_path, 'r') as file:
        keywords = file.read().splitlines()
    return keywords


# Кластеризация семантического ядра
def cluster_keywords(keywords, num_clusters, stop_words: list):
    result_keys = []
    other = []
    for keyword in keywords:
        stop_words = list(map(lambda x: x.lower(), stop_words))
        doc = nlp(keyword.lower())
        result_key = []
        stop = False
        for token in doc:
            if token.text.lower() in stop_words:
                stop = True
                other.append(keyword)
                break
            else:
                result_key.append(token.lemma_)

        if not stop:
            result_key = ' '.join([token.lemma_ for token in doc if token and not token.is_stop])
            result_keys.append(result_key)

    vectorizer = TfidfVectorizer()
    X = vectorizer.fit_transform(result_keys)
    kmeans = KMeans(n_clusters=num_clusters)
    kmeans.fit(X)
    return kmeans.labels_, result_keys, other


def main(file_path):
    # Укажите путь к файлу с ключевыми словами
    stop_words = ['Buy']
    keywords = read_keywords(file_path)
    num_clusters = round(len(keywords)/10)  # Укажите количество желаемых кластеров
    result_info = cluster_keywords(keywords, num_clusters, stop_words=stop_words)

    clusters = defaultdict(list)
    for keyword, label in zip(result_info[1], result_info[0]):
        clusters[label].append(keyword)

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet[f'A{1}'] = 'Cluster ID'
    sheet[f'B{1}'] = 'Group'
    sheet[f'C{1}'] = 'Key'

    index = 2
    for label, cluster_keywords_list in clusters.items():
        result = f"Cluster {label + 1}: {', '.join(cluster_keywords_list)}\n"
        print(result)

        for key in cluster_keywords_list:
            sheet[f'A{index}'] = label + 1
            sheet[f'B{index}'] = cluster_keywords_list[0]
            sheet[f'C{index}'] = key
            index += 1

    for key in result_info[2]:
        sheet[f'A{index}'] = 0
        sheet[f'B{index}'] = 'Stop_world'
        sheet[f'C{index}'] = key
        index += 1
    # Сохраняем файл
    workbook.save('example.xlsx')
    workbook.close()


if __name__ == "__main__":
    text = ' '.join(read_keywords(seo))
    lang = detect(text)

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

