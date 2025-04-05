import json
import speech_recognition as sr
from pptx import Presentation
import spacy

# Загрузка NLP-модели
nlp = spacy.load("ru_core_news_sm")

# Конфигурация
CONFIG_FILE = "config.json"
PPTX_FILE = "auto_presentation.pptx"

def load_config():
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = json.load(f)
        # Преобразуем ключевые слова в леммы
        for section in config["sections"]:
            lemmatized_keywords = []
            for phrase in section["keywords"]:
                doc = nlp(phrase)
                lemmatized_keywords.extend([token.lemma_ for token in doc])
            section["lemmas"] = list(set(lemmatized_keywords))
        return config

def recognize_speech():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Говорите...")
        audio = r.listen(source)
        
    try:
        text = r.recognize_google(audio, language="ru-RU").lower()
        print(f"Распознано: {text}")
        return text
    except sr.UnknownValueError:
        return ""
    except sr.RequestError:
        print("Ошибка сервиса распознавания")
        return ""

def create_slide(presentation, section):
    try:
        slide_layout = presentation.slide_layouts[1]  # Заголовок + текст
        slide = presentation.slides.add_slide(slide_layout)
        
        # Заголовок
        title_shape = slide.shapes.title
        title_shape.text = section["title"]
        
        # Контент
        content_shape = slide.placeholders[1]
        tf = content_shape.text_frame
        for item in section["content"]:
            p = tf.add_paragraph()
            p.text = item
            p.level = 0
    except Exception as e:
        print(f"Ошибка создания слайда: {e}")

def main():
    config = load_config()
    prs = Presentation()
    
    try:
        while True:
            text = recognize_speech()
            if "стоп" in text:
                break
                
            if text:
                doc = nlp(text)
                user_lemmas = [token.lemma_ for token in doc]
                print(f"Леммы пользователя: {user_lemmas}")
                
                # Поиск подходящего раздела
                for section in config["sections"]:
                    if any(lemma in user_lemmas for lemma in section["lemmas"]):
                        print(f"Найден раздел: {section['title']}")
                        create_slide(prs, section)
                        prs.save(PPTX_FILE)  # Сохраняем после каждого изменения
                        break
                
    except KeyboardInterrupt:
        pass
    
    prs.save(PPTX_FILE)
    print(f"Презентация сохранена как {PPTX_FILE}")

if __name__ == "__main__":
    main()
