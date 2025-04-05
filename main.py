import os
import requests
import subprocess
import speech_recognition as sr
from pptx import Presentation

# Настройки
PPTX_FILE = "auto_presentation.pptx"
OLLAMA_URL = "http://localhost:11434/api/generate"
MODEL_NAME = "llama3.2"

def recognize_speech():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("\n🎤 Говорите... (скажите 'стоп' для выхода)")
        audio = r.listen(source, phrase_time_limit=15)
    try:
        return r.recognize_google(audio, language="ru-RU")
    except:
        return ""

def generate_slide_data(text):
    try:
        # Генерация заголовка
        title_prompt = f"""<|begin_of_text|>
        [INST] Выдели основную тему из текста (3-5 слов):
        "{text}". Ответ дай только самой темой без пояснений. [/INST]
        """
        
        title_response = requests.post(
            OLLAMA_URL,
            json={
                "model": MODEL_NAME,
                "prompt": title_prompt,
                "stream": False,
                "options": {"temperature": 0.3}
            }
        ).json().get("response", "").strip().replace('"', '')

        # Генерация контента
        content_prompt = f"""<|begin_of_text|>
        [INST] Сгенерируй 3 ключевых пункта для слайда на тему "{title_response}"
        на основе текста: "{text}". Формат: маркированный список. [/INST]
        """
        
        content_response = requests.post(
            OLLAMA_URL,
            json={
                "model": MODEL_NAME,
                "prompt": content_prompt,
                "stream": False,
                "options": {"temperature": 0.5}
            }
        ).json().get("response", "").split("\n")

        return title_response, [line for line in content_response if line.strip()]
    
    except Exception as e:
        print(f"Ошибка генерации: {e}")
        return "Ошибка", ["Не удалось сгенерировать контент"]

def create_slide(presentation, title, content):
    slide_layout = presentation.slide_layouts[1]
    slide = presentation.slides.add_slide(slide_layout)
    slide.shapes.title.text = title[:50]  # Заголовок
    content_box = slide.placeholders[1]
    for line in content[:3]:  # Контент (первые 3 пункта)
        content_box.text_frame.add_paragraph().text = line[:100]

def refresh_powerpoint():
    script = f'''
    tell application "Microsoft PowerPoint"
        activate
        close active presentation saving yes
        open POSIX file "{os.path.abspath(PPTX_FILE)}"
        delay 3
        tell active presentation
            set slideCount to count of slides
            if slideCount > 0 then
                set curSlide to slide index of slide range of selection of document window 1
                go to slide view of document window 1 number slideCount
            end if
        end tell
    end tell
    '''
    subprocess.run(["osascript", "-e", script])


def main():
    # Создаём или загружаем существующую презентацию
    if os.path.exists(PPTX_FILE):
        prs = Presentation(PPTX_FILE)
    else:
        prs = Presentation()

    try:
        while True:
            text = recognize_speech().strip()
            if not text:
                continue
            
            print(f"\n🔊 Распознано: {text}")
            
            if "стоп" in text.lower():
                break
            
            # Генерация данных для слайда
            title, content = generate_slide_data(text)
            print(f"\n📄 Заголовок: {title}")
            print("📌 Контент:", "\n • ".join(content))
            
            # Создание нового слайда
            create_slide(prs, title, content)
            
            # Сохранение изменений и обновление PowerPoint
            prs.save(PPTX_FILE)
            refresh_powerpoint()
            print("✅ Слайд сохранен и отображен!")
            
    finally:
        prs.save(PPTX_FILE)
        print(f"\n💾 Презентация сохранена как: {PPTX_FILE}")

if __name__ == "__main__":
    main()
