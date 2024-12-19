import win32com.client
import os
import speech_recognition as sr
import pyautogui
import time

def record_text():
    recognizer = sr.Recognizer()
    microphone = sr.Microphone()
    with microphone as source:
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
    try:
        return recognizer.recognize_google(audio, language="ru-RU").lower()
    except sr.UnknownValueError:
        return "не понял"
    except sr.RequestError:
        return "ошибка"

def open_powerpoint_presentation(file_path):
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True

    presentation = powerpoint.Presentations.Open(file_path)
    slides_count = presentation.Slides.Count

    # Set slideshow settings
    slideshow_settings = presentation.SlideShowSettings
    slideshow_settings.ShowWithNarration = False
    slideshow_settings.ShowWithAnimation = False
    slideshow_settings.RangeType = 2  # ppShowAll
    slideshow_settings.StartingSlide = 1
    slideshow_settings.EndingSlide = slides_count

    # Start slideshow
    presentation.SlideShowSettings.Run()

    while True:
        command = record_text()
        print(f"Распознана команда: {command}")

        if command == "назад":
            pyautogui.hotkey('left')
        elif command == "вперёд":
            pyautogui.hotkey('right')
        elif command.startswith("перейти к слайду"):
            try:
                slide_number = int(command.split()[3])
                if 1 <= slide_number <= slides_count:
                    presentation.SlideShowWindow.View.GotoSlide(slide_number)
                    print(f"Переход к слайду {slide_number}")
                else:
                    print("Неверный номер слайда")
            except IndexError:
                print("Не указан номер слайда")
            except ValueError:
                print("Неверный номер слайда")
        elif command == "exit":
            powerpoint.Quit()
            break
        else:
            print("Неверная команда")

        # Give some time for PowerPoint to process the commands
        time.sleep(0.5)

if __name__ == "__main__":
    file_path = input("Введите полный путь к презентации PowerPoint: ")
    if os.path.exists(file_path):
        open_powerpoint_presentation(file_path)
    else:
        print("Файл не найден")
