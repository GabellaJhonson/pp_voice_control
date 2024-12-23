import win32com.client
import os
import speech_recognition as sr
import pyautogui
import time

import read_csv

def record_text():
    recognizer = sr.Recognizer()
    microphone = sr.Microphone()
    with microphone as source:
        # recognizer.adjust_for_ambient_noise(source)
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

    # Read Activation words

    commands_dict = read_csv.csv_to_hashmap()
    while True:
        command_txt = record_text()
        command = ""
        print(f"Распознана команда: {command_txt}")
        if commands_dict.get(command_txt.split()[-1]):
            command = str (commands_dict.get(command_txt.split()[-1])[0])


        if command == "previous":
            pyautogui.hotkey('left')
        elif command == "next":
            pyautogui.hotkey('right')
        elif command.startswith("slide"):
            command_txt = "перейти к слайду " + command.split()[1]
        elif command == "exit":
            powerpoint.Quit()
            break
        if command_txt.startswith("перейти к слайду"):
            try:
                slide_number = int(command_txt.split()[-1])
                if 1 <= slide_number <= slides_count:
                    presentation.SlideShowWindow.View.GotoSlide(slide_number)
                    print(f"Переход к слайду {slide_number}")
                else:
                    print("Неверный номер слайда")
            except IndexError:
                print("Не указан номер слайда")
            except ValueError:
                print("Неверный номер слайда")
        else:
            print("Неверная команда")

        # Give some time for PowerPoint to process the commands
        time.sleep(0.2)

if __name__ == "__main__":
    file_path = input("Введите полный путь к презентации PowerPoint: ")
    if os.path.exists(file_path):
        open_powerpoint_presentation(file_path)
    else:
        print("Файл не найден")
