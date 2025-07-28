import tkinter as tk
import tkinter.messagebox as msgbox
import speech_recognition as sr
import webbrowser

root = tk.Tk()
root.title('음성 어시스턴트')
root.geometry('400x700')
root.resizable(False,False)

new_img = tk.PhotoImage(file='title_label.png')
new_img2 = new_img.subsample(3,3)
title_label = tk.Label(root,image=new_img2)
title_label.pack(pady=10)

mic_im = tk.PhotoImage(file='BUTTON_IMG.png')
mic_img = mic_im.subsample(3,3)
btn1 = tk.Button(root,image=mic_img,relief='flat',bd=0)
btn1.pack()

tel = tk.PhotoImage(file='tell-001.png').subsample(3,3)
pres = tk.PhotoImage(file='press-001.png').subsample(3,3)
status_label = tk.Label(root,image=pres)
status_label.pack()

def click():
    print('버튼 클릭됨')
    recognizer = sr.Recognizer()
    recognizer.energy_threshold = 1000
    recognizer.pause_threshold = 1
    with sr.Microphone() as source:
        print('말하세요')
        status_label.config(image=tel)
        root.update()
        audio = recognizer.listen(source,timeout=7)
        try:
            text = recognizer.recognize_google(audio,language='ko_KR')
            print(text)
            msgbox.showinfo('입력된 메시지',text)
            status_label.config(image=pres)
            root.update()
            for word in [
                '찾아','찾아줘','찾아줘요','찾아봐','찾아주세요',
                '찾기','좀 찾아','검색','검색해줘','검색해 줘',
                '검색하기','검색해줘요','검색 부탁해','검색 좀 해줄래',
                '검색해','가 뭔지 알려줘','이 뭔지 알려줘','가 뭐야','이 뭐야'
                ]:
                if word in text:
                    keyword = text.split(word)[0].strip()
                    url = f'https://www.google.com/search?q={keyword}'
                    print(keyword)
                    webbrowser.open(url)
            if text == '종료하기':
                root.destroy()
        except:
            print('다시시도해주세요')
            msgbox.showerror('다시 시도해주세요')

btn1.config(command=click)
root.mainloop()