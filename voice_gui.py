# GUI
import tkinter as tk 
import tkinter.messagebox as msgbox  

# 음성 인식
import speech_recognition as sr 

# 웹 브라우저
import webbrowser

# OS 수준 기능
import os
import subprocess  # 외부 프로그램

# 여러 작업을 동시에 수행
import threading

# 볼륨 조절
from ctypes import cast, POINTER  # C 언어 스타일 포인터
from comtypes import CLSCTX_ALL  
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume  # 볼륨 제어

# COM 초기화
import pythoncom

# 시스템 함수 호출
import ctypes

# 경로 관련 조작
from pathlib import Path 

# 파일 복사, 이동, 삭제
import shutil


root = tk.Tk()
root.title('음성어시스턴트')
root.geometry('400x800')
root.resizable(False, False)

is_listening = False
title_label = tk.Label(root, text='음성\n어시스턴트', font=('Serif', 50,))
title_label.pack(pady=10)

btn1 = tk.Button(root, text='눌러서\n시작', font=('Serif', 45), width=10, height=3, relief='raised', bd=2)
btn1.pack(pady=10)


def available_program_show():
    pg_win = tk.Toplevel(root)
    pg_win.title('사용가능 목록')
    pg_win.geometry('400x600')
    pg_win.resizable(False, False)

    frame = tk.Frame(pg_win)
    frame.pack(fill='both', expand=True)

    canvas = tk.Canvas(frame)
    scrollbar = tk.Scrollbar(frame, orient='vertical', command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side='left', fill='both', expand=True)
    scrollbar.pack(side='right', fill='y')

    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    canvas.bind_all("<MouseWheel>", _on_mousewheel)

    tips = [
        '메모장', '계산기', '그림판',
        '워드패드', '스크린 캡처 도구',
        '화면 돋보기', '화면 키보드',
        '작업 관리자', '명령 프롬프트',
        '파워셸', '파일 탐색기',
        '디스크 정리', '레지스트리 편집기',
        '시스템 정보', '장치 관리자',
        '서비스', '이벤트 뷰어',
        '크롬', '엣지',
        '파이어폭스', '오페라',
        '네이버', '카카오톡', '라인',
        '텔레그램', '디스코드',
        '스카이프', '슬랙',
        '줌', '팀즈',
        '왓츠앱',
        '스팀', '에픽 게임즈',
        '배틀넷', '오리진',
        '유플레이', '겜포스',
        '넷플릭스',
        '유튜브',
        '디즈니 플러스',
        '엑셀', '파워포인트',
        '워드', '한글',
        '한셀', '한쇼',
        '스포티파이',
        '비주얼 스튜디오 코드',
        'vs 코드'
    ]

    for tip in tips:
        label = tk.Label(scrollable_frame, text=tip, font=('Serif', 28), justify='left')
        label.pack(pady=5)



def show_tip():
    tip_window = tk.Toplevel(root)
    tip_window.title('명령어 목록')
    tip_window.geometry('400x600')
    tip_window.resizable(False, False)

    tips = [
        "[검색어] 찾아(줘) - 구글에서 검색",
        "유튜브에서 [검색어] 찾아(줘) - 유튜브에서 검색",
        "컴퓨터 꺼줘 - 컴퓨터 종료",
        "컴퓨터 다시 시작 - 컴퓨터 재부팅",
        "컴퓨터 잠금 - 컴퓨터 잠금",
        "소리 켜 - 소리 켜기",
        "소리 꺼 - 소리 끄기",
        "소리 줄여 - 소리 줄이기",
        "소리 높여 - 소리 높이기",
        '프로그램 이름 + 열어or켜줘',
        "메모장 열어 - 메모장 켜줘",
        "계산기 열어 - 계산기 실행",
        "종료하기 - 프로그램 종료",
        '휴지통 비우기 - 휴지통 비우기',
        '임시파일 정리해 줘 - 임시파일 정리',
        '디스크 정리해 줘 - 디스크 정리'
    ]

    for tip in tips:
        label = tk.Label(tip_window, text=tip, font=('Serif', 12),justify='left')
        label.pack(pady=5)
    btn3 = tk.Button(tip_window,text='현재 사용가능한 프로그램 보기',font='Arial 20',command=available_program_show)
    btn3.pack(pady=20)

bt2 = tk.Button(root, text='모든 명령어', font=('Serif', 20), width=22, height=2, command=show_tip, relief='raised', bd=2)
bt2.pack(pady=10)

def set_volume(action):
    pythoncom.CoInitialize()

    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    current_volume = volume.GetMasterVolumeLevelScalar()

    if action == '켜':
        volume.SetMute(0, None)
    elif action == '꺼':
        volume.SetMute(1, None)
    elif action == '줄여':
        volume.SetMasterVolumeLevelScalar(max(0.0, current_volume - 0.25), None)
    elif action == '높여':
        volume.SetMasterVolumeLevelScalar(min(1.0, current_volume + 0.25), None)

def clear_temp_files():
    temp_path = Path(os.getenv('TEMP'))
    system_temp_path = Path(os.getenv('SystemRoot')) / 'Temp'

    try:
        for item in temp_path.iterdir():
            try:
                if item.is_file() or item.is_symlink():
                    item.unlink()
                elif item.is_dir():
                    shutil.rmtree(item)
            except Exception:
                pass  

        for item in system_temp_path.iterdir():
            try:
                if item.is_file() or item.is_symlink():
                    item.unlink()
                elif item.is_dir():
                    shutil.rmtree(item)
            except Exception:
                pass
        
        msgbox.showinfo("임시파일 정리", "임시파일 정리가 완료되었습니다.")
    except Exception as e:
        msgbox.showerror("오류", f"임시파일 정리 중 오류가 발생했습니다:\n{e}")

def voice_loop():
    global is_listening
    recognizer = sr.Recognizer()
    recognizer.energy_threshold = 300
    recognizer.dynamic_energy_threshold = True
    recognizer.pause_threshold = 0.8
    with sr.Microphone() as source:
        try:
            print("듣는 중...")
            btn1.config(text='음성\n인식 중', font=('Serif', 45,))
            root.update()

            audio = recognizer.listen(source, timeout=5, phrase_time_limit=5)
            text = recognizer.recognize_google(audio, language='ko_KR')
            print(f"인식됨: {text}")

            search_keywords = [
                '찾아줘', '찾아줘요', '찾아봐', '찾아주세요', '검색해줘', '검색해 줘',
                '검색하기', '검색해줘요', '검색 부탁해', '검색 좀 해줄래',
                '가 뭔지 알려줘', '이 뭔지 알려줘', '가 뭐야', '이 뭐야',
                '뭔지 알려줘', '이게 뭐야', '이게 뭔지 알려줘',
                '가 뭔지 알려줘', '이게 뭔지', '가 뭔지', '찾기', '검색해', '검색', '찾아',
            ]
            search_keywords.sort(key=len, reverse=True)

            if text in ['안녕', '안녕하세요', '여보세요', '하이', '하이요']:
                msgbox.showinfo('인사', '안녕하세요! 무엇을 도와드릴까요?')

            elif text in ['도움말', '명령어', '명령어 목록', '도움말 보기']:
                show_tip()

            elif '유튜브에서' in text:
                query = text.replace('유튜브에서', '').strip()
                for key in search_keywords:
                    if key in query:
                        query = query.replace(key, '').strip()
                        break
                if query:
                    url = f"https://www.youtube.com/results?search_query={query.replace(' ', '+')}"
                    webbrowser.open(url)
                    msgbox.showinfo('알림', f'유튜브에서 "{query}"를 검색했어요.')
                else:
                    msgbox.showinfo('알림', '유튜브에서 무엇을 검색할지 말씀해주세요.')

            elif any(word in text for word in search_keywords):
                keyword_found = False
                keyword = ""
                for word in search_keywords:
                    if text.endswith(word):
                        keyword = text[:-len(word)].strip()
                        keyword_found = True
                        break
                    elif word in text:
                        keyword = text.replace(word, '').strip()
                        keyword_found = True
                        break

                if keyword_found and keyword:
                    webbrowser.open(f'https://www.google.com/search?q={keyword}')
                    msgbox.showinfo('검색 결과', f'"{keyword}"에 대한 검색 결과를 보여드렸어요.')
                else:
                    msgbox.showinfo('알림', '무엇을 검색할지 말씀해주세요.')

            elif text == '종료하기':
                root.destroy()
                return
            
            elif text in ['컴퓨터 꺼줘', '컴퓨터 종료', '컴퓨터 꺼', '컴퓨터 꺼질래']:
                os.system('shutdown /s /t 3')
                msgbox.showinfo('알림', '컴퓨터가 3초 후에 종료됩니다.')

            elif text in ['컴퓨터 다시 시작', '컴퓨터 재부팅', '컴퓨터 리부팅', '컴퓨터 리스타트']:
                os.system('shutdown /r /t 3')
                msgbox.showinfo('알림', '컴퓨터가 3초 후에 다시 시작됩니다.')

            elif text in ['컴퓨터 잠금', '컴퓨터 잠궈', '컴퓨터 슬립','컴퓨터 절전모드']:
                os.system('rundll32.exe user32.dll,LockWorkStation')
                msgbox.showinfo('알림', '컴퓨터를 잠궜습니다.')

            elif text in ['휴지통 비우기', '휴지통 비워', '휴지통 청소','휴지통 비워줘']:
                try:
                    ctypes.windll.user32.MessageBoxW(0, "휴지통을 비웁니다.", "알림", 0x40 | 0x1)
                    os.system('rd /s /q C:\\$Recycle.Bin')
                    msgbox.showinfo('알림', '휴지통을 비웠어요.')
                except Exception as e:
                    print(f"휴지통 비우기 오류: {e}")
                    msgbox.showerror('오류', '휴지통을 비우는 데 실패했습니다. 관리자 권한이 필요할 수 있습니다.')

            elif '임시 파일' in text or 'temp' in text:
                if '정리해 줘' in text or '삭제해 줘' in text or '청소해 줘' in text:
                    clear_temp_files() 

            elif '디스크' in text:
                if '정리해 줘' in text or '삭제해 줘' in text or '청소해 줘' in text:
                    subprocess.Popen('cleanmgr.exe')
                    msgbox.showinfo('알림','디스크를 정리했어요.')

            elif any(cmd in text for cmd in ['소리', '음량']):
                if '켜' in text or '해제' in text:
                    set_volume('켜')
                    msgbox.showinfo('알림', '소리를 켰어요.')
                elif '꺼' in text or '음소거' in text:
                    set_volume('꺼')
                    msgbox.showinfo('알림', '소리를 껐어요.')
                elif '줄여' in text or '낮춰' in text:
                    set_volume('줄여')
                    msgbox.showinfo('알림', '소리를 줄였어요.')
                elif '높여' in text or '올려' in text:
                    set_volume('높여')
                    msgbox.showinfo('알림', '소리를 높였어요.')
                else:
                    msgbox.showinfo('알림', '무슨 동작을 할지 모르겠어요. (켜, 꺼, 줄여, 높여 중 하나)')

            elif any(cmd in text for cmd in ['켜', '열어', '실행']):
                app_map = {
                    '메모장': 'notepad.exe', '계산기': 'calc.exe', '그림판': 'mspaint.exe',
                    '워드패드': 'write.exe', '스크린 캡처 도구': 'snippingtool.exe',
                    '화면 돋보기': 'magnify.exe', '화면 키보드': 'osk.exe',
                    '작업 관리자': 'taskmgr.exe', '명령 프롬프트': 'cmd.exe',
                    '파워셸': 'powershell.exe', '파일 탐색기': 'explorer.exe',
                    '디스크 정리': 'cleanmgr.exe', '레지스트리 편집기': 'regedit.exe',
                    '시스템 정보': 'msinfo32.exe', '장치 관리자': 'devmgmt.msc',
                    '서비스': 'services.msc', '이벤트 뷰어': 'eventvwr.msc',
                    '크롬': 'start chrome', '엣지': 'start msedge',
                    '파이어폭스': 'start firefox', '오페라': 'start opera',
                    '네이버': 'start https://www.naver.com',
                    '카카오톡': r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\카카오톡.lnk", '라인': 'start line',
                    '텔레그램': 'start telegram', '디스코드': 'start discord',
                    '스카이프': 'start skype', '슬랙': 'start slack',
                    '줌': 'start zoom', '팀즈': 'start teams',
                    '왓츠앱': 'start whatsapp',
                    '스팀': 'start steam', '에픽 게임즈': 'start epicgameslauncher',
                    '배틀넷': 'start Battle.net', '오리진': 'start origin',
                    '유플레이': 'start uplay', '겜포스': 'start gamedesk',
                    '넷플릭스': 'start https://www.netflix.com',
                    '유튜브': 'start https://www.youtube.com/',
                    '디즈니 플러스': 'start https://www.disneyplus.com',
                    '엑셀' : 'start excel', '파워포인트':'start powerpnt',
                    '워드' : 'start winword', '한글' : 'start Hwp',
                    '한셀' : 'start hcell', '한쇼' : 'start hshow',
                    '스포티파이' : 'start spotify',
                    '비주얼 스튜디오 코드' : 'start code',
                    'vs 코드' : 'start code',
                }

                found = False
                for keyword, command in app_map.items():
                    if keyword in text:
                        if keyword == '카카오톡':
                            os.startfile(app_map['카카오톡'])
                        elif command.startswith('start http'):
                            webbrowser.open(command.split(' ')[1])
                        elif command.startswith('start'):
                            subprocess.Popen(command, shell=True)
                        elif command.endswith('.msc'):
                            subprocess.Popen(['mmc.exe', command])
                        else:
                            subprocess.Popen(command)
                        msgbox.showinfo('알림', f'{keyword}를 열었어요.')
                        found = True
                        break

                if not found:
                    msgbox.showinfo('알림', '알 수 없는 앱이에요.')
            else:
                msgbox.showinfo('알림', '이해하지 못했어요. 다시 말씀해주세요.')

        except sr.WaitTimeoutError:
            print("타임아웃: 아무 말도 안함")
            msgbox.showinfo("알림", "아무 말씀도 하지 않으셨습니다.")
        except sr.UnknownValueError:
            print("음성 인식 실패")
            msgbox.showerror("오류", "음성을 인식할 수 없습니다. 다시 시도해 주세요.")
        except sr.RequestError as e:
            print(f"인터넷 오류: {e}")
            msgbox.showerror("오류", "인터넷 연결을 확인해 주세요. 오류: " + str(e))
        except Exception as e:
            print(f"오류 발생: {e}")
            msgbox.showerror("오류", f"예상치 못한 오류가 발생했습니다: {e}")
        finally:
            is_listening = False
            btn1.config(text='눌러서\n시작', font=('Serif', 45,))
            root.update()
            print("대기 상태로 복귀")

def toggle_listening():
    global is_listening
    if not is_listening:
        is_listening = True
        threading.Thread(target=voice_loop, daemon=True).start()

btn1.config(command=toggle_listening)
root.mainloop()