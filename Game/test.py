import time
import threading
import win32gui
import win32process
import win32con
import win32api
import psutil
import ctypes
import PySimpleGUI as sg
import pyautogui

def get_hwnds_by_title_keyword(keyword):
    hwnds = []
    def callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if keyword in title:
                hwnds.append(hwnd)
    win32gui.EnumWindows(callback, None)
    return hwnds

def activate_window(hwnd):
    try:
        if not hwnd or not win32gui.IsWindow(hwnd):
            return False
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        fg_thread = win32api.GetCurrentThreadId()
        hwnd_thread, _ = win32process.GetWindowThreadProcessId(hwnd)
        ctypes.windll.user32.AttachThreadInput(fg_thread, hwnd_thread, True)
        win32gui.SetForegroundWindow(hwnd)
        win32gui.SetFocus(hwnd)
        win32gui.BringWindowToTop(hwnd)
        ctypes.windll.user32.AttachThreadInput(fg_thread, hwnd_thread, False)
        return True
    except Exception as e:
        print(f"激活窗口出错: {e}")
        return False

# def press_ctrl_3():
#     VK_CONTROL = 0x11
#     VK_3 = 0x33
#     win32api.keybd_event(VK_CONTROL, 0, 0, 0)
#     win32api.keybd_event(VK_3, 0, 0, 0)
#     time.sleep(0.05)
#     win32api.keybd_event(VK_3, 0, win32con.KEYEVENTF_KEYUP, 0)
#     win32api.keybd_event(VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
def press_ctrl_3():
    pyautogui.hotkey('ctrl', '3')

def loop_send(hwnds, stop_event, delay, keyword):
    last_non_game_hwnd = None

    def is_game_window(hwnd):
        # 这里用标题判断是否游戏窗口
        if not hwnd or not win32gui.IsWindow(hwnd):
            return False
        title = win32gui.GetWindowText(hwnd)
        return keyword in title

    idx = 0
    while not stop_event.is_set():
        curr_hwnd = win32gui.GetForegroundWindow()
        if curr_hwnd and win32gui.IsWindow(curr_hwnd) and not is_game_window(curr_hwnd):
            last_non_game_hwnd = curr_hwnd

        if not hwnds:
            print("没有找到任何游戏窗口，停止发送")
            stop_event.set()
            break

        hwnd_game = hwnds[idx % len(hwnds)]
        idx += 1

        if activate_window(hwnd_game):
            press_ctrl_3()
        else:
            print(f"游戏窗口激活失败，句柄: {hwnd_game}")

        if last_non_game_hwnd and win32gui.IsWindow(last_non_game_hwnd) and last_non_game_hwnd != hwnd_game:
            activate_window(last_non_game_hwnd)
            #pyautogui.click()

        print(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - 已给窗口 '{win32gui.GetWindowText(hwnd_game)}' 发送 Ctrl+3，焦点还原到 {last_non_game_hwnd}")
        time.sleep(delay)

def main():
    sg.theme('DarkBlue3')
    layout = [
        [sg.Text("窗口标题关键词:"), sg.InputText("问道经典服", key="keyword", size=(20,1))],
        [sg.Text("间隔时间(秒):"), sg.InputText("60", key="delay", size=(10,1))],
        [sg.Button("启动", size=(8,1)), sg.Button("停止", size=(8,1)), sg.Button("退出", size=(8,1))],
        [sg.Text("", key="time", size=(20,1), justification='center', font=("Consolas", 12), text_color='yellow')],
        [sg.Output(size=(60, 15), font=("Consolas", 10), expand_x=True, expand_y=True)]
    ]
    window = sg.Window("MyGame", layout, element_justification='center', finalize=True, resizable=True)

    stop_event = threading.Event()
    worker_thread = None

    while True:
        event, values = window.read(timeout=1000)
        if event in (sg.WINDOW_CLOSED, "退出"):
            stop_event.set()
            break

        current_time = time.strftime("%Y-%m-%d %H:%M:%S")
        window["time"].update(current_time)

        if event == "启动" and (worker_thread is None or not worker_thread.is_alive()):
            hwnds = get_hwnds_by_title_keyword(values["keyword"])
            if not hwnds:
                print(f"未找到包含关键词 '{values['keyword']}' 的窗口")
            else:
                try:
                    delay = float(values["delay"])
                except ValueError:
                    print("间隔时间必须是数字，已使用默认60秒")
                    delay = 60
                stop_event.clear()
                worker_thread = threading.Thread(target=loop_send,
                                                 args=(hwnds, stop_event, delay, values["keyword"]),
                                                 daemon=True)
                worker_thread.start()
                print(f"找到 {len(hwnds)} 个窗口，开始循环，每 {delay} 秒给每个窗口发送一次 Ctrl+3")

        if event == "停止":
            stop_event.set()
            print("脚本已停止")

    window.close()

if __name__ == "__main__":
    main()
