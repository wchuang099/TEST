#pyinstaller --onefile --windowed --hidden-import=win32timezone MyGame.py

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

def get_hwnds_by_process_name(target_process_name):
    hwnds = []
    def enum_handler(hwnd, lParam):
        if win32gui.IsWindowVisible(hwnd):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            try:
                proc = psutil.Process(pid)
                if proc.name().lower() == target_process_name.lower():
                    hwnds.append(hwnd)
            except Exception:
                pass
        return True
    win32gui.EnumWindows(enum_handler, None)
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

def press_ctrl_3():
    VK_CONTROL = 0x11
    VK_3 = 0x33
    win32api.keybd_event(VK_CONTROL, 0, 0, 0)
    win32api.keybd_event(VK_3, 0, 0, 0)
    time.sleep(0.05)
    win32api.keybd_event(VK_3, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)

def press_ctrl_tab():
    VK_CONTROL = 0x11
    VK_TAB = 0x09
    win32api.keybd_event(VK_CONTROL, 0, 0, 0)
    win32api.keybd_event(VK_TAB, 0, 0, 0)
    time.sleep(0.05)
    win32api.keybd_event(VK_TAB, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)

def loop_send(hwnd_game, stop_event, delay, game_process_name):
    last_non_game_hwnd = None

    def is_game_window(hwnd):
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            proc = psutil.Process(pid)
            return proc.name().lower() == game_process_name.lower()
        except Exception:
            return False

    while not stop_event.is_set():
        curr_hwnd = win32gui.GetForegroundWindow()

        if curr_hwnd and win32gui.IsWindow(curr_hwnd) and not is_game_window(curr_hwnd):
            last_non_game_hwnd = curr_hwnd

        if activate_window(hwnd_game):
            press_ctrl_3()
            time.sleep(1)
            press_ctrl_tab()
        else:
            print("游戏窗口激活失败，停止发送")
            stop_event.set()
            break

        if last_non_game_hwnd and win32gui.IsWindow(last_non_game_hwnd) and last_non_game_hwnd != hwnd_game:
            activate_window(last_non_game_hwnd)
            pyautogui.click()

        print(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - 已发送 Ctrl+3 和 Ctrl+Tab，焦点还原到 {last_non_game_hwnd}")
        time.sleep(delay)

def main():
    sg.theme('DarkBlue3')
    layout = [
        [sg.Text("目标进程名:"), sg.InputText("AtTab.exe", key="proc_name", size=(20,1))],
        [sg.Text("间隔时间(秒):"), sg.InputText("60", key="delay", size=(10,1))],
        [sg.Button("启动", size=(8,1)), sg.Button("停止", size=(8,1)), sg.Button("退出", size=(8,1))],
        [sg.Text("", key="time", size=(20,1), justification='center', font=("Consolas", 12), text_color='yellow')],
        [sg.Output(size=(60, 15), font=("Consolas", 10), expand_x=True, expand_y=True)]
    ]
    window = sg.Window("MyGame", layout, element_justification='center', finalize=True, resizable=True)

    stop_event = threading.Event()
    worker_thread = None

    while True:
        event, values = window.read(timeout=1000)  # 每秒刷新一次
        if event in (sg.WINDOW_CLOSED, "退出"):
            stop_event.set()
            break

        # 更新时间显示
        current_time = time.strftime("%Y-%m-%d %H:%M:%S")
        window["time"].update(current_time)

        if event == "启动" and (worker_thread is None or not worker_thread.is_alive()):
            hwnds = get_hwnds_by_process_name(values["proc_name"])
            if not hwnds:
                print(f"未找到进程名为 {values['proc_name']} 的窗口")
            else:
                hwnd_game = hwnds[0]
                try:
                    delay = float(values["delay"])
                except ValueError:
                    print("间隔时间必须是数字，已使用默认60秒")
                    delay = 60
                stop_event.clear()
                worker_thread = threading.Thread(target=loop_send,
                                                 args=(hwnd_game, stop_event, delay, values["proc_name"]),
                                                 daemon=True)
                worker_thread.start()
                print(f"找到窗口句柄: {hwnd_game}，开始循环，每 {delay} 秒发送一次 Ctrl+3 和 Ctrl+Tab")

        if event == "停止":
            stop_event.set()
            print("脚本已停止")

    window.close()

if __name__ == "__main__":
    main()
