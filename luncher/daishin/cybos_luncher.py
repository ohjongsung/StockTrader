import pywinauto.timings
import pyautogui
import time


def cybos_login():
    print(pywinauto.__version__)

    # 비번
    pw = '1234'
    cert = '1234'

    # 파일에서 비번들 로드할경우
    with open('../.ignores/pw.txt') as f:
        items = list(f.readlines())
        pw = items[1].strip()
        cert = items[2].strip()

    # 어플리케이션 실행
    cybos_plus_app = pywinauto.Application()
    cybos_plus_app.start(r'D:\DAISHIN\STARTER\ncStarter.exe /prj:cp')
    print('cp load done')

    # wait
    time.sleep(1)

    # 엔터
    pyautogui.typewrite('\n', interval=0.1)
    print('enter done')

    # 다이얼로그
    def ret_wind():
        return cybos_plus_app.connect(title='CYBOS Starter').Dialog

    dlg = pywinauto.timings.wait_until_passes(10, 0.5, ret_wind)
    print('done dlg')

    # pass edit
    pass_edit = dlg.Edit2
    pass_edit.set_focus()
    pass_edit.type_keys(pw)
    print('done pw')

    # cert edit
    cert_edit = dlg.Edit3
    cert_edit.set_focus()
    cert_edit.type_keys(cert)
    print('done cert')

    # login
    btn = dlg.Button
    btn.click()
    print('done login click')


if __name__ == "__main__":
    cybos_login()
