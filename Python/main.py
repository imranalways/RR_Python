# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import webbrowser


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('Imran')

# webbrowser.open_new_tab("http://www.google.com")
chromedir= "C:/Program Files/Google/Chrome/Application/chrome.exe %s"
webbrowser.get(chromedir).open("https://www.youtube.com/")
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
