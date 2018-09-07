import requests
import io
import sys
import pygame
import time

# Change the default encoding of standard output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')

# Real time check these objects
# url_enter = 'http://hdzx.xicp.cn:56751/sjbmxt/wap1.asp'
# url_register = 'http://hdzx.xicp.cn:56751/sjbmxt/banjilist2.asp'
url_enter = 'http://hdzx.xicp.cn:56751/sjbmxt/wap.asp'
url_register = 'http://hdzx.xicp.cn:56751/sjbmxt/banjilist.asp'

# Get the string local source URL enter/register
fout_s_url_enter = open("source\\source_url_enter", "r+", encoding='utf8')
fout_s_url_register = open("source\\source_url_register", "r+", encoding='utf8')
S_url_enter = fout_s_url_enter.read(2210)
S_url_register = fout_s_url_register.read()


def get_http_text(url):
    try:
        res = requests.get(url, timeout=10)
        res.encoding = 'utf-8'
        return res.text
    except res.HTTPError as e:
        return e.code
    except res.URLError as e:
        return e.reason
    else:
        res.encoding = 'utf-8'
        return res.text


def clean_up_str(input_str):
    output_str = '''input_str.strip().
        replace(' ', '').
        replace('\n', '').
        replace('\t', '').
        replace('\r', '').
        strip()'''
    return output_str


if __name__ == "__main__":
    file = r'source\\xiyouji_song.mp3'
    pygame.mixer.init()
    track = pygame.mixer.music.load(file)
    while True:
        time.sleep(5)
        # loop real time check URL
        S_real_time_enter = get_http_text(url_enter)
        S_real_time_enter = clean_up_str(S_real_time_enter)
        time.sleep(5)
        S_real_time_register = get_http_text(url_register)
        S_real_time_register = clean_up_str(S_real_time_register)
        if S_url_enter in S_real_time_enter:
            print("same enter")
        else:
            print("diff enter")
#            pygame.mixer.music.play()
#            time.sleep(20)
#            pygame.mixer.music.stop()
#            pygame.mixer.music.unpause()
            print(S_url_enter)
            print("diff enter")
            print(S_real_time_enter)
            print("diff enter")
        if S_url_register in S_real_time_register:
            print("same register")
        else:
            print("diff register")
#            pygame.mixer.music.play()
#            time.sleep(20)
#            pygame.mixer.music.stop()
#            pygame.mixer.music.unpause()
            print("diff register")
            print(S_url_register)
            print("diff register")
            print(S_real_time_register)
            print("diff register")
            break
