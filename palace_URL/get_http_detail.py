import requests
import io
import sys

# Change the default encoding of standard output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')

# Real time check these objects
url_enter = 'http://hdzx.xicp.cn:56751/sjbmxt/wap.asp'
url_register = 'http://hdzx.xicp.cn:56751/sjbmxt/banjilist.asp'


def get_http_text(url):
    res = requests.get(url)
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


# Get the string local source URL enter/register
if __name__ == "__main__":
    fout_enter = open('get\\source_url_enter', "w+", encoding='utf8')
    fout_register = open('get\\source_url_register', "w+", encoding='utf8')
    fout_enter.write(clean_up_str(get_http_text(url_enter)))
    fout_register.write(clean_up_str(get_http_text(url_register)))
