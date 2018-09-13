import io
import sys


# Change the default encoding of standard output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')

shooter01_01 = open("friends0101.txt", "r+", encoding='utf8')
new_shooter01_01 = open("new_friends0101.txt", "w+", encoding='utf8')
line_num = 0
dialogue_switch__AB = False
while True:
    temp_str = shooter01_01.readline()
    str1 = temp_str.split('}')
    if dialogue_switch__AB:
        new_shooter01_01.write("Dialog A: " + str1[-1] + '\n')
    else:
        new_shooter01_01.write("Dialog B: " + str1[-1] + '\n')
    dialogue_switch__AB = ~dialogue_switch__AB
    if shooter01_01.readline() == '':
        break
shooter01_01.close()
new_shooter01_01.close()
