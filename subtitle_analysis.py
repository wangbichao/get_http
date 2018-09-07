import io
import sys
import re


# Change the default encoding of standard output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf8')

shooter01_01 = open("Shooter01_01.txt", "r+", encoding='utf8')
new_shooter01_01 = open("new_shooter01_01.txt", "w+", encoding='utf8')
line_num = 0
dialogue_switch__AB = False
new_shooter01_01.write("Dialog A: ")
while True:

    temp_str = shooter01_01.readline()
    if temp_str.find("00:") == -1:
            line_num = line_num+1
            if (line_num % 2) == 0:
                if re.findall('\.$', temp_str) or\
                 re.findall('\?$', temp_str) or\
                 re.findall('\!$', temp_str):
                    if dialogue_switch__AB:
                        new_shooter01_01.write(temp_str + '\n' + "Dialog A: ")
                    else:
                        new_shooter01_01.write(temp_str + '\n' + "Dialog B: ")
                    dialogue_switch__AB = ~dialogue_switch__AB
                else:
                    new_shooter01_01.write(temp_str)
    if shooter01_01.readline() == '':
        break
shooter01_01.close()
new_shooter01_01.close()

