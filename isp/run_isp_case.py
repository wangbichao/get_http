import sys
import os


base_str = "./tserverlite -d=1 -v=debug -test="


def help():
    help_str = "\n\n"
    help_str += "run_isp_case.py is A simple command line written in python.The current version is 1.0.\n"
### "Since too many arguments are not remembered when running the arguments,this command line is used to call the arguments.\n"
    help_str = "\n"
    help_str += "Usage: python run_isp_case.py [options] \n"
    help_str += "For example: \n"
    help_str += "    python run_isp_case.py isp3.1-2 -lcd -skip=1 \n"
    help_str += "\n\n"
    help_str += "Options: \n"
    help_str += "    -l  Save all command LOG to log.txt.\n"
    help_str += "    -c  Clear ISP memory after allocation.\n"
    help_str += "    -d  Dump DMA output frame buffer to file. File format is determined by DMA output format internally.\n"
    help_str += "    -V  Check and verify DMA FIFO overflow.\n"
    help_str += "    -R  File to save all ISP register access events.\n"
    help_str += "    -skip  The NUM of frames to be abandoned.\n"
    help_str += "\n\n"
    print(help_str)


def creat_case_cmd():
    if "." in sys.argv[1]:
        if "G" in sys.argv[1]:
            run_cmd = "gdbserver :3344 " + base_str + '"' + sys.argv[1][1:] + '"'
        else:
            run_cmd = base_str + '"' + sys.argv[1] + '"'
        if "l" in sys.argv[2]:
            run_cmd += " -log"
        if "c" in sys.argv[2]:
            run_cmd += " -clearIspMemory"
        if "d" in sys.argv[2]:
            run_cmd += " -dumpIspResult"
        if "V" in sys.argv[2]:
            run_cmd += " -checkDMAFIFO"
        if "R" in sys.argv[2]:
            run_cmd += " -logIspRegisterAccess"
        if len(sys.argv) > 3:
            if "skip" in sys.argv[3]:
                skip_frame = sys.argv[3].split('=')
                run_cmd = run_cmd + " -abandonIspFrames=" + skip_frame[-1] + " "
        if len(sys.argv) > 4:
            run_cmd += ' '.join(sys.argv[4:])
        return run_cmd
    else:
        print("!!!No input case Num!!!")
        help()
        return("arguments error.")


if __name__ == "__main__":
    if len(sys.argv) == 1:
        help()
    else:
        run = creat_case_cmd()
        print(run)
        os.system(run)
