#!/usr/bin/python
import sys
import os


base_str = "./tserverlite -skip_iommu_check -zfb -nogc -skip_consoleblank_check -no_toollib -load_ucode=off -gpu_copyengine=sdma -v=debug -dcedebug=1 -log -d=1 -skiptextbuffering=1 -use_kernel_driver=0 -test="


def help():
    help_str = "\n\n"
    help_str += "dcn.py is A simple command line written in python.The current version is 1.0.\n"
### "Since too many arguments are not remembered when running the arguments,this command line is used to call the arguments.\n"
    help_str = "\n"
    help_str += "Usage: python dcn.py [options] \n"
    help_str += "For example: \n"
    help_str += "    python dcn.py -t/-f DIO2.1/file \n"
    help_str += "\n\n"
    help_str += "Options: \n"
    help_str += "    -t  execute the case.\n"
    help_str += "    -f  execute all cases on the file.\n"
    help_str += "\n\n"
    print(help_str)


def creat_case_cmd():
    if len(sys.argv) > 2:
        if "p" in sys.argv[1]:
            run_cmd = base_str + ' '.join(sys.argv[2:])
            print(run_cmd)
        if "t" in sys.argv[1]:
            if "." in sys.argv[2]:
                run_cmd = base_str + ' '.join(sys.argv[2:])
                os.system(run_cmd)
            else:
                print("!!!No input case Num!!!")
                help()
                return("arguments error.")
        if "f" in sys.argv[1]:
            regression_fail = open(sys.argv[2], "r+")
            for run_test in regression_fail:
                run_test = run_test.strip('\n')
                run_cmd = base_str + run_test + " -v=error " + ' '.join(sys.argv[3:])
#                print(run_cmd)
                os.system(run_cmd)
    else:
        print("!!!No input case Num!!!")
        help()
        return("arguments error.")


if __name__ == "__main__":
    if len(sys.argv) == 1:
        help()
    else:
        creat_case_cmd()
#        print(run)
#        os.system(run)
