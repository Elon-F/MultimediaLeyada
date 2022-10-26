import os
import subprocess

LOC = r"C:\Users\Elonf\Desktop\autoit-scripts-fixed"
COMPILER = r"C:\Program Files (x86)\AutoIt3\Aut2Exe\Aut2exe.exe"
ARGS = ["/x64"]  # /gui?


def get_file_data_versatile(file):
    codecs = ["charmap", "utf8", "utf16", "ansi"]
    # Read in the file
    for codec in codecs:
        try:
            with open(file, 'r', encoding=codec) as f:
                return f.read()
        except UnicodeEncodeError:
            continue
    print(f"ERROR ON {file}")

def write_file_data_versatile(file, data):
    codecs = ["charmap", "utf8", "utf16", "ansi"]
    # Read in the file
    for codec in codecs:
        try:
            with open(file, 'w', encoding=codec) as f:
                f.write(data)
                return
        except UnicodeEncodeError:
            continue
    print(f"ERROR ON {file}")


def relativise_paths(file, source, target):
    filedata = get_file_data_versatile(file)

    # Replace the target string
    filedata = filedata.replace(source, target)

    # Write the file out again
    write_file_data_versatile(file, filedata)


if __name__ == '__main__':
    files = []
    print("test")
    # r=root, d=directories, f = files
    for r, d, f in os.walk(LOC):
        for file in f:
            if any(ext in file for ext in ['.txt', '.bat', '.au3', '.ps1']):
                files.append(os.path.join(r, file))

    for file in files:
        print(file)
        relativise_paths(file, r"D:\autoit-scripts" "\\", ".\\")
        relativise_paths(file, r'$drive &"\autoit-scripts' "\\", "\".\\")
        if os.path.splitext(file)[1] == ".au3":
            print("Script file detected! compiling...")
            args = [COMPILER, f'/in', file, *ARGS]
            subprocess.call(args)
            print("Done! \n")
