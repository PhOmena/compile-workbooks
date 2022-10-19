# imports
from compile_workbooks import compile_workbooks
import time

# title
print("|", "-"*40, "Workbooks Compile", "-"*40, "|")

# inputs
path = input("Type the path of the files that you wanna compile (ex: C:\\Users\\...): ").replace("/", "\\")
file_name = input("Name of the final file (ex: compiled): ")

# treating the path
path = path + "\\" if path.endswith("\\") else path

# treating the file
file_name = file_name[:-5] if ".xlsx" in file_name else file_name

# function
compile_workbooks(path, file_name + ".xlsx")

# show infos
print(f"Final filename: {file_name}" + ".xlsx")

# end title
print("|", "-"*98, "|")

# time before terminal gets close
time.sleep(10)

