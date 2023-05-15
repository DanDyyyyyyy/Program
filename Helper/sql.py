import cx_Oracle
import sys
import os


try:
    if sys.platform.startswith("###"):
        lib_dir = os.path.join(os.environ.get("HOME"), "###",
                               "###")
        cx_Oracle.init_oracle_client(lib_dir=lib_dir)
    elif sys.platform.startswith("win32"):
        lib_dir=r"C:\###"
        cx_Oracle.init_oracle_client(lib_dir=lib_dir)
except Exception as err:
    print("Whoops!")
    print(err);
    sys.exit(1);