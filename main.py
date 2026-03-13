from scriptv1 import run_v1
from scriptv2 import run_v2

if __name__ == "__main__":
    version = input("1 = Excel | 2 = URL: ")

    if version == "1":
        run_v1()
    elif version == "2":
        run_v2()
    else:
        print("Versão inválida")