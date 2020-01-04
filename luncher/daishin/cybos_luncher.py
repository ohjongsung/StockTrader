

from luncher.daishin.creon import Creon


def cybos_login():
    creon = Creon()
    creon.connect('id', 'pwd', 'pwdcert')


if __name__ == "__main__":
    cybos_login()
