import os

envvars = ['EWS_USERNAME',
           'EWS_PASSWORD',
           'DB_HOST',
           'DB_NAME',
           'DB_USER',
           'DB_PASSWORD',
           'SMTP_URL',
           'SMTP_USER',
           'SMTP_PASSWORD']


def print_envvars():
    for envvar in envvars:
        print(f'{envvar} is set to {os.environ.get(envvar)}')


def check_exist():
    for envvar in envvars:
        if not os.environ.get(envvar):
            raise OSError(f'Missing required environment variable {envvar}')
    return True


if __name__ == '__main__':
    if check_exist():
        print("All required environment variables set")
