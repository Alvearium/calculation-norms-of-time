from ftplib import FTP

# server = socket(
#     AF_INET, SOCK_STREAM
# )
# server.bind(
#     ('89.111.137.139', 1024)
# )
# server.listen(2)
# user, addr = server.accept()
# print(f"CONNECTED:\n{user},\n{addr}")
def send_file(file):
    server = '89.111.137.139'
    username = 'ftpuser'
    password = 'admin'

    ftp = FTP(server)
    ftp.connect(server, 22)
    ftp.login(username, password)
    ftp.cwd('./root/user_files/')

    my_file = open(file, 'wb')
    ftp.retrbinary('RETR ' + my_file.name, my_file.write, 1024)
    ftp.quit()
    my_file.close()