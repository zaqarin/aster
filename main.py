import paramiko

def ssh_file_transfer(address, username, password, remote_path, local_path):
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    client.connect(address, username=username, password=password)

    ftp = client.open_sftp()
    ftp.get(remote_path, local_path)
    ftp.close()
    client.close()

# Example usage
ssh_file_transfer("host", "username", "password", "path to ade file")
