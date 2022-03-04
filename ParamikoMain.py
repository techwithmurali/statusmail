import paramiko
import paramiko_globals as globals
import os,sys

class SSHConnection(object):
    def __init__(self, p_hostname, p_username, p_password, port=22):
        self.ssh_client=paramiko.SSHClient()
        self.ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        self.ssh_client.connect(hostname=p_hostname,username=p_username,password=p_password)
        self.ftp_client=self.ssh_client.open_sftp()

    def close(self):
        self.ftp_client.close() 
        self.ssh_client.close()

    def local_to_server(self,localpath,serverpath,lstFiles):
        try:
            for file in lstFiles:
                self.ftp_client.put(os.path.join(localpath,file),(serverpath + '/' + file))
        finally:
            self.close()        

    def server_to_local(self,localpath,serverpath,lstFiles):
        try:
            for file in lstFiles:
                self.ftp_client.get((serverpath + '/' + file),os.path.join(localpath,file))
        finally:
            self.close() 

    def exec_command(self,lstCommand):
        try:
            for command in lstCommand:
                (stdin,stdout,stderr) = self.ssh_client.exec_command(command)
                cmd_output = stdout.read()
                err  = stderr.read().decode()
                print(' command : {} - Output:  {}'.format(command,cmd_output))
                with open(globals.output_file,'w+') as fp:
                    fp.write('Command - {} - Output: {} '.format(command,str(cmd_output)))
                    if err:
                        print(err)
                        fp.write('Command - {} - Error: {} '.format(command,str(err)))
        finally:
            self.close() 


def local_to_server(localpath, destpath,lstFiles):
    ftpclient =  SSHConnection(globals.hostname,globals.username,globals.password)
    ftpclient.local_to_server(localpath, destpath,lstFiles)

def server_to_local(localpath, destpath,lstFiles):
    ftpclient =  SSHConnection(globals.hostname,globals.username,globals.password)
    ftpclient.server_to_local(localpath, destpath,lstFiles)
 
if __name__ == "__main__":
    if sys.argv[1] == 'Local_Server':
        local_to_server(globals.localpath,globals.serverpath,globals.lst_Files )
    elif sys.argv[1] == 'Server_Local':
        server_to_local(globals.localpath_dest, globals.serverpath,globals.lst_Files)
    print('completed')
