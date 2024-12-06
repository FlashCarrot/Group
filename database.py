#database.py

import os
import sys
from dotenv import load_dotenv, find_dotenv
#import cx_Oracle 
import oracledb #pip install oracledb
import sshtunnel #pip install sshtunnel

class OracleDB(oracledb.Connection): 
    

    #we will do this in class
    # Or you can try on your own - create a .env file and copy all credentials to it
    load_dotenv()
    NYU_SERVER = os.environ.get('NYU_SERVER')
    local_port = int(os.environ.get('local_port'))
    ssh_username = os.environ.get('ssh_username')
    ssh_password = os.environ.get('ssh_password')
    remote_port = int(os.environ.get('remote_port'))
    remote_address = os.environ.get('remote_address')
    SID = os.environ.get('SID')
    db_username = os.environ.get('db_username')
    db_password = os.environ.get('db_password')

    dsn_tns = oracledb.makedsn(remote_address, local_port, SID)
 
    #create the ssh tunnel handler
    server = sshtunnel.SSHTunnelForwarder(NYU_SERVER,
                                        ssh_username=ssh_username,
                                        ssh_password=ssh_password,
                                        remote_bind_address=(remote_address, remote_port),
                                        local_bind_address=('', local_port))
 
 
    def __init__(self):
        #start ssh tunnel
        self.server.start() 
        super().__init__(user = self.db_username, password = self.db_password, dsn =self.dsn_tns)

    def get_connection(self):
        return self

    #when connection ends - stop ssh tunnel
    def __exit__(self, exc_type, exc_value, traceback):
        print("existing....closing connection and ssh server")
        self.server.stop()
        


        
if __name__ == "__main__":
    with OracleDB().get_connection() as connection:
        print("DB Connected")