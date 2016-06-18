# coding=utf-8
import sys,os  
import ConfigParser  


'''
[http_maker]
user = xxx
password = xxx

[test_writer]
test = xxx
'''
class ConfigManager:  
    def __init__(self, path):  
        self.path = path  
        self.cf = ConfigParser.ConfigParser()  
        self.cf.read(self.path)  
    '''
    @return: return '' if not found
    '''
    def get(self, field, key):  
        result = ""  
        try:  
            result = self.cf.get(field, key)  
        except:  
            result = ""  
        return result  
    def set(self, field, key, value):  
        try:  
            self.cf.set(field, key, value)  
            self.cf.write(open(self.path,'w'))  
        except:  
            raise 
        return True  
              
              
  

if __name__ == "__main__":  
  c = ConfigManager(r"c:\http_maker_conf.ini") 
  print c.get('http_maker', 'user')
  print c.get('http_maker','password')
  c.set('test_writer', 'test', 'testxxx')