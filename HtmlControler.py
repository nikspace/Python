# coding=utf-8
from HTMLParser import HTMLParser

#继承自HTMLParser 实现各个回调函数
class HtmlControler(HTMLParser):

    _title_text = False  
    _title_value = ''  
    def handle_starttag(self,tag,attr):  
        if tag == 'title':  
            self._title_text = True  
            #print (dict(attr))  
              
    def handle_endtag(self,tag):  
        if tag == 'title':  
            self._title_text = False  
              
    def handle_data(self,data):  
        if self._title_text:  
            self._title_value = data  
    def get_title_text(self,html):
        self.feed(html)
        return self._title_value;
    
if __name__ == "__main__":           
    parser = HtmlControler()
    print parser.get_title_text('<html><head><title>test nik title</title></head><body><p>Some <a href=\"#\">html</a> tutorial...<br>END</p></body></html>')
    parser.close()