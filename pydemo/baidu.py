from urllib import request

response = request.urlopen("http://www.baidu.com/")
fi = open("priject.txt","w")
page = fi.write(str(response.read()))
fi.close()
