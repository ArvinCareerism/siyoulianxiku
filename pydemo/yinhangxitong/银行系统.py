

'''
用户
类名：User
属性：姓名    身份证号    电话号
行为： 


银行
类名：bank
属性：用户列表   提款机
行为：


卡
类名：Card
属性：卡号   密码     余额
行为：


提款机
类名：ATM
属性：用户字典
行为：开户   查询   取款    存款    转账    改密    锁定
      解锁   补卡   销户    退出
	  
界面
类名：View
属性：
行为： 管理员界面    管理员登录    系统功能界面
	  
'''

from view import View
from atm import ATM
import pickle
import time
import os

def main():

	filepath = os.path.join(os.getcwd(),"allusers.txt")
	
	f = open(filepath,"rb")
	allUsers = pickle.load(f)
	
	atm = ATM(allUsers)
	view = View()
	view.printAdminView()

	if view.adminOption():
		return -1
		
	while True:
		view.printsysFunctionView()
		option = input("请输入您的操作：")
		if option == '1':
			atm.createUser()
		elif option == '2':
			print("查询")
			atm.searchUserInfo()
		elif option == '3':
			print("取款")
			atm.getMoney()
		elif option == '4':
			print("存储")
		elif option == '5':
			print("转账")
		elif option == '6':
			print("改密")
		elif option == '7':
			print("锁定")
			atm.lockUser()
		elif option == '8':
			print("解锁")
			atm.unlockUser()
		elif option == '9':
			print("补卡")
		elif option == '0':
			print("销户")
		elif option == 't':
			print("退出")
			if not view.adminOption():
				
				#print(filepath)
				f = open(filepath,"wb")
				pickle.dump(atm.allUsers,f)
				f.close()
				
				return -1
				
		time.sleep(1)










if __name__ == '__main__':
	main()




















