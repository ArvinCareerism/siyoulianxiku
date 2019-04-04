import os


def getAllDirDE(path):
	stack = []
	stack.append(path)
	
	while len(stack) != 0:
		dirPath = stack.pop()
		filesList = os.listdir(dirPath)
		
		for fileName in filesList:
			fileAbsPath = os.path.join(dirPath,fileName)
			if os.path.isdir(fileAbsPath):
				print("目录: " + fileName)
				stack.append(fileAbsPath)
			else:
				print("普通： " + fileName)
				
getAllDirDE(r"F:\未分类")