from docx import Document
from docx.shared import Pt
from cv2 import imdecode, imwrite
from numpy import frombuffer, uint8
from json import dumps
from os import rename, remove, mkdir
from os.path import exists, join
from sys import argv

# 字体对应的磅数，可以查看 https://zhidao.baidu.com/question/550765269.html
fontDict = {
    "一号": 26,
    "小二": 18,
    "三号": 16,
    "小三": 15,
    "四号": 14,
    "普通文本": 11
}


class WordExtractor:
    def __init__(self, word_name):
        """
        这个类用于提取出word文档里面的题目，注意这个word文档最好是从飞书文档转换过来的
        按照一定规则的word文档
        :param word_name: word文档的文件路径
        """
        # 这个Document输入的是word文档的文件路径，他会解析这个word文档，用于我们后面提取这个word文档的题目
        self.word = Document(word_name)
        # 这个字典是用于把字号的具体数值大小跟字号的名字对应起来，如{"330200":"一号"}
        self.fontLevelNum = {Pt(value): key for key, value in fontDict.items()}
        # 这个列表用于存放word文档里面的所有图片，图片的格式是三维数组形式
        self.imgsList = self.getImgs()

    def getImgs(self) -> list:
        """
        按照从上往下的顺序提取出word文档里面的所有图片，并用opencv将其读取成三维数组的形式
        :return: 装着所有图片数组的列表
        """
        keyList = sorted(self.word.part._rels.keys(), key=lambda x: int(x[3:]))
        imgsList = []
        for key in keyList:
            if "media" in self.word.part._rels[key].target_ref:
                array = frombuffer(self.word.part._rels[key].target_part.blob, dtype=uint8)
                img = imdecode(array, 1)
                imgsList.append(img)
        return imgsList

    def saveImg(self, saveName: str):
        """
        保存图片到当前路径下的saveImgs文件夹内，如果这个文件夹不存在，会自动创建
        :param saveName: 保存的图片的图片名字
        :return: None
        """
        # 如果saveImgs这个文件夹不存在就创建
        if exists(self.header) is False:
            mkdir(self.header)
        # 因为.imwrite保存图片的图片名字不能有中文，所以先保存成temp.png
        # 保存图片列表里面的第一张图片，并把这张图片从列表里面删除
        imwrite("temp.png", self.imgsList.pop(0))
        # 把上面保存的temp.png图片名字更改为它真正应该是的名字
        try:
            rename("temp.png", join(self.header, saveName))
        except FileExistsError:
            remove(join(self.header, saveName))
            rename("temp.png", join(self.header, saveName))

    def formatProblem(self, problem: list, level: list) -> dict:
        """
        输入包含一道题目的列表，列表里面的元素是字符串
        这个函数就是根据这个列表来提取出这个题目的所有信息
        并将这些信息按照标准的格式保存到字典
        :param problem: 输入的列表，里面元素是字符串
        :param level: 列表，里面存放着这个题目所在的位置（在哪个父节点下）
        :return: 返回整理好的题目，类型是字典
        """
        # 创建一个空字典，用于存放标准化后的题目
        question = {}
        # 题目默认类型是填空题，如果有选项的话就是选择题
        questionType = "填空题"
        # 图片的序号，因为一道题可能有多个图片，区分开这些图片就用这个序号
        imgNumber = 0
        for index, Str in enumerate(problem):
            Str = Str.replace("：", ":")
            if Str.startswith("（"):
                # 如果进入这个if语句，表示这个是题目的开头，因为题目开头一般就是（1）运行下面程序......
                question["questionTitle"] = {
                    "title": Str,
                    "imgName": []
                }
            elif Str.endswith(".png"):
                # 图片的名字：他的所有父节点的名称加上序号，如"一、编程基础+（一）变量+1、创建变量+（1）+0.png"
                imgName = "+".join(level) + "+" + str(imgNumber) + ".png"
                imgNumber += 1
                # 看一下上一个循环的字符串，来判断这个图片是属于题目的问题部分还是属于选项里面
                lastStr = problem[index - 1]
                if lastStr in ["A:", "B:", "C:", "D:"]:
                    # 这张图片属于选项部分
                    if lastStr == 'A:':
                        questionType = "选择题"
                        question['option'] = {
                            lastStr: imgName
                        }
                    else:
                        question['option'][lastStr] = imgName
                elif lastStr.startswith("（"):
                    # 这张图片属于题目的问题部分而且这个题目的问题部分只有一张图片
                    question["questionTitle"]['imgName'].append(imgName)
                elif lastStr.endswith(".png"):
                    # 这张图片属于题目的问题部分而且这个题目的问题部分有多张图片，循环遍历这个题目的问题部分的所有图片
                    count = 2
                    while True:
                        if index < count:
                            raise ValueError("这道题的图片放的位置有问题")
                        if problem[index - count].startswith("（"):
                            break
                        count += 1
                    question["questionTitle"]['imgName'].append(imgName)
                # 将图片列表中的第一张图片抽取出来保存
                self.saveImg(imgName)
            elif Str[:2] in ["A:", "B:", "C:", "D:"] and Str not in ["A:", "B:", "C:", "D:"]:
                # 如果这个题目是选择题而且选项不是图片而是文字就进入这个语句
                if Str[:2] == 'A:':
                    questionType = "选择题"
                    question['option'] = {
                        Str[:2]: Str[2:]
                    }
                else:
                    question['option'][Str[:2]] = Str[2:]

            elif index == len(problem) - 1 and "答案" in Str:
                # 如果这个字符串在题目的末尾而且这个字符串里面又答案两个字就表示这个是答案部分
                question['answer'] = Str[3:]
        # 根据上面循环的结果，题目的类型已经确定下来了，把类型写入字典
        question["questionType"] = questionType
        return question

    def getQuestions(self):
        """
        从word文档中提取出所有题目，最终保存为json文件
        :return: None
        """
        # 定义一个列表用来装n级标题，比如当前循环到的题目的一级标题、二级标题等等
        level = ["", "", "", "", ""]
        # 定义一个字典用来存储所有的题目
        questions = {}
        # 定义一个列表，这个列表的作用是作为中间变量，把单个题目装进去然后放到formatProblem函数里面解析
        tempList = []
        for p in self.word.paragraphs:
            # 遍历到的这一段的字体大小，根据这个字体大小可以判断是正文还是n级标题，如果是None表示是图片
            fontSize = p.runs[0].font.size
            # 遍历到的这一段字符串，比如"一、编程基础"或"（一）变量"等
            content = p.text.replace("：", ":").strip()
            if "计算机视觉" in content:
                break

            if fontSize is None and content == "":
                # 如果是图片，fontSize是空而且content没有内容
                tempList.append("+".join(level) + ".png")

            elif '一号' == self.fontLevelNum[fontSize]:
                # 文档的文件名，暂时没有用处
                self.header = content
                level[0] = content

            elif '小二' == self.fontLevelNum[fontSize]:
                # 一级标题，如一、编程基础
                level[1] = content
                questions[content] = {}
            elif '三号' == self.fontLevelNum[fontSize]:
                # 二级标题，如（一）变量
                level[2] = content
                questions[level[1]][content] = {}
            elif '小三' == self.fontLevelNum[fontSize]:
                # 三级标题，如1、创建变量
                level[3] = content
                questions[level[1]][level[2]][content] = []
            elif '四号' == self.fontLevelNum[fontSize]:
                # 四级标题，就是具体的题目了，这里是题目的开头
                if "（" in content:
                    tempList.append(content)
                    level[4] = content[:content.find("）") + 1]
            elif "普通文本" == self.fontLevelNum[fontSize]:
                # 普通文本可能是答案也可能是文本选项，比如"A: 可能相同"
                tempList.append(content)
                if "答案" in content:
                    questions[level[1]][level[2]][level[3]].append(self.formatProblem(tempList, level))
                    # 一道题解析结束，这个作为中间变量的列表就需要情况，以便于存储下一道题目
                    tempList = []
        # 用dumps把json对象转成json格式的字符串，加入indent=4比较好看，层次分明，最后一个是为了正常显示中文而不是Unicode编码
        with open(level[0] + ".json", 'w', encoding='utf-8') as file:
            file.write(dumps(questions, indent=4, ensure_ascii=False))


for file in [file for file in argv if file.endswith(".docx")]:
    print(file)
    word = WordExtractor(file)
    word.getQuestions()
    print(f"{file}文件转换成功,json文件输出为{word.header}.json文件内，图片保存在{word.header}文件夹内")
