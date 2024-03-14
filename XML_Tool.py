import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTextEdit, QVBoxLayout, QWidget, QPushButton, QFileDialog
import pandas as pd
from lxml import etree
from PyQt5.QtWidgets import QMessageBox

class AppWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.title = '让XML的信息无处遁形！'
        self.left = 100
        self.top = 100
        self.width = 640
        self.height = 480
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        # 搞一个文本编辑区域
        self.textEdit = QTextEdit(self)
        self.textEdit.setReadOnly(True)

        # 搞一个按钮，用于打开文件选择对话框
        self.buttonOpen = QPushButton('找到XML！', self)
        self.buttonOpen.clicked.connect(self.openFileNameDialog)

        # 搞一个按钮，用于导出结果到 Excel 文件
        self.buttonExport = QPushButton('导出到Excel', self)
        self.buttonExport.clicked.connect(self.exportToExcel)

        # 设置布局
        layout = QVBoxLayout()
        layout.addWidget(self.buttonOpen)
        layout.addWidget(self.textEdit)
        layout.addWidget(self.buttonExport)

        # 设置布局会用到的小东西
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def openFileNameDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "QFileDialog.getOpenFileName()", "",
                                                  "XML Files (*.xml);;All Files (*)", options=options)
        try:
            self.parse_xml(fileName)
        except Exception as e:
            print("XML的这里有问题:", e)

    def parse_xml(self, file_path):
        # 解析XML文件
        tree = etree.parse(file_path)
        root = tree.getroot()
        result = ""

        # 处理视频轨道信息
        video_info = self.parse_track(root, 'video')
        result += "视频轨道信息:\n"
        result += video_info

        # 处理音频轨道信息
        audio_info = self.parse_track(root, 'audio')
        result += "音频轨道信息:\n"
        result += audio_info

        self.textEdit.setText(result)
        self.result = result  # 保存结果供导出到 Excel 使用

    def parse_track(self, root, track_type):
        result = ""
        # 根据 track_type 选择不同的轨道
        track_elements = root.xpath(f'//{track_type}//clipitem')

        # 遍历轨道上的 clipitem 元素
        for clipitem in track_elements:
            name = clipitem.find('name')
            if name is None:
                continue

            is_valid = self.filter_clip(name)
            if is_valid:
                result += self.return_clip_info(name)

        return result

    def filter_clip(self, item):
        # 判断片段名是否符合规范
        # :param item: 一个包含片段名信息的对象，预期为一个具有.text属性的对象，该属性存储了片段的名称。
        # :return: 返回一个布尔值，如果片段名符合规范（即不包含任何素材文件的常见后缀），则返回True；否则返回False。
        filter_out_params = ['.mp4', '.avi', '.mov', '.mkv', '.flv', '.wmv', '.ari', '.dng', '.mxf', '.r3d', '.arw', '.jpg', '.jpeg', '.dpx', '.cine', '.raw', '.wav', '.wave', '.mp3', '.aac', '.ogg', '.flac', '.wma', '.aiff', '.au', '.m4a', '.ape']  # 属于素材的后缀，需要过滤掉这些后缀的片段名

        is_valid = True  # 默认假设片段名是有效的

        data = item.text  # 获取片段名

        # 遍历所有需要过滤的后缀，如果片段名中包含任何一种后缀，则认为片段名不符合规范
        for param in filter_out_params:
            if data.find(param) != -1:
                is_valid = True  # 找到匹配的后缀，片段名无效
                return is_valid
            else:
                is_valid = False  # 未找到匹配的后缀，片段名暂时有效

        return is_valid  # 返回最终判断结果

    def return_clip_info(self, name):
        # 用于存储结果的字符串
        result_text = ""

        # 获取大元素（假设它包含所需的信息）
        parent = name.getparent()

        # 对于每个name，尝试找到对应的子元素
        width = parent.xpath('.//width/text()')
        height = parent.xpath('.//height/text()')
        timebase = parent.xpath('.//timebase/text()')
        start = parent.xpath('.//start/text()')
        end = parent.xpath('.//end/text()')
        depth = parent.xpath('.//depth/text()')
        channelcount = parent.xpath('.//channelcount/text()')
        samplerate = parent.xpath('.//samplerate/text()')
        alphatype = parent.xpath('.//alphatype/text()')
        pproBypass = parent.xpath('.//pproBypass/text()')
        when = parent.xpath('.//when/text()')
        value = parent.xpath('.//value/text()')
        authoringApp = parent.xpath('.//parameter[@authoringApp]/@authoringApp')
        effecttype = parent.xpath('.//effecttype/text()')
        parameterid = parent.xpath('.//parameterid/text()')

        # 给时码提取更复杂的结构
        timecode_string = "无"
        timecode = parent.xpath('.//timecode/string/text()')
        if timecode:
            timecode_string = timecode[0]

        # 将信息添加到结果字符串中
        result_text += "\n"
        result_text += f"片段名称: {name.text}\n"
        result_text += "Video:\n"
        result_text += f"  素材长度: {width[0] if width else '无'}\n"
        result_text += f"  素材宽度: {height[0] if height else '无'}\n"
        result_text += f"  素材帧数: {timebase[0] if timebase else '无'}\n"
        result_text += f"  素材开始: {start[0] if start else '无'}\n"
        result_text += f"  素材结束: {end[0] if end else '无'}\n"
        result_text += f"  素材时码: {timecode_string}\n"
        result_text += f"  透明通道: {alphatype[0] if alphatype else '无'}\n"
        result_text += f"  绕过特效: {pproBypass[0] if pproBypass else '未知'}\n"
        result_text += f"  特效类型: {effecttype[0] if effecttype else '未知'}\n"
        result_text += f"  始关键帧: {when[0] if when else '未检测到关键帧'}\n"
        result_text += f"  可变参数: {value[0] if value else '未知'}\n"
        result_text += f"  参数名称: {parameterid[0] if parameterid else '未知'}\n"
        result_text += f"  交付软件: {authoringApp[0] if authoringApp else '未知'}\n"
        result_text += "Audio:\n"
        result_text += f"  音频位深: {depth[0] if depth else '无'}\n"
        result_text += f"  音频通道: {channelcount[0] if channelcount else '无'}\n"
        result_text += f"  音频采样: {samplerate[0] if samplerate else '无'}\n"
        result_text += "---\n\n"

        return result_text

    def exportToExcel(self):
        # 创建 DataFrame
        data = []
        result_lines = self.result.split("\n")
        current_dict = {}
        for line in result_lines:
            if line.startswith("片段名称:"):
                if current_dict:
                    data.append(current_dict)
                    current_dict = {}
                current_dict["片段名称"] = line.split(":")[1].strip()  # 加入片段名称信息
            elif line.strip() and ":" in line:
                key, value = line.split(":", 1)
                current_dict[key.strip()] = value.strip()
        if current_dict:
            data.append(current_dict)

        df = pd.DataFrame(data)

        # 导出到 Excel 文件
        file_path, _ = QFileDialog.getSaveFileName(self, "导出到Excel", filter="Excel文件 (*.xlsx)")
        if file_path:
            df.to_excel(file_path, index=False)
            QMessageBox.information(self, "成功", "嘿，您猜怎么着，导出成功了！")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = AppWindow()
    ex.show()
    sys.exit(app.exec_())
