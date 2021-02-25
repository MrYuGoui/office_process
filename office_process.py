import os
import win32com.client


class DesignFile():
    def __init__(self, name):
        self.name = name
        self.wdApp = win32com.client.gencache.EnsureDispatch('Word.Application')
        self.doc = self.wdApp.Documents.Open(os.path.join(os.getcwd(), "模板.docx"))
        self.wdApp.Selection.Find.ClearFormatting()
        self.wdApp.Selection.Find.Replacement.ClearFormatting()

    def process_word(self):
        # 替换文字
        self.wdApp.Selection.Find.Execute(
            FindText="某某", MatchCase=False, MatchWholeWord=False,MatchWildcards=False,MatchSoundsLike=False,
            MatchAllWordForms=False, Forward=True, Wrap=1,Format=True,ReplaceWith=self.name, Replace=2)

        # 插入拓扑图，需要在word相应位置插入标签“AddImage”
        RangeImage = self.wdApp.ActiveDocument.Bookmarks("AddImage").Range
        self.wdApp.Selection.InlineShapes.AddPicture(FileName=os.path.join(os.getcwd(), "模板.jpg"), LinkToFile=False, SaveWithDocument=True, Range=RangeImage)

        # 替换页眉
        # self.wdApp.ActiveDocument.Sections(3).Headers(win32com.client.constants.wdHeaderFooterPrimary).Range.Find.ClearFormatting()
        # self.wdApp.ActiveDocument.Sections(3).Headers(win32com.client.constants.wdHeaderFooterPrimary).Range.Find.Replacement.ClearFormatting()
        # self.wdApp.ActiveDocument.Sections(3).Headers(win32com.client.constants.wdHeaderFooterPrimary).Range.Find.Execute("页眉", 0, 0, 0, 0, 0, 1, 1, 1, "", 2)

        # 另存为doc和pdf
        print(self.name)
        self.doc.SaveAs(os.path.join(os.getcwd(), f"{self.name}.docx"))
        # self.doc.ExportAsFixedFormat(os.path.join(os.getcwd(), f"{self.name}.pdf"), win32com.client.constants.wdExportFormatPDF,
        #                         Item=win32com.client.constants.wdExportDocumentWithMarkup,
        #                         CreateBookmarks=win32com.client.constants.wdExportCreateHeadingBookmarks)
        self.wdApp.Quit()

    def start_func(self):
        self.process_word()


def process_excel():
    EachFileList = []
    xlApp = win32com.client.Dispatch('Excel.Application')
    # 选择excel文件，选中sheet
    xlBook = xlApp.Workbooks.Open(os.path.join(os.getcwd(), "模板.xlsx"))
    xlSheet = xlBook.Worksheets("模板")
    # 获取当前sheet下有效的行数（-1是排除了第一行表头）
    Num = xlSheet.usedrange.rows.count - 1
    print("本次共有", Num, "个文件处理:\n")
    for i in range(Num):
        EachFileDict = {}
        EachFileDict["姓名"] = xlSheet.Cells(i + 2, 1).Value
        # 若需要在表格赋值：
        # xlSheet.Cells(i + 2, 2).Value = "处理"
        EachFileList.append(EachFileDict)

    # 若需要保存及输出PDF
    # xlBook.SaveAs(f"模板.xlsx")
    # xlBook.Worksheets(['模板']).Select()
    # xlApp.ActiveSheet.ExportAsFixedFormat(0,f"表格.pdf")
    xlApp.Quit()
    return EachFileList


if __name__ == '__main__':
    pe_res = process_excel()
    for each in pe_res:
        df = DesignFile(each["姓名"])
        df.start_func()
