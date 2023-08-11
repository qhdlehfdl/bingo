import win32com.client as win32
from random import *

hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
hwp.XHwpWindows.Item(0).Visible = True
hwp.Run("CharShapeBold")

# 여백 설정
hwp.HAction.GetDefault("PageSetup", hwp.HParameterSet.HSecDef.HSet)  # 액션생성
hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyClass", 24)  # 적용범위 구분. 없어도 됨
hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyTo", 3)  # 적용범위. 필수
hwp.HParameterSet.HSecDef.PageDef.LeftMargin = hwp.MiliToHwpUnit(
    5.0)  # 파라미터셋 설정
hwp.HParameterSet.HSecDef.PageDef.RightMargin = hwp.MiliToHwpUnit(5.0)
hwp.HParameterSet.HSecDef.PageDef.TopMargin = hwp.MiliToHwpUnit(5.0)
hwp.HParameterSet.HSecDef.PageDef.BottomMargin = hwp.MiliToHwpUnit(5.0)
hwp.HParameterSet.HSecDef.PageDef.HeaderLen = hwp.MiliToHwpUnit(5.0)
hwp.HParameterSet.HSecDef.PageDef.FooterLen = hwp.MiliToHwpUnit(5.0)
hwp.HParameterSet.HSecDef.PageDef.GutterLen = hwp.MiliToHwpUnit(5.0)
# 해당액션 실행(파라미터셋 적용)
hwp.HAction.Execute("PageSetup", hwp.HParameterSet.HSecDef.HSet)


for i in range(1, 51):
    char_shape = hwp.CharShape
    char_shape.SetItem("Height", 1500)
    hwp.CharShape = char_shape
    hwp.Run("ParagraphShapeAlignCenter")
    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
    hwp.HParameterSet.HInsertText.Text = "팀 빙고"
    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

    # 표 만들기
    hwp.HAction.GetDefault(
        "TableCreate", hwp.HParameterSet.HTableCreation.HSet)  # 표 생성 시작
    hwp.HParameterSet.HTableCreation.Rows = 5  # 행 갯수
    hwp.HParameterSet.HTableCreation.Cols = 5  # 열 갯수
    hwp.HParameterSet.HTableCreation.WidthType = 2
    hwp.HParameterSet.HTableCreation.HeightType = 1
    # 셀 크기 설정
    hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", 5)
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(
        0, hwp.MiliToHwpUnit(25.17))
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(
        1, hwp.MiliToHwpUnit(25.17))
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(
        2, hwp.MiliToHwpUnit(25.17))
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(
        3, hwp.MiliToHwpUnit(25.17))
    hwp.HParameterSet.HTableCreation.ColWidth.SetItem(
        4, hwp.MiliToHwpUnit(25.17))

    hwp.HParameterSet.HTableCreation.CreateItemArray("RowHeight", 5)
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(
        0, hwp.MiliToHwpUnit(24.49))
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(
        1, hwp.MiliToHwpUnit(24.49))
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(
        2, hwp.MiliToHwpUnit(24.49))
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(
        3, hwp.MiliToHwpUnit(24.49))
    hwp.HParameterSet.HTableCreation.RowHeight.SetItem(
        4, hwp.MiliToHwpUnit(24.49))
    hwp.HParameterSet.HTableCreation.TableProperties.TreatAsChar = 1  # 글자처럼 취급
    hwp.HAction.Execute(
        "TableCreate", hwp.HParameterSet.HTableCreation.HSet)  # 위 코드 실행

    alist = []
    visited = [False]*25
    for i in range(25):
        a = randint(1, 50)
        while a in alist:
            a = randint(1, 50)
        alist.append(a)

    while True:
        char_shape = hwp.CharShape
        char_shape.SetItem("Height", 3000)
        hwp.CharShape = char_shape
        hwp.Run("ParagraphShapeAlignCenter")
        while True:
            index = randint(0, 24)
            if (visited[index] == False):
                visited[index] = True
                # 숫자넣기
                text = str(alist[index])
                hwp.HParameterSet.HInsertText.Text = text
                hwp.HAction.Execute(
                    "InsertText", hwp.HParameterSet.HInsertText.HSet)
                break

        # 현재위치가 마지막 셀이라면
        if hwp.KeyIndicator()[-1][1:].split(")")[0] == "E5":
            # hwp.Run("MoveDocEnd")
            # hwp.Run("BreakPara")
            hwp.Run("MovePageUp")
            break
        hwp.Run("TableRightCell")
