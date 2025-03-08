from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

class LineFormatter:
    def __init__(self, spacing_enabled=False):
        self.spacing_enabled = spacing_enabled
        self.symbol_triggers = {':': True}

    def add_symbol_triggered_border(self, paragraph, symbol):
        if self.spacing_enabled and self.symbol_triggers.get(symbol):
            p = paragraph._element.get_or_add_pPr()
            # 添加右边框并设置24pt间距
            border_right = OxmlElement('w:right')
            border_right.set(qn('w:val'), 'single')
            border_right.set(qn('w:sz'), '8')
            border_right.set(qn('w:space'), '24')
            border_right.set(qn('w:color'), 'auto')
            
            # 添加底部边框到底部
            border_bottom = OxmlElement('w:bottom')
            border_bottom.set(qn('w:val'), 'single')
            border_bottom.set(qn('w:sz'), '8')
            border_bottom.set(qn('w:space'), '0')
            border_bottom.set(qn('w:color'), 'auto')
            
            p.append(border_right)
            p.append(border_bottom)

    def set_spacing_border(self, paragraph):
        if not self.spacing_enabled:
            return
        
        pPr = paragraph._p.get_or_add_pPr()
        # 设置左右缩进为0
        ind_elem = OxmlElement('w:ind')
        ind_elem.set(qn('w:left'), '0')
        ind_elem.set(qn('w:right'), '0')
        pPr.append(ind_elem)
        
        # 添加四周边框实现全宽效果
        for pos in ['top', 'left', 'bottom', 'right']:
            border = OxmlElement(f'w:{pos}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), 'auto')
            pPr.append(border)

    def set_spacing_property(self, paragraph):
        if not self.spacing_enabled:
            return
        
        # 添加段落底部边框
        border = OxmlElement('w:bottom')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'auto')
        p_pr = paragraph._p.get_or_add_pPr()
        p_pr.append(border)
        # 添加缩进控制实现全宽
        ind = OxmlElement('w:ind')
        ind.set(qn('w:left'), '0')
        ind.set(qn('w:right'), '0')
        ind.set(qn('w:firstLine'), '0')  # 清除首行缩进
        p_pr.append(ind)

        p_pr = paragraph._p.get_or_add_pPr()
        spacing = OxmlElement('w:spacing')
        spacing.set(qn('w:before'), '60')  # 6磅
        spacing.set(qn('w:after'), '60')   # 6磅
        p_pr.append(spacing)