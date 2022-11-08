import xlsxwriter as xl

def string_format(workbook,color,heading=False):
    format = None
    if heading == True:
        format = workbook.add_format( 
            {'bold': True,
            'font_color': 'white',
            'bg_color': color,
            'center_across': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'border': 2}
        )
    else:
        format = workbook.add_format(
            {'bg_color': color,
            'border': 1}
        )
    return format

def number_format(workbook,color):
    format = workbook.add_format(
        {'bg_color': color,
        'border': 1,
        'num_format': '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'}
    )
    return format

def currency_format(workbook,color):
    format = workbook.add_format(
        {'bg_color': color,
        'border': 1,
        'num_format': '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'}
    )
    return format

def heading_format(workbook):
    format = workbook.add_format( 
        {'bold': True,
        'font_color': 'white',
        'bg_color': '#366092',
        'center_across': True,
        'text_wrap': True,
        'valign': 'vcenter',
        'border': 2}
    )
    return format