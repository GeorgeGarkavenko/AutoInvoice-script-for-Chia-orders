from glob import glob
from sys import argv

from xlrd import open_workbook
from xlwt import easyxf
from xlutils.copy import copy

def main():
    """
    Make invoice from order for Chia
    """
    
    orders = glob('order*.xls')
    invoices = glob('invoice*.xls')
    
    for order_name in orders:
        
        name_parts = order_name.split('_')
        name_parts[0] = 'invoice'
        invoice_name = "_".join(name_parts)
        
        if invoice_name not in invoices:
            
            rb = open_workbook(order_name, formatting_info = True)
            rs = rb.sheet_by_index(0)
            
            wb = copy(rb)
            ws = wb.get_sheet(0)
            
            header_format = easyxf('font: bold True, height ' + str(20 * 14))
            total_format = easyxf('font: bold True, height ' + str(20 * 12))
            
            ws.cols[2].width = ws.cols[2].width - 300
            ws.cols[5].width = int(ws.cols[5].width * 2 / 3)
            
            order_title = rs.cell(0, 0).value
            invoice_number = order_title.split()[1]
            invoice_title = \
                b'\xd0\x9d\xd0\xb0\xd0\xba\xd0\xbb\xd0\xb0\xd0\xb4\xd0\xbd\xd0\xb0\xd1\x8f: '.decode('utf8') \
                + invoice_number
            ws.write(0, 0, invoice_title, header_format)
            
            total = rs.cell(rs.nrows - 2, rs.ncols - 2).value * 0.8
            ws.write(rs.nrows, rs.ncols - 3, \
                     b'\xd0\x92\xd1\x81\xd0\xb5\xd0\xb3\xd0\xbe \xd0\xba \xd0\xbe\xd0\xbf\xd0\xbb\xd0\xb0\xd1\x82\xd0\xb5: '.decode('utf8')
, total_format)
            ws.write(rs.nrows, rs.ncols - 1, total, total_format)
            
            ws.write(rs.nrows - 1, rs.ncols - 1, -20)
            
            wb.save(invoice_name)
            print('Created: ' + invoice_name)
    
if __name__ == "__main__":
    main(*argv[1:])
