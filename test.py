from competitor import GoogleMaps
from openpyxl import load_workbook

big_c_mobile_link = r"D:\projects\competitor_maps_review\excel\Competitor\Competitor\Big C Mobiles Stores Link.xlsx"
big_c_mobile_col = 3
big_c_mobile_title = "Big C Mobile "


Mobile_xpress_link = r"D:\projects\competitor_maps_review\excel\Competitor\Competitor\Mobile Xpress Stores Link.xlsx"
Mobile_xpress_col = 4
Mobile_xpress_title = "Mobile Xpress "

gm = GoogleMaps()

gm.google_report(link=big_c_mobile_link,col=big_c_mobile_col,title=big_c_mobile_title)