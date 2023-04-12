import xlwings as xw
import time

app = xw.App(visible=True, add_book=False)
workbook = app.books.add()
worksheet = workbook.sheets.active

worksheet.range('A1:Q33').api.Font.Name = '宋体'

worksheet.range('A1').expand('right').api.Font.Bold = True
worksheet.range('A1:Q33').api.Font.Size = 8
worksheet.range('A1:Q33').api.HorizontalAlignment = -4108
worksheet.range('A1:Q33').api.VerticalAlignment = -4108
worksheet.range('A1').value = '建立部门'
worksheet.range('A1:B1').merge()
worksheet.range('C1').value = '洁净科技事业部'
worksheet.range('D1').value = '管理部门'
worksheet.range('E1:G1').merge()
worksheet.range('H1').value = '编辑'
worksheet.range('H1:I1').merge()
worksheet.range('J1:M1').merge()
worksheet.range('N1').value = '审核'
worksheet.range('O1:Q1').merge()

worksheet.range('A2').value = '版本'
worksheet.range('A2:B2').merge()
worksheet.range('C2').value = 1
worksheet.range('D2').value = '版本日期'
worksheet.range('E2').value = '2022/4/17'
worksheet.range('E2:G2').merge()
worksheet.range('H2').value = '变更人'
worksheet.range('H2:I2').merge()
worksheet.range('J2').value = '路飞'
worksheet.range('J2:M2').merge()
worksheet.range('N2').value = '页码'
worksheet.range('Q2').value = '1/1'
worksheet.range('O2:Q2').merge()

worksheet.range('A4').value = '标题'
worksheet.range('A4:B4').merge()
worksheet.range('C4').value = 'GAH高温有隔板过滤器 BOM材料清单'
worksheet.range('C4:K4').merge()
worksheet.range('L4').value = '编码：'
worksheet.range('L4:M4').merge()
worksheet.range('N4').value = 'GAH1F8-A1S4-B1'
worksheet.range('N4:Q4').merge()

worksheet.range('A5').value = '工艺'
worksheet.range('A5:A8').merge()
worksheet.range('B5').value = '产品名称'
worksheet.range('B5:C5').merge()
worksheet.range('B6').value = '高温中效有隔板过滤器'
worksheet.range('B6:C6').merge()
worksheet.range('B7').value = '分隔物'
worksheet.range('B7:C7').merge()
worksheet.range('B8').value = '瓦楞铝箔'
worksheet.range('B8:C8').merge()

worksheet.range('D5').value = '结构'
worksheet.range('D5:E5').merge()
worksheet.range('D6').value = '内翻式'
worksheet.range('D6:E6').merge()
worksheet.range('D7').value = '进风面护网'
worksheet.range('D7:E7').merge()
worksheet.range('D8').value = '镀锌碾平网'
worksheet.range('D8:E8').merge()

worksheet.range('F5').value = '级别'
worksheet.range('F5:G5').merge()
worksheet.range('F6').value = 'F8'
worksheet.range('F6:G6').merge()

worksheet.range('H5').value = '滤材'
worksheet.range('H5:I5').merge()
worksheet.range('H6').value = 'F8玻纤B'
worksheet.range('H6:I6').merge()

worksheet.range('F7').value = '出风面护网'
worksheet.range('F7:I7').merge()
worksheet.range('F8').value = '镀锌碾平网'
worksheet.range('F8:I8').merge()

worksheet.range('J5').value = '外框架'
worksheet.range('J5:K5').merge()
worksheet.range('J6').value = '镀锌外框(292mm)无法兰'
worksheet.range('J6:K6').merge()

worksheet.range('J7').value = '密封条'
worksheet.range('J7:K7').merge()
worksheet.range('J8').value = '出风面玻纤管'
worksheet.range('J8:K8').merge()

worksheet.range('L5').value = '产品规格尺寸'
worksheet.range('L5:Q5').merge()
W_H_D = input('请输入要导入的规格(如Width*Height*Depth)：')
WHD = W_H_D.split('*')
W = WHD[0]
H = WHD[1]
D = WHD[2]
# print(float(W),float(H),float(D))
# print(float(W)-7,float(H)+10,float(D)-5)

worksheet.range('L6').value = W_H_D
worksheet.range('L6:Q6').merge()

worksheet.range('L7').value = '加强筋'
worksheet.range('L7:M7').merge()
worksheet.range('N7').value = '产品图纸'
worksheet.range('N7:Q7').merge()

worksheet.range('L8').value = '无加强筋'
worksheet.range('L8:M8').merge()
# worksheet.range('N7').value = '产品图纸'
worksheet.range('N8:Q8').merge()

worksheet.range('A10').value = '序号'
worksheet.range('B10').value = '材料类型'
worksheet.range('C10').value = '材料名称'
worksheet.range('D10').value = '原材料规格mm'
worksheet.range('E10').value = '半成品规格mm'
worksheet.range('E10:G10').merge()
worksheet.range('H10').value = '用量/只'
worksheet.range('H10:I10').merge()
worksheet.range('J10').value = '单位'
worksheet.range('K10').value = '标准mm'
worksheet.range('L10').value = '材质'
worksheet.range('M10').value = '图示'
worksheet.range('N10').value = '图号'
worksheet.range('O10').value = '供应商'
worksheet.range('P10').value = '物料编号'
worksheet.range('Q10').value = '采购单位'

worksheet.range('A11').value = 'B'
worksheet.range('B11').value = '顶框'
worksheet.range('C11').value = '镀锌外框'
worksheet.range('D11').value = '292厚度0.6mm'
worksheet.range('E11').value = W+'x'+D
worksheet.range('E11:G11').merge()
worksheet.range('H11').value = '2'
worksheet.range('H11:I11').merge()
worksheet.range('J11').value = '块'
worksheet.range('K11').value = '+0~-2'
worksheet.range('L11').value = '镀锌钢板'
worksheet.range('M11').value = 'Picture'
worksheet.range('N11').value = 'TCKB-0003/TCKB-0004/CKB-0001'
worksheet.range('O11').value = '乐之伟'
worksheet.range('P11').value = 'LF0002'
worksheet.range('Q11').value = '套'

worksheet.range('A12').value = 'C'
worksheet.range('B12').value = '侧框'
worksheet.range('C12').value = '镀锌外框'
worksheet.range('D12').value = '292厚度0.6mm'
worksheet.range('E12').value = H+'x'+D
worksheet.range('E12:G12').merge()
worksheet.range('H12').value = '2'
worksheet.range('H12:I12').merge()
worksheet.range('J12').value = '块'
worksheet.range('K12').value = '+0~-2'
worksheet.range('L12').value = '镀锌钢板'
worksheet.range('M12').value = 'Picture'
worksheet.range('N12').value = 'TCKB-0003/TCKB-0004/CKB-0001'
worksheet.range('O12').value = '乐之伟'
worksheet.range('P12').value = 'LF0002'
worksheet.range('Q12').value = '套'

worksheet.range('A13').value = 'D'
worksheet.range('B13').value = '无法兰'

worksheet.range('A14').value = 'E'
worksheet.range('B14').value = '角件'
worksheet.range('C14').value = '三角筋'
worksheet.range('D14').value = '厚度1.0mm'
worksheet.range('E14').value = '60x18'
worksheet.range('E14:G14').merge()
worksheet.range('H14').value = '8'
worksheet.range('H14:I14').merge()
worksheet.range('J14').value = '个'
worksheet.range('K14').value = '+0~-2'
worksheet.range('L14').value = '镀锌钢板'
worksheet.range('M14').value = 'Picture'
worksheet.range('N14').value = 'TCKB-0003/TCKB-0004/CKB-0001'
worksheet.range('O14').value = '乐之伟'
worksheet.range('P14').value = 'LF0002'
worksheet.range('Q14').value = '套'

worksheet.range('A15').value = 'F1'
worksheet.range('B15').value = '加固'
worksheet.range('C15').value = '铆钉'
worksheet.range('D15').value = 'Φ3.2x9'
worksheet.range('E15').value = '60x18'
worksheet.range('E15:G15').merge()
worksheet.range('H15').value = '12'
worksheet.range('H15:I15').merge()
worksheet.range('J15').value = '个'
# worksheet.range('K15').value = '+0~-2'
worksheet.range('L15').value = '铝'
worksheet.range('M15').value = 'Picture'
# worksheet.range('N15').value = 'TCKB-0003/TCKB-0004/CKB-0001'
# worksheet.range('O15').value = '乐之伟'
worksheet.range('P15').value = 'LF0003'
worksheet.range('Q15').value = '个'

worksheet.range('A16').value = 'F2'
worksheet.range('B16').value = '加固'
worksheet.range('C16').value = '铆钉'
worksheet.range('D16').value = 'Φ3.2x7'
worksheet.range('E16').value = '60x18'
worksheet.range('E16:G16').merge()
worksheet.range('H16').value = '24'
worksheet.range('H16:I16').merge()
worksheet.range('J16').value = '个'
# worksheet.range('K16').value = '+0~-2'
worksheet.range('L16').value = '铝'
worksheet.range('M16').value = 'Picture'
# worksheet.range('N15').value = 'TCKB-0003/TCKB-0004/CKB-0001'
# worksheet.range('O15').value = '乐之伟'
worksheet.range('P16').value = 'LF0004'
worksheet.range('Q16').value = '个'

worksheet.range('A17').value = 'G'
worksheet.range('A18').value = 'H'
worksheet.range('A19').value = 'I'
worksheet.range('A20').value = 'J'
worksheet.range('A21').value = 'K'
worksheet.range('B17').value = '滤材'
worksheet.range('B17:B21').merge()
worksheet.range('C17').value = 'F8玻纤B'
worksheet.range('C17:C21').merge()
worksheet.range('D17').value = '604mmx500m'
worksheet.range('D17:D21').merge()
worksheet.range('E17').value = '门幅'
worksheet.range('F17').value = '='
worksheet.range('G17').value = float(W)-7
worksheet.range('E18').value = '折高'
worksheet.range('F18').value = '='
worksheet.range('G18').value = '7'
worksheet.range('E19').value = '折幅'
worksheet.range('F19').value = '='
worksheet.range('G19').value = float(D)-30
worksheet.range('E20').value = '总高'
worksheet.range('F20').value = '='
worksheet.range('G20').value = float(W)-19
worksheet.range('E21').value = '折数'
worksheet.range('F21').value = '='
worksheet.range('G21').value = (float(W)+25-6)/7/2
worksheet.range('H17').value = '23.05'
worksheet.range('H17:H21').merge()
worksheet.range('I17').value = '13.92'
worksheet.range('I17:I21').merge()
worksheet.range('J17').value = 'm/m²'
worksheet.range('J17:J21').merge()
worksheet.range('K17').value = '+0.5~-0.5'
worksheet.range('K18').value = '+0.1~-0.1'
worksheet.range('K19').value = '+1~-1'
worksheet.range('K20').value = '+5~-5'
worksheet.range('K21').value = '+2~-2'
worksheet.range('L17').value = '玻璃纤维'
worksheet.range('L17:L21').merge()
worksheet.range('M17').value = 'Picture'
worksheet.range('M17:M21').merge()
worksheet.range('N17').value = ''
worksheet.range('N17:N21').merge()
worksheet.range('O17').value = 'HV/重庆再升/中材'
worksheet.range('O17:O21').merge()
worksheet.range('P17').value = 'LM0007'
worksheet.range('P17:P21').merge()
worksheet.range('Q17').value = '卷'
worksheet.range('Q17:Q21').merge()

worksheet.range('A22').value = 'L'
worksheet.range('A23').value = 'M'
worksheet.range('A24').value = 'N'
worksheet.range('B22').value = '分隔材料'
worksheet.range('B22:B24').merge()
worksheet.range('C22').value = '瓦楞铝箔'
worksheet.range('C22:C24').merge()
worksheet.range('D22').value = '270'
worksheet.range('D22:D24').merge()
worksheet.range('E22').value = '折高'
worksheet.range('F22').value = '='
worksheet.range('G22').value = '6.75'
worksheet.range('E23').value = '长度'
worksheet.range('F23').value = '='
worksheet.range('G23').value = float(H)-6
worksheet.range('E24').value = '宽度'
worksheet.range('F24').value = '='
worksheet.range('G24').value = float(D)-28
worksheet.range('H22').value = '88'
worksheet.range('H22:I24').merge()
worksheet.range('J22').value = 'm'
worksheet.range('J22:J24').merge()
worksheet.range('K22').value = '+0.3~-0.3'
worksheet.range('K23').value = '+2~-2'
worksheet.range('K24').value = '+1~-1'
worksheet.range('L22').value = '铝箔'
worksheet.range('L22:L24').merge()
worksheet.range('M22').value = 'Picture'
worksheet.range('M22:M24').merge()
worksheet.range('N22').value = ''
worksheet.range('N22:N24').merge()
worksheet.range('O22').value = '东升'
worksheet.range('O22:O24').merge()
worksheet.range('P22').value = 'LQ0001'
worksheet.range('P22:P24').merge()
worksheet.range('Q22').value = '卷'
worksheet.range('Q22:Q24').merge()

worksheet.range('A25').value = 'O'
worksheet.range('B25').value = '护网'
worksheet.range('C25').value = '镀锌碾平网'
worksheet.range('D25').value = 'Ø1.3,网眼8x16'
worksheet.range('E25').value = str(float(W)-7)+'x'+str(float(H)-7)
worksheet.range('E25:G25').merge()
worksheet.range('H25').value = '2'
worksheet.range('H25:I25').merge()
worksheet.range('J25').value = '片'
worksheet.range('K25').value = '+0~-2'
worksheet.range('L25').value = '镀锌钢板'
worksheet.range('M25').value = 'Picture'
worksheet.range('N25').value = 'TCHA-0004'
worksheet.range('O25').value = '富达/久恩'
worksheet.range('P25').value = 'LN0004'
worksheet.range('Q25').value = '片'

worksheet.range('A26').value = 'Q'
worksheet.range('B26').value = '胶水'
worksheet.range('C26').value = '704胶水'
worksheet.range('D26').value = '5kg/桶'
# worksheet.range('E26').value = '603x603'
# worksheet.range('E26:G26').merge()
worksheet.range('H26').value = '0.67'
worksheet.range('H26:I26').merge()
worksheet.range('J26').value = 'Kg'
worksheet.range('K26').value = '+5%~-5%'
worksheet.range('L26').value = '镀单组份硫化硅橡胶'
worksheet.range('M26').value = 'Picture'
# worksheet.range('N26').value = 'TCHA-0004'
worksheet.range('O26').value = '芮意森'
worksheet.range('P26').value = 'LK0009'
worksheet.range('Q26').value = 'Kg'

worksheet.range('A27').value = 'R'
worksheet.range('A28').value = 'S'
worksheet.range('B27').value = '密封条'
worksheet.range('B27:B28').merge()
worksheet.range('C27').value = '玻纤管'
worksheet.range('C27:C28').merge()
worksheet.range('D27').value = 'Φ8x200'
worksheet.range('D27:D28').merge()
worksheet.range('E27').value = ''
worksheet.range('E27:G27').merge()
worksheet.range('E28').value = ''
worksheet.range('E28:G28').merge()
worksheet.range('H27').value = '2.44'
worksheet.range('H27:I27').merge()
worksheet.range('H28').value = ''
worksheet.range('H28:I28').merge()
worksheet.range('J27').value = 'm'
worksheet.range('J27:J28').merge()
worksheet.range('K27').value = '+2~-2'
worksheet.range('K28:K28').merge()
worksheet.range('L27').value = '玻璃纤维'
worksheet.range('L27:L28').merge()
worksheet.range('M27').value = 'Picture'
worksheet.range('M27:M28').merge()
worksheet.range('N27').value = ''
worksheet.range('N27:N28').merge()
worksheet.range('O27').value = '四维铜业'
worksheet.range('O27:O28').merge()
worksheet.range('P27').value = 'LS0004'
worksheet.range('P28').value = ''
worksheet.range('Q27').value = '米'
worksheet.range('Q28').value = ''

worksheet.range('A29').value = 'T'
worksheet.range('A30').value = 'U'
worksheet.range('B29').value = '加强筋'
worksheet.range('B29:B30').merge()
worksheet.range('C29').value = '无加强筋'
worksheet.range('C29:C30').merge()
worksheet.range('D29').value = '-'
worksheet.range('D29:D30').merge()
worksheet.range('E29').value = '-'
worksheet.range('E29:G29').merge()
worksheet.range('E30').value = '-'
worksheet.range('E30:G30').merge()
worksheet.range('H29').value = '-'
worksheet.range('H29:I29').merge()
worksheet.range('H30').value = '-'
worksheet.range('H30:I30').merge()
worksheet.range('J29').value = '-'
worksheet.range('J29:J30').merge()
worksheet.range('K29').value = '-'
worksheet.range('K29:K30').merge()
worksheet.range('L29').value = '-'
worksheet.range('L29:L30').merge()
worksheet.range('M29').value = '-'
worksheet.range('M29:M30').merge()
worksheet.range('N29').value = '-'
worksheet.range('N29:N30').merge()
worksheet.range('O29').value = '-'
worksheet.range('O29:O30').merge()
worksheet.range('P29').value = '-'
worksheet.range('P30').value = '-'
worksheet.range('Q29').value = '-'
worksheet.range('Q30').value = '-'

worksheet.range('A31').value = 'V'
worksheet.range('B31').value = '包装'
worksheet.range('C31').value = '塑料袋'
worksheet.range('D31').value = '厚薄0.03'
worksheet.range('E31').value = str(float(W)+330)+'x'+str(float(H)+370)
worksheet.range('E31:G31').merge()
worksheet.range('H31').value = '1'
worksheet.range('H31:I31').merge()
worksheet.range('J31').value = '只'
worksheet.range('K31').value = '+5~-0'
worksheet.range('L31').value = '聚乙烯'
worksheet.range('M31').value = 'Picture'
worksheet.range('N31').value = '-'
worksheet.range('O31').value = '高博'
worksheet.range('P31').value = 'LB0021'
worksheet.range('Q31').value = '只'

worksheet.range('A32').value = 'W'
worksheet.range('B32').value = '包装'
worksheet.range('C32').value = '纸箱'
worksheet.range('D32').value = '厚薄7'
worksheet.range('E32').value = str(float(W)+10)+'x'+str(float(D)+20)+'x'+str(float(H)+15)
worksheet.range('E32:G32').merge()
worksheet.range('H32').value = '1'
worksheet.range('H32:I32').merge()
worksheet.range('J32').value = '只'
worksheet.range('K32').value = '+5~-0'
worksheet.range('L32').value = '双瓦楞'
worksheet.range('M32').value = 'Picture'
worksheet.range('N32').value = '-'
worksheet.range('O32').value = '高博'
worksheet.range('P32').value = 'LB0024'
worksheet.range('Q32').value = '只'

worksheet.range('A33').value = 'X'
worksheet.range('B33').value = '包装'
worksheet.range('C33').value = 'LG胶带'
worksheet.range('D33').value = '60mmx100m'
worksheet.range('E33').value = '-'
worksheet.range('E33:G33').merge()
worksheet.range('H33').value = '5.51'
worksheet.range('H33:I33').merge()
worksheet.range('J33').value = 'm'
worksheet.range('K33').value = '+5~-0'
worksheet.range('L33').value = '聚丙烯'
worksheet.range('M33').value = 'Picture'
worksheet.range('N33').value = '-'
worksheet.range('O33').value = '友海'
worksheet.range('P33').value = 'LB0016'
worksheet.range('Q33').value = '卷'


datalist = ['A1','D1','H1','N1','A2','D2','H2',
            'N2','A4','C4','L4','N4','A5','B5',
            'D5','F5','H5','J5','L5','L6','B7',
            'D7','F7','J7','L7','N7','A10','B10',
            'C10','D10','E10','H10','J10','K10',
            'L10','M10','N10','O10','P10','Q10']
for i in datalist:
    # print(i)
    worksheet.range(i).font.name = '宋体'
    worksheet.range(i).font.bold = True
    worksheet.range(i).font.size = 8
    worksheet.range(i).api.HorizontalAlignment = -4108
    worksheet.range(i).api.VerticalAlignment = -4108
worksheet.autofit()

for i in worksheet.range('A1:Q33'):
    for j in range(7, 12):
        i.api.Borders(j).LineStyle = 2
        i.api.Borders(j).Weight = 2
for i in range(33):
    worksheet.range(i+1, 1).row_height = 30

worksheet.range('A1:Q1').api.Borders(8).LineStyle = 1
worksheet.range('A1:Q1').api.Borders(8).Weight = 3

worksheet.range('A2:Q2').api.Borders(9).LineStyle = 1
worksheet.range('A2:Q2').api.Borders(9).Weight = 3

worksheet.range('A4:Q4').api.Borders(8).LineStyle = 1
worksheet.range('A4:Q4').api.Borders(8).Weight = 3
worksheet.range('A4:Q4').api.Borders(9).LineStyle = 1
worksheet.range('A4:Q4').api.Borders(9).Weight = 3

worksheet.range('A8:Q8').api.Borders(9).LineStyle = 1
worksheet.range('A8:Q8').api.Borders(9).Weight = 3

worksheet.range('A9:Q9').api.Borders(10).LineStyle = 3
worksheet.range('A9:Q9').api.Borders(10).Weight = 1
worksheet.range('A9:Q9').api.Borders(7).LineStyle = 3
worksheet.range('A9:Q9').api.Borders(7).Weight = 1

worksheet.range('A10:Q10').api.Borders(8).LineStyle = 1
worksheet.range('A10:Q10').api.Borders(8).Weight = 3
worksheet.range('A10:Q10').api.Borders(7).LineStyle = 1
worksheet.range('A10:Q10').api.Borders(7).Weight = 1
worksheet.range('A10:Q10').api.Borders(10).LineStyle = 1
worksheet.range('A10:Q10').api.Borders(10).Weight = 1
worksheet.range('A10:Q10').api.Borders(9).LineStyle = 1
worksheet.range('A10:Q10').api.Borders(9).Weight = 3
worksheet.range('A10:Q10').api.Borders(11).LineStyle = 1
worksheet.range('A10:Q10').api.Borders(11).Weight = 2

worksheet.range('A33:Q33').api.Borders(9).LineStyle = 1
worksheet.range('A33:Q33').api.Borders(9).Weight = 3

worksheet.range('A1:A33').api.Borders(7).LineStyle = 1
worksheet.range('A1:A33').api.Borders(7).Weight = 3

worksheet.range('Q1:Q33').api.Borders(10).LineStyle = 1
worksheet.range('Q1:Q33').api.Borders(10).Weight = 3

worksheet.range('L4:Q4').color = 207, 218, 249
worksheet.range('L6:Q6').color = 207, 218, 249

worksheet.range('P34').value = '制单日期：'
worksheet.range('Q34').value = time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
worksheet.range('P34:Q34').api.Font.Name = '宋体'
worksheet.range('P34:Q34').api.Font.Size = 8
worksheet.range('P34:Q34').api.HorizontalAlignment = -4108
worksheet.range('P34:Q34').api.VerticalAlignment = -4108

# worksheet.pictures.add('F:\\Python草稿练习\\LOGO.png',left=worksheet.range('C4').left,top=worksheet.range('C4').top)
worksheet.pictures.add('F:\\Python草稿练习\\LOGO.png',left=73,top=90)
workbook.save(fr'F:\\Python草稿练习\\样品采购BOM清单（练）.xlsx')
workbook.close()
app.quit()

