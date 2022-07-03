import json
import xlwt

workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('sheet1',cell_overwrite_ok=True)

worksheet.write(0, 1, label='Bitcoin_marketCap')
worksheet.write(0, 2, label='Ethereum_marketCap')
worksheet.write(0, 3, label='Tether_marketCap')
worksheet.write(0, 4, label='USD_Coin_marketCap')
worksheet.write(0, 5, label='BNB_marketCap')
worksheet.write(0, 6, label='Binance_USD_marketCap')
worksheet.write(0, 7, label='Cardano_marketCap')
worksheet.write(0, 8, label='XRP_marketCap')
worksheet.write(0, 9, label='Solana_marketCap')
worksheet.write(0, 10, label='Dogecoin_marketCap')
worksheet.write(0, 11, label='otherTotalMarketCap')
worksheet.write(0, 12, label='global')
worksheet.write(0, 13, label='Bitcoin_percent')
worksheet.write(0, 14, label='Ethereum_percent')
worksheet.write(0, 15, label='Tether_percent')
worksheet.write(0, 16, label='USD_percent')
worksheet.write(0, 17, label='BNB_percent')
worksheet.write(0, 18, label='Binance_USD_percent')
worksheet.write(0, 19, label='Cardano_percent')
worksheet.write(0, 20, label='XRP_percent')
worksheet.write(0, 21, label='Solana_percent')
worksheet.write(0, 22, label='Dogecoin_percent')
worksheet.write(0, 23, label='otherTotal_percent')

line = 0

for file_num in range(2,20):
    with open('data/'+str(file_num)+'.json','r',encoding='utf8')as fp:
        json_data = json.load(fp)


    nums_1=len(json_data['data']['quotes'])
    nums_2=12

    for num_1 in range(1,nums_1+1):
        worksheet.write(num_1+line, 0, label=json_data['data']['quotes'][num_1 - 1]['timestamp'])
        for num_2 in range(1, nums_2+1):
            print(json_data['data']['quotes'][num_1 - 1]['quote'][num_2-1]['marketCap'])
            worksheet.write(num_1+line, num_2, label=str(json_data['data']['quotes'][num_1 - 1]['quote'][num_2-1]['marketCap']))

    line = line + nums_1

# 保存
workbook.save('sum.xls')
