from application import application_methods
import os

# clean the stock list get from the
stocks_data = application_methods.load_input_data()

stocks_data.pop(0)
stocks_data.pop(-1)

result = ''
for i in stocks_data:
    result += i.split('|')[0] + '\n'


file_path = r'./stock_list.txt'

with open(file_path, 'w') as f:
    f.write(result)

print(f"文件已保存到：{file_path}")