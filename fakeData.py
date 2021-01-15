from faker import Faker
import xlwt

fake = Faker(locale='zh_CN')
workbook = xlwt.Workbook(encoding='ascii')
worksheet = workbook.add_sheet('Worksheet')
n=int(1)
if n==0:
    for i in range(1000):
        name = fake.name()
        worksheet.write(i, 0, label=name)
        address = fake.address()
        worksheet.write(i, 1, label=address)
        bankCard = fake.credit_card_number(card_type=None)
        worksheet.write(i, 2, label=bankCard)
    workbook.save("info_0.xls")

else:
    for i in range(1000):
        company = fake.company()
        worksheet.write(i, 0, label=company)
        web = fake.domain_name(levels=1)
        worksheet.write(i, 1, label=web)
        work = fake.job()
        worksheet.write(i, 2, label=work)
    workbook.save("info_1.xls")

