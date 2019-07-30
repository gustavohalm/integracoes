from xlrd import open_workbook
import xlwt

while True:
    apelido = input('Digite o apelido da empresa: ')
    if apelido != "":
        break
while True:
    arquivo = input('Nome do arquivo original(Ex: arquivoxxx): ')
    if arquivo != "":
        break

while True:
    first = int( input('Primeira linha de lançamento: ') ) - 1
    if first != "":
        break

while True:
    col_conta = int( input('Digite a Coluna a inserir a Conta Contábil: ')) - 1
    if col_conta != "":
        break
while True:
    col_complemento = int( input('Digite a Coluna  do complemento')) - 1
    if col_complemento != "":
        break

arquivo = arquivo + '.xls'

wb = xlwt.Workbook()
s = wb.add_sheet("sheet" )

lancamentos = open_workbook(arquivo)
sheet = lancamentos.sheet_by_index(0)


for row in range(first, sheet.nrows):
    for col in range(0, sheet.ncols):

        if col == col_conta:
            cell = sheet.cell(row, col_conta)

            if(cell.value.upper() == "ARMAZENAGEM"):
                s.write( row,col_conta,"04.01.01.002.00037")
                print("ok")

            elif(cell.value.upper() == "ESCRITORIO") or (cell.value.upper() == "ESCRITÓRIO"):
                s.write( row,col_conta, "04.01.01.002.00010")
                print("ok")

            elif (cell.value.upper() == "VENDA DE FEIJAO") or (cell.value.upper() == "FEIJAO") or (cell.value.upper() == "VENDA DE FEIJÃO") or (cell.value.upper() == "FEIJÃO"):
                s.write(row, col_conta, "03.01.01.002.0002")
                print("ok")

            elif (cell.value.upper() == "VENDA DE SORGO") or (cell.value.upper() == "SORGO"):
                s.write(row, col_conta, "03.01.01.002.0008")
                print("ok")

            elif (cell.value.upper() == "MANUTENÇÃO") or (cell.value.upper() == "MANUTENÇÃO DE MAQUINAS") or (cell.value.upper() == "FERRAMENTAS"):
                s.write( row,col_conta,"04.01.01.002.0021" )
                print("ok")

            elif (cell.value.upper() == "FERTILIZANTES") or (cell.value.upper() == "FERTILIZANTE"):
                s.write( row,col_conta,"04.01.01.002.0017")
                print("ok")

            elif (cell.value.upper() == "IPVA"):
                s.write(row, col_conta, "04.01.01.002.0027")
                print("ok")

            elif (cell.value.upper() == "VENDA DE ALGODAO") or (cell.value.upper() == "ALGODAO") or (cell.value.upper() == "VENDA DE ALGODÃO") or (cell.value.upper() == "ALGODÃO"):
                s.write(row, col_conta, "03.01.01.002.0011")
                print("ok")

            elif (cell.value.upper() == "TRIBUTOS")  or (cell.value.upper() == "DARF"):
                s.write(row, col_conta, "04.01.01.002.0024")
                print("ok")

            elif (cell.value.upper() == "COMBUSTIVEL") or (cell.value.upper() == "COMBUSTÍVEL") or (cell.value.upper() == "OLEO DIESEL") or (cell.value.upper() == "LUBRIFICANTE"):
                s.write(row, col_conta, "04.01.01.002.0020")
                print("ok")

            elif (cell.value.upper() == "CONSORCIO") or (cell.value.upper() == "CONSÓRCIO"):
                s.write(row, col_conta, "04.01.01.002.00032")
                print("ok")

            elif (cell.value.upper() == "ELEKTRO") or (cell.value.upper() == "ELEKTRO") or (cell.value.upper() == "INTERNET") or (cell.value.upper() == "AGUA") or (cell.value.upper() == "VIVO"):
                s.write(row, col_conta, "04.01.01.002.00013")
                print("ok")

            elif (cell.value.upper() == "FRETE"):
                s.write(row, col_conta, "04.01.01.002.00023")
                print("ok")

            elif (cell.value.upper() == "ASSISTENCIA TECNICA") or (cell.value.upper() == "TECNICO"):
                s.write(row, col_conta, "04.01.01.002.00043")
                print("ok")

            elif (cell.value.upper() == "ARRENDAMENTO"):
                s.write(row, col_conta, "04.01.01.002.00048")
                print("ok")

            elif (cell.value.upper() == "CONTÁBIL") or (cell.value.upper() == "HONORÁRIOS CONTÁBEIS") or (cell.value.upper() == "HONORARIOS CONTABEIS")or (cell.value.upper() == "CONTABIL"):
                s.write(row, col_conta, "04.01.01.002.00044")
                print("ok")

            elif (cell.value.upper() == "COLHEITA") or (cell.value.upper() == "PULVERIZAÇÃO") or (cell.value.upper() == "SEMENTE") or (cell.value.upper() == "PLANTIO") or (cell.value.upper() == "TRATAMENTO") or (cell.value.upper() == "SISTEMATIZAÇÃO"):
                s.write(row, col_conta, "04.01.01.002.00015")
                print("ok")

            elif (cell.value.upper() == "DEFENSIVO"):
                s.write(row, col_conta, "04.01.01.002.00018")
                print("ok")

            elif (cell.value.upper() == "JUROS"):
                s.write(row, col_conta, "04.01.01.002.00031")
                print("ok")

            elif (cell.value.upper() == "MONSANTO") or (cell.value.upper() == "VENDA"):
                s.write(row, col_conta, "03.01.01.002.00007")
                print("ok")

            else:
                s.write(row, col_conta, "00.00.00.000.0000")
                print(" ")

        elif col == col_complemento:
            cell = sheet.cell(row, col_complemento)
            str_cell = str(cell.value)
            doc_type = str_cell.split(' ', 3)[0]

            if doc_type == "NF":
                s.write(row, 7, doc_type)
                val_doc = str_cell.split(' ', 3)[1]

            elif doc_type == "CH-":
                s.write(row, 7, doc_type)

                val_doc = str_cell.split(' ', 3)[1]

            elif doc_type == "CODIGO":
                s.write(row, 7, doc_type)
                val_doc = str_cell.split(' ', 3)[1]
            s.write(row, col, str_cell)
            s.write(row, 8, val_doc)

        else:
            cell = sheet.cell(row, col)
            s.write(row, col, cell.value)


wb.save(apelido + ".xls")





