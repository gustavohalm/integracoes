from random import randint
import xlwt
from datetime import date


class Funcionario(object):
    def __init__(self, cod, nome, funcao, dep_ir, salario_hr, salario, salario_hora, inss_s_salario_hr, inss_s_salario, irff_s_salario_hr, irff_s_salario, ferias_pagas_mes_ant, ferias_pagas_mes_ant_1_3, abono_pecuniario_mes_ant, abono_pecuniario_mes_ant_1_3, salario_familia, licenca_maternidade, atestado_medico, insalubridade, rescisao, vale, assistencia_medica, inss_ferias_mes_ant, liquido_ferias_mes_ant, pensao_alimenticia, desconto_alimentacao, desconto_energia_agua, dsr_s_horas_extras, hora_extra_fixa, horas_extras_50_hr, horas_extras_50,
                 horas_extras_100_hr, horas_extras_100, dsr_faltas_hr, dsr_faltas, faltas_horas_hr, faltas_horas, dsr_adc_noturno, adc_noturno_hr, adc_noturno, indenizacao_itinere, base_inss_empresa, base_inss_func, base_inss_func_13_sal, base_fgts_13, base_fgts, fgts, base_irrf, deducoes, proventos, descontos, liquido):
        self.cod = cod
        self.nome = nome
        self.funcao = funcao
        self.dep_ir = dep_ir
        self.salario_hr = salario_hr
        self.salario = salario
        self.salario_hora = salario_hora
        self.inss_s_salario_hr = inss_s_salario_hr
        self.inss_s_salario = inss_s_salario
        self.irff_s_salario_hr = irff_s_salario_hr
        self.irff_s_salario = irff_s_salario
        self.ferias_pagas_mes_ant = ferias_pagas_mes_ant
        self.ferias_pagas_mes_ant_1_3 = ferias_pagas_mes_ant_1_3
        self.abono_pecuniario_mes_ant = abono_pecuniario_mes_ant
        self.abono_pecuniario_mes_ant_1_3 = abono_pecuniario_mes_ant_1_3
        self.salario_familia = salario_familia
        self.licenca_maternidade = licenca_maternidade
        self.atestado_medico = atestado_medico
        self.insalubridade = insalubridade
        self.rescisao = rescisao
        self.vale = vale
        self.assistencia_medica = assistencia_medica
        self.inss_ferias_mes_ant = inss_ferias_mes_ant
        self.liquido_ferias_mes_ant = liquido_ferias_mes_ant
        self.pensao_alimenticia = pensao_alimenticia
        self.desconto_alimentacao = desconto_alimentacao
        self.descon_energia_agua = desconto_energia_agua
        self.dsr_s_horas_extras = dsr_s_horas_extras
        self.hora_extra_fixa = hora_extra_fixa
        self.horas_extras_50_hr = horas_extras_50_hr
        self.horas_extras_50 = horas_extras_50
        self.horas_extras_100_hr = horas_extras_100_hr
        self.horas_extras_100 = horas_extras_100
        self.dsr_faltas_hr = dsr_faltas_hr
        self.dsr_faltas = dsr_faltas
        self.faltas_horas_hr = faltas_horas_hr
        self.faltas_horas = faltas_horas
        self.dsr_adc_noturno = dsr_adc_noturno
        self.adc_noturno_hr = adc_noturno_hr
        self.adc_noturno = adc_noturno
        self.indenizacao_itinere = indenizacao_itinere
        self.base_inss_empresa = base_inss_empresa
        self.base_inss_func = base_inss_func
        self.base_inss_func_13_sal = base_inss_func_13_sal
        self.base_fgts_13 = base_fgts_13
        self.base_fgts = base_fgts
        self.fgts = fgts
        self.base_irrf = base_irrf
        self.deducoes = deducoes
        self.proventos = proventos
        self.descontos = descontos
        self.liquido = liquido


# fim da classe
list_func = []
while True:
    apelido = input('Digite o apelido da empresa: ')
    if apelido != "":
        break
while True:
    arquivo = input('Nome do arquivo original(Ex: arquivoxxx): ')
    if arquivo != "":
        break
while True:
    entrada = input('Nome do arquivo de saida(Ex: xxx_admin): ')
    if entrada != "":
        break
while True:
    mes = input('Digite o mes referente a Folha(Ex: 12): ')
    if mes != "":
        break

arquivo = arquivo + ".txt"
folha = open(arquivo)
linhas = folha.readlines()

print(" ------ Progresso ------")
cod = " "
nome = " "
funcao = " "
dep_ir = " "
salario = " "
salario_hr = " "
salario_hora = " "
inss_s_salario = " "
inss_s_salario_hr = " "
irff_s_salario_hr = " "
irff_s_salario = " "
ferias_pagas_mes_ant = " "
ferias_pagas_mes_ant_1_3 = " "
abono_pecuniario_mes_ant = " "
abono_pecuniario_mes_ant_1_3 = " "
salario_familia = " "
licenca_maternidade = " "
atestado_medico = " "
insalubridade = " "
rescisao = " "
vale = " "
assistencia_medica = " "
inss_ferias_mes_anr = " "
liquido_ferias_mes_ant = " "
pensao_alimenticia = " "
desconto_alimenticia = " "
desconto_energia_agua = " "
dsr_s_horas_extras = " "
horas_extras_fixa = " "
horas_extras_50_hr = " "
horas_extras_50 = " "
horas_extras_100_hr = " "
horas_extras_100 = " "
dsr_faltas_hr = " "
dsr_faltas = " "
faltas_horas_hr = " "
faltas_horas = " "
dsr_adc_noturno = " "
adc_noturno_hr = " "
adc_noturno = " "
indenizacao_itinere = ' '
base_inss_empresa = " "
base_inss_func = " "
base_inss_func_13_sal = " "
base_fgts_13 = " "
base_fgts = " "
fgts = " "
base_irrf = " "
deducoes = " "
proventos = " "
descontos = " "
liquido = " "

for i in range(0, len(linhas)):
    linha = linhas[i]
    try:
        cod = linha[linha.index("Cod:") + len("Cod:"): linha.index("Nome")].strip("  ")
        nome = linha[linha.index("Nome:") + len("Nome:"): linha.index("Funcao")].strip("  ")
        funcao = linha[linha.index("Funcao:") + len("Funcao:"): linha.index(" Dep.IR")].strip("  ")
        dep_ir = linha[linha.index("Dep.IR:") + len("Dep.IR:"): linha.index(" |")].strip("  ")
    except:
        print("", end='')

    try:
        salarios = linha[linha.index("1 Salário ") + len("1 Salário "): linha.index(" |")].strip(" ")
        salario =  salarios.split(" ", 1)[1]
        salario_hr = salarios.split(" ", 1)[0]

    except:
        print("", end='')

    try:
        salario_familia =  linha[linha.index("4 Salário Família ") + len("4 Salário Família "): linha.index(" |")].strip(" ").split(" ", 1)[1]

    except:
        print("", end='')
    try:
        licenca_maternidade = linha[linha.index("37 Salário Maternidade ") + len("37 Salário Maternidade "): linha.index(" |")].strip(" ")
    except:
        print("", end='')
    try:
        inss_s_salario = linha[linha.index("INSS Sobre Salário ") + len("INSS Sobre Salário "): len(linha) - 2].strip(" ").split(" ", 1)[1]
        inss_s_salario_hr = linha[linha.index("INSS Sobre Salário ") + len("INSS Sobre Salário "): len(linha) - 2].strip(" ").split(" ", 1)[0]

    except:
        print("", end='')

    try:
        irff_s_salario_hr = linha[linha.index("13 IRRF Sobre Salário ") + len("13 IRRF Sobre Salário "): len(linha) - 2].strip(" ").split(" ", 1)[0]
        irff_s_salario =  linha[linha.index("13 IRRF Sobre Salário ") + len("13 IRRF Sobre Salário "): len(linha) - 2].strip(" ").split(" ", 1)[1]
    except:
        print("", end='')

    try:
        salarios_hora = linha[linha.index("2 Salário Hora ") + len("2 Salário Hora "): linha.index(" |")].strip(" ")
        salario_hora =  salarios_hora.split(" ", 1)[1]
    except:
        print("", end='')

    try:
        ferias_pagas_mes_ant = linha[linha.index("157 Férias Pagas Mês Anterior") + len("157 Férias Pagas Mês Anterior"): linha.index(" |")].strip("  ").split(" ", 1)[1]
    except:
        print("", end='')

    try:
        ferias_pagas_mes_ant_1_3 =  linha[linha.index("158 1/3 Ferias Pagas Mês Anterior") + len("158 1/3 Ferias Pagas Mês Anterior"): linha.index(" |")].strip("  ")
    except:
        print("", end='')

    try:
        abono_pecuniario_mes_ant =  linha[linha.index("161 Abono Pecuniário Mês Anterior ") + len("161 Abono Pecuniário Mês Anterior "): linha.index(" |")].strip("  ").split(" ", 1)[1]
    except:
        print('', end='')

    try:
        abono_pecuniario_mes_ant_1_3 = linha[linha.index("162 1/3 Abono Pecuniário Mês Ant. ") + len("162 1/3 Abono Pecuniário Mês Ant. "): linha.index(" |")].strip("  ").split(" ", 1)[1]
    except:
        print('', end='')

    try:
        dsr_s_horas_extras =  linha[linha.index("D.S.R. Sobre Horas Extras") + len("D.S.R. Sobre Horas Extras"): linha.index(" |")].strip("  ")

    except:
        print("", end='')
    try:
        atestado_medico =  linha[linha.index("1163 Atestado Médico ") + len("1163 Atestado Médico "): linha.index(" |")].strip("  ").split(" ", 1)[1]
    except:
        print("", end='')

    try:
        insalubridade = linha[linha.index("1039 Adc Insalubridade ") + len("1039 Adc Insalubridade "): linha.index(" |")].strip("  ")
    except:
        print("", end='')

    try:
        rescisao =  linha[linha.index("73 Liquido de Rescisão ") + len("73 Liquido de Rescisão "): len(linha) - 2].strip("  ")

    except:
        print("", end='')

    try:
        vale = linha[linha.index("12 Adiantamento Anterior ") + len("12 Adiantamento Anterior "): len(linha) - 2].strip("  ")

    except:
        print("", end='')

    try:
        assistencia_medica = linha[linha.index("1363 Assistencia Medica ") + len("1363 Assistencia Medica "): len(linha) - 2].strip(" ")
    except:
        print("", end='')
    try:
        inss_ferias_mes_anr =  linha[linha.index("45 INSS Sobre Férias ") + len("45 INSS Sobre Férias "): len(linha) - 2].strip(" ")
    except:
        print("", end='')
    try:
        liquido_ferias_mes_ant =  linha[linha.index("53 Liquido de Férias ") + len("53 Liquido de Férias "): len(linha) - 2].strip(" ")
    except:
        print("", end='')
    try:
        pensao_alimenticia =linha[linha.index("1188 Pensão Alimenticia % M ") + len("1188 Pensão Alimenticia % M "): len(linha) - 2].strip(" ").split(" ", 1)[1]
    except:
        print("", end='')

    try:
        desconto_alimenticia =  linha[linha.index("1154 Desconto Alimentação ") + len("1154 Desconto Alimentação "): len(linha) - 2].strip(" ")
    except:
        print("", end='')

    try:
        desconto_energia_agua = linha[linha.index("1023 Desconto Energia/Agua ") + len("1023 Desconto Energia/Agua "): len(linha) - 2].strip(" ")

    except:
        print("", end='')

    try:
        horas_extras_50_hr = linha[linha.index("Horas Extras 50%") + len("Horas Extras 50%"): linha.index(" | ")].strip("  ").split(" ", 1)[0]
        horas_extras_50 = "R$" + linha[linha.index("Horas Extras 50%") + len("Horas Extras 50%"): linha.index(" | ")].strip("  ").split(" ", 1)[1]

    except:
        print("", end='')

    try:
        horas_extras_fixa =  linha[linha.index("1273 Hora Extra Fixa 50% ") + len("1273 Hora Extra Fixa 50% "): linha.index(" | ")].strip("  ").split(" ", 1)[1]
    except:
        print("", end='')

    try:
        horas_extras_100_hr = linha[linha.index("Hora Extras 100%") + len("Horas Extras 100%"): linha.index(" |")].strip("  ").split(" ", 1)[0]
        horas_extras_100 = linha[linha.index("Hora Extras 100%") + len("Horas Extras 100%"): linha.index(" |")].strip("  ").split(" ", 1)[1]
    except:
        print("", end='')

    try:
        faltas_horas_hr = linha[linha.index("39 Faltas (Dias) ") + len("39 Faltas (Dias) "): len(linha) - 2].strip(" ").split(" ", 1)[0]
        faltas_horas =linha[linha.index("39 Faltas (Dias) ") + len("39 Faltas (Dias) "): len(linha) - 2].strip(" ").split(" ", 1)[1]

    except:
        print("", end='')

    try:
        dsr_faltas_hr = linha[linha.index("1055 Faltas (DSR)") + len("1055 Faltas (DSR)"): len(linha) - 2].strip(" ").split(" ", 1)[0]
        dsr_faltas =  linha[linha.index("1055 Faltas (DSR)") + len("1055 Faltas (DSR)"): len(linha) - 2].strip(" ").split(" ", 1)[1]

    except:
        print("", end='')

    try:
        dsr_adc_noturno = linha[linha.index("DSR Adicional Noturno") + len("DSR Adicional Noturno"): linha.index(" |")].strip("  ")

    except:
        print("", end='')

    try:
        adc_noturno_hr = linha[linha.index("Adicional Noturno 25%") + len("Adicional Noturno 25%"): linha.index(" |")].strip("  ").split(" ", 1)[0]
        adc_noturno = linha[linha.index("Adicional Noturno 25%") + len("Adicional Noturno 25%"): linha.index(" |")].strip("  ").split(" ", 1)[1]

    except:
        print("", end='')

    try:
        indenizacao_itinere = linha[linha.index("Indenização Horas Itinere ") + len("Indenização Horas Itinere"): linha.index(" |")].strip("  ").split(" ", 1)[1]
    except:
        print("", end="")

    try:
        base_inss_empresa = linha[linha.index("Base INSS Empresa: ") + len("Base INSS Empresa: "): linha.index(" Base INSS Func")].strip(" ")
        base_inss_func = linha[linha.index("Base INSS Func.: ") + len("Base INSS Func. : "): linha.index(" Base INSS Func 13o Sal:")].strip("  ")
        base_inss_func_13_sal =linha[linha.index("Base INSS Func 13o Sal:") + len("Base INSS Func 13o Sal:"): linha.index(" |")].strip("  ")

    except:
        print("", end='')

    try:
        base_fgts_13 =  linha[linha.index("Base FGTS 13o: ") + len("Base FGTS 13o: "): linha.index(" Base F.G.T.S.:")].strip(" ")
        base_fgts =  linha[linha.index("Base F.G.T.S.: ") + len("Base F.G.T.S.: "): linha.index(" |")].strip("  ")
        base_fgts = base_fgts[0:base_fgts.index("F.G.T.S.:")]
        fgts = linha[linha.index("F.G.T.S.: ") + len("F.G.T.S.: "): linha.index(" |")].strip("  ")

    except:
        print("", end='')

    try:
        fgts = fgts[fgts.index("F.G.T.S.: ") + len("F.G.T.S.: "): -1].strip(" ")

    except:
        print("", end='')
    try:
        base_irrf = linha[linha.index("Base IRRF:") + len("Base IRRF:"): linha.index(" Deducoes:")].strip("  ")
        deducoes = linha[linha.index("Deducoes: ") + len("Deducoes: "): linha.index(" |")].strip("  ")
    except:
        print("", end='')

    try:
        proventos =  linha[linha.index("Proventos:") + len("Proventos:"): linha.index(" Descontos:")].strip("  ")
        descontos =  linha[linha.index("Descontos:") + len("Descontos:"): linha.index(" Liquido:")].strip("  ")
        liquido = linha[linha.index("Liquido: ") + len("Liquido: "): linha.index(" |")].strip("  ")

        # list_func.append( Funcionario(cod, nome, funcao, dep_ir, salario_hr, salario, salario_hora, inss_s_salario_hr, inss_s_salario, dsr_s_horas_extras,horas_extras_50_hr, horas_extras_50, horas_extras_100_hr, horas_extras_100, faltas_horas_hr,faltas_horas, dsr_adc_noturno,  adc_noturno_hr,adc_noturno,indenizacao_itinere, base_inss_empresa, base_inss_func, base_inss_func_13_sal, base_fgts_13, base_fgts, fgts, base_irrf, deducoes, proventos, descontos, liquido ))

        #  Funcionario(self, cod, nome, funcao, dep_ir, salario_hr, salario, salario_hora, inss_s_salario_hr, inss_s_salario,
        #  ferias_pagas_mes_ant, ferias_pagas_mes_ant_1_3, abono_pecuniario_mes_ant, abono_pecuniario_mes_ant_1_3,
        #  salario_familia, licenca_maternidade, atestado_medico, insalubridade, rescisao, vale,
        #  assistencia_medica, inss_ferias_mes_ant, liquido_ferias_mes_ant, pensao_alimenticia,
        #  desconto_alimentacao, desconto_energia_agua, dsr_s_horas_extras, hora_extra_fixa, horas_extras_50_hr,
        #  horas_extras_50, horas_extras_100_hr, horas_extras_100, dsr_faltas_hr, dsr_faltas, faltas_horas_hr,
        #  faltas_horas, dsr_adc_noturno, adc_noturno_hr, adc_noturno, indenizacao_itinere, base_inss_empresa,
        #  base_inss_func, base_inss_func_13_sal, base_fgts_13, base_fgts, fgts, base_irrf, deducoes, proventos,
        #  descontos, liquido)

        list_func.append(
            Funcionario(cod, nome, funcao, dep_ir, salario_hr, salario, salario_hora, inss_s_salario_hr, inss_s_salario, irff_s_salario_hr, irff_s_salario,
                        ferias_pagas_mes_ant, ferias_pagas_mes_ant_1_3, abono_pecuniario_mes_ant, abono_pecuniario_mes_ant_1_3,
                        salario_familia, licenca_maternidade, atestado_medico, insalubridade, rescisao, vale,
                        assistencia_medica, inss_ferias_mes_anr, liquido_ferias_mes_ant, pensao_alimenticia,
                        desconto_alimenticia, desconto_energia_agua, dsr_s_horas_extras, horas_extras_fixa, horas_extras_50_hr,
                        horas_extras_50, horas_extras_100_hr, horas_extras_100, dsr_faltas_hr, dsr_faltas, faltas_horas_hr,
                        faltas_horas, dsr_adc_noturno, adc_noturno_hr, adc_noturno, indenizacao_itinere, base_inss_empresa,
                        base_inss_func, base_inss_func_13_sal, base_fgts_13, base_fgts, fgts, base_irrf, deducoes, proventos,
                        descontos, liquido)
        )
        print('Funcionario adicionado')
        cod = " "
        nome = " "
        funcao = " "
        dep_ir = " "
        salario = " "
        salario_hr = " "
        salario_hora = ' '
        inss_s_salario = " "
        inss_s_salario_hr = " "
        irff_s_salario_hr = " "
        irff_s_salario = " "
        ferias_pagas_mes_ant = " "
        ferias_pagas_mes_ant_1_3 = " "
        abono_pecuniario_mes_ant = " "
        abono_pecuniario_mes_ant_1_3 = " "
        salario_familia = " "
        licenca_maternidade = " "
        atestado_medico = " "
        insalubridade = " "
        rescisao = " "
        vale = " "
        assistencia_medica = " "
        inss_ferias_mes_anr = " "
        liquido_ferias_mes_ant = " "
        pensao_alimenticia = " "
        desconto_alimenticia = " "
        desconto_energia_agua = " "
        dsr_s_horas_extras = " "
        horas_extras_50_hr = " "
        horas_extras_fixa = " "
        horas_extras_50 = " "
        horas_extras_100_hr = " "
        horas_extras_100 = " "
        dsr_faltas_hr = " "
        dsr_faltas = " "
        faltas_horas_hr = " "
        faltas_horas = " "
        dsr_adc_noturno = " "
        adc_noturno_hr = " "
        adc_noturno = " "
        indenizacao_itinere = " "
        base_inss_empresa = " "
        base_inss_func = " "
        base_inss_func_13_sal = " "
        base_fgts_13 = " "
        base_fgts = " "
        fgts = " "
        base_irrf = " "
        deducoes = " "
        proventos = " "
        descontos = " "
        liquido = " "
    except:
        print("", end='')

    print("_", end='')

funcionarios = xlwt.Workbook()
style = xlwt.XFStyle()  # ?????
styleC = xlwt.XFStyle()  # ?????

font = xlwt.Font()  # ???????
font.name = 'Times New Roman'
font.bold = True
font.color_index = 0

borders = xlwt.Borders()
borders.left = 2
borders.right = 2
borders.top = 2
borders.bottom = 5

pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = xlwt.Style.colour_map['gray25']
style.font = font
style.borders = borders
styleC.font = font
styleC.borders = borders
styleC.pattern = pattern
alignment = xlwt.Alignment()
styleC.alignment.wrap = 1

sheet = funcionarios.add_sheet("Folha")
# list_cabecalho = [  'Cod', 'Nome', 'Funcao','INSS S/ Salario(Hrs/Dias)', 'INSS S/ Salario(Valor)', 'D.S.R S/ Horas Extras', 'H. Extras 50%(Horas)', 'H. Extras 50%(Valor)', 'H. Extras 100%(Hrs)', 'H. Extras 100%(Valor)', 'Faltas/Atrasos (Horas)', 'Faltas/Atasos(Valor)' , 'D.S.R ADICIONAL NOTURNO','Adicional Noturno(Hrs.)', 'Adicional Noturno(Valor)' , 'Indenização Hr. Itinere', 'FGTS', 'Salario(Dias)', 'Salario(Valor)' 'Deducoes', 'Proventos', 'Descontos', 'Liquido' ]
list_cabecalho = ['Cod', 'Nome', 'Funcao', 'Salário(Qtde.)', 'Salário(Valor)', 'Horas extras fixas', 'Hr. Extras 50%(Qtde)', 'Hr. Extras 50%(Valor)', 'Hr. Extras 100%(Qtde)', 'Hr. Extras 100%(Valor)', 'Férias pagas mês ant', '1/3 férias pagas mês ant', 'Abono pecuniario mes ant', '1/3 pecuniario mes ant', 'Salario Familia', 'Licença maternidade', 'Atesdado Médico', 'Dsr', 'Adic Not. 25%', 'Hora In itinere', 'Recisão', 'Insalubr', 'TOTAL PROVENTOS', 'INSS Salário(Qtde)', 'INSS Salario(valor)', 'Vale', 'IRRF Salario(Qtdade)', 'IRRF Salario(Valor)',
                  'Faltas(Qtde)', 'Faltas(Valor)', 'DSR(Qtde)', 'DSR(Valor)', 'Assist. Médica', 'INSS Férias mês ant.', 'Liquido Férias Mê amt', 'Pensão Alimenticia', 'Desconto Aliment.', 'Desconto energia/agua', 'Total Descontos', 'Total Liquido']

for i in range(0, len(list_cabecalho)):
    sheet.write(2, i, list_cabecalho[i], style=styleC)
sheet.col(0).width = 276 * 8

sheet.row(2).height_mismatch = True
sheet.row(2).height = 25 * 25
for f in range(0, len(list_func)):
    func = list_func[f]
    # list_funcionario = [ func.cod, func.nome, func.funcao,  func.inss_s_salario_hr,func.inss_s_salario, func.dsr_s_horas_extras, func.horas_extras_50_hr,func.horas_extras_50, func.horas_extras_100_hr, func.horas_extras_100, func.faltas_horas_hr,func.faltas_horas,func.dsr_adc_noturno, func.adc_noturno_hr, func.adc_noturno, func.indenizacao_itinere, func.fgts,func.salario_hr, func.salario, func.deducoes, func.proventos, func.descontos, func.liquido ]

    #  Funcionario(self, cod, nome, funcao, dep_ir, salario_hr, salario, salario_hora, inss_s_salario_hr, inss_s_salario,
    #  ferias_pagas_mes_ant, ferias_pagas_mes_ant_1_3, abono_pecuniario_mes_ant, abono_pecuniario_mes_ant_1_3,
    #  salario_familia, licenca_maternidade, atestado_medico, insalubridade, rescisao, vale,
    #  assistencia_medica, inss_ferias_mes_ant, liquido_ferias_mes_ant, pensao_alimenticia,
    #  desconto_alimentacao, desconto_energia_agua, dsr_s_horas_extras, hora_extra_fixa, horas_extras_50_hr,
    #  horas_extras_50, horas_extras_100_hr, horas_extras_100, dsr_faltas_hr, dsr_faltas, faltas_horas_hr,
    #  faltas_horas, dsr_adc_noturno, adc_noturno_hr, adc_noturno, indenizacao_itinere, base_inss_empresa,
    #  base_inss_func, base_inss_func_13_sal, base_fgts_13, base_fgts, fgts, base_irrf, deducoes, proventos,
    #  descontos, liquido)

    # list_funcionario = [ func.cod, func.nome, func.funcao,  func.inss_s_salario_hr,func.inss_s_salario, func.dsr_s_horas_extras, func.horas_extras_50_hr,func.horas_extras_50, func.horas_extras_100_hr, func.horas_extras_100, func.faltas_horas_hr,func.faltas_horas,func.dsr_adc_noturno, func.adc_noturno_hr, func.adc_noturno, func.indenizacao_itinere, func.fgts,func.salario_hr, func.salario, func.deducoes, func.proventos, func.descontos, func.liquido ]
    list_funcionario = [
        func.cod, func.nome, func.funcao, func.salario_hr, func.salario, func.hora_extra_fixa, func.horas_extras_50_hr, func.horas_extras_50, func.horas_extras_100_hr, func.horas_extras_100, func.ferias_pagas_mes_ant, func.ferias_pagas_mes_ant_1_3, func.abono_pecuniario_mes_ant, func.abono_pecuniario_mes_ant_1_3, func.salario_familia, func.licenca_maternidade, func.atestado_medico, func.dsr_adc_noturno, func.adc_noturno, func.indenizacao_itinere, func.rescisao, func.insalubridade, func.proventos, func.inss_s_salario_hr, func.inss_s_salario,
        func.vale, func.irff_s_salario_hr, func.irff_s_salario, func.faltas_horas_hr, func.faltas_horas, func.dsr_faltas_hr, func.dsr_faltas, func.assistencia_medica, func.inss_ferias_mes_ant, func.liquido_ferias_mes_ant, func.pensao_alimenticia, func.desconto_alimentacao, func.descon_energia_agua, func.descontos, func.liquido
    ]

    for h in range(0, len(list_funcionario)):
        if (h == 22):
            sheet.write(f + 3, h, list_funcionario[h], style=styleC)
        elif (h == 39):
            sheet.write(f + 3, h, list_funcionario[h], style=styleC)
        else:
            sheet.write(f + 3, h, list_funcionario[h], style=style)
        sheet.col(h + 1).width = 276 * 12
    sheet.row(f + 3).height_mismatch = True
    sheet.row(f + 3).height = 25 * 10
sheet.col(1).width = 276 * 32
sheet.col(2).width = 276 * 20
data_atual = date.today()

nomearquivo = entrada + ".xls"
funcionarios.save(nomearquivo)
print("Conversão completa, arquivo ", nomearquivo)
input("Aperte Enter para sair")


