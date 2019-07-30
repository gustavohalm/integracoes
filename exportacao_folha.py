import xlwt

class Folha(object):
    def __init(self, salario, horas_extras_50, horas_extras_100, dsr_s_horas_extras, dsr_adicional_noturno, adicional_periculosidade, gratificacao_de_funcao, hora_rod_initinere, adc_insaalubridade, gratificacao, adicional_noturno_25,
               quinquenio, hora_extra_com_convencao, inss_s_salario, irrf_s_salario, liquido_ferias_mes_ant, descontos_energia_agua, ferias_pagas_mes_ant, ferias_pagas, proventos, descontos, liquido):
        self.salario = salario
        self.horas_extras_50 = horas_extras_50
        self.horas_extras_100 = horas_extras_100
        self.dsr_s_horas_extras = dsr_s_horas_extras
        self.dsr_adicional_noturno = dsr_adicional_noturno
        self.adicional_periculosidade = adicional_periculosidade
        self.gratificacao_de_funcao = gratificacao_de_funcao
        self.hora_rod_initinere = hora_rod_initinere
        self.adc_insalubridade =adc_insaalubridade
        self.gratificacao = gratificacao
        self.adicional_noturno_25 = adicional_noturno_25
        self.quinquenio = quinquenio
        self.hora_extra_com_convencao = hora_extra_com_convencao
        self.inss_s_salario = inss_s_salario
        self.irrf_s_salario = irrf_s_salario
        self.liquido_ferias_mes_ant = liquido_ferias_mes_ant
        self.descontos_energia_agua = descontos_energia_agua
        self.ferias_pagas_mes_ant = ferias_pagas_mes_ant
        self.ferias_pagas = ferias_pagas
        self.proventos = proventos
        self.descontos = descontos
        self.liquido = liquido


while  True:
        apelido = input('Digite o apelido da empresa: ')
        if apelido != "":
                break
while True:
        arquivo = input('Nome do arquivo original(Ex: arquivoxxx): ')
        if arquivo != "":
                break
while True:
        entrada = input('Nome do arquivo de saida(Ex: xxx_admin): ')
        if entrada !="":
                break
while True:
        mes = input('Digite o mes  a fazer o lançamento (Ex: 12/2019): ')
        if mes != "":
                break

arquivo = arquivo + ".txt"
folha = open(arquivo)
linhas = folha.readlines()
list_planilha = []

salario = "0,0"
horas_extras_50 = "0,0"
horas_extras_100 = "0,0"
dsr_s_horas_extras = "0,0"
dsr_adicional_noturno = "0,0"
adicional_periculosidade = "0,0"
gratificacao_de_funcao ="0,0"
hrora_rod_initinere = "0,0"
adc_insalubridade = "0,0"
gratificacao = "0,0"
adicional_noturno_25 ="0,0"
quinquenio = "0,0"
hora_extra_com_convencao ="0,0"
pensao_alimentcia_2  = "0,0"
pensao_alimentcia_4 = "0,0"
pensao_alimentcia_3 = "0,0"
inss_s_salario = "0,0"
inss_s_salario_rescisao = "0,0"
irrf_s_salario = "0,0"
irrf_s_salario_rescisao = "0,0"
liquido_ferias_mes_ant = "0,0"
descontos_energia_agua = "0,0"
ferias_pagas_mes_ant = "0,0"
farmacia = "0,0"
ferias_pagas = "0,0"
adiantamento_anterior = "0,0"
desconto_alimentacao = "0,0"
pensao_alimentcia = "0,0"
pensao_alimentcia_m = "0,0"
pensao_alimentcia_liq = "0,0"
vale_transporte_v = "0,0"
vale_transporte = "0,0"
proventos = "0,0"
descontos = "0,0"
descontos_moradia = "0,0"
descontos_aluguel = "0,0"
descontos_aluguel_2 = "0,0"
liquido = "0,0"
liquido_ferias = "0,0"
liquido_rescisao ="0,0"
fgts = "0,0"
fgts_13 = "0,0"
contribuicao_confeterativa = "0,0"
contribuicao_confeterativa_l = "0,0"
contribuicao_confeterativa_2 = "0,0"
contribuicao_confeterativa_3 = "0,0"
contribuicao_confeterativa_4 = "0,0"
emprestimo = "0,0"
contribuicao_confeterativa_d = "0,0"
contribuicao_confeterativa_v = "0,0"
contribuicao_assistencial = "0,0"
adiantamento_pagto_ferias = "0,0"
desc_desp_jardinagem = "0,0"
irrf_desc_ferias = "0,0"
fgsts_rescisao = "0,0"
fgsts_13_rescisao = "0,0"
str_folha="0,0"
total_inss = "0,0"
assistencia_medica = "0,0"
liquido_inss = "0,0"

planilha = xlwt.Workbook()
sheet = planilha.add_sheet("sheet1")
for i in range(0, len(linhas)):
    linha = linhas[i]
    str_folha += linhas[i]
    try:
        salario = linha[linha.index("1 Salário ") + len("1 Salário "): linha.index(" |")].strip(" ").split(" ", 1)[1]
    except:
        print('', end='')


    try:
        inss_s_salario = linha[linha.index("11 INSS Sobre Salário ") + len("11 INSS Sobre Salário "): len(linha) - 2].strip(" ").split(" ", 1)[1]
    except:
        print('', end='')

    try:
        irrf_s_salario =  linha[linha.index("13 IRRF Sobre Salário ") + len("13 IRRF Sobre Salário "): len(linha) - 2].strip(" ").split(" ", 1)[1]
    except:
        print('', end='')
    try:
        irrf_s_salario_rescisao = linha[linha.index("70 IRRF Sobre Salário (Rescisão) ") + len("70 IRRF Sobre Salário (Rescisão) "): len(linha) - 2].strip(" ").split(" ", 1)[1]
    except:
        print('', end='')
    try:
        dsr_s_horas_extras = linha[linha.index("D.S.R. Sobre Horas Extras ") + len( "D.S.R. Sobre Horas Extras "): linha.index(" |")].strip("  ")

    except:
        print("", end='')

    try:
        horas_extras_50 =  linha[linha.index("17 Horas Extras 50% ") + len("17 Horas Extras 50% "): linha.index(" | ")].strip("  ").split(" ", 1)[1]
    except:
        print('', end='')

    try:
        horas_extras_100 =  linha[linha.index("82 Hora Extras 100% ") + len("82 Hora Extras 100% "): linha.index(" | ")].strip("  ").split(" ", 1)[1]
    except:
        print('', end='')

    try:
        dsr_adicional_noturno = linha[linha.index("152 DSR Adicional Noturno ") + len("152 DSR Adicional Noturno "): linha.index(" | ")].strip("  ")
    except:
        print('', end='')

    try:
        adicional_periculosidade = linha[linha.index("9 Adicional Periculosidade  ") + len("9 Adicional Periculosidade  "): linha.index(" | ")].strip("  ")

    except:
        print('', end='')
    try:
        gratificacao_de_funcao =  linha[linha.index("9 Adicional Periculosidade  ") + len("9 Adicional Periculosidade  "): linha.index(" | ")].strip("  ")
    except:
        print('', end='')

    try:
        hora_rod_initinere = linha[linha.index("1037 Hora Rodoviaria/In Itinere  ") + len("1037 Hora Rodoviaria/In Itinere  "): linha.index(" | ")].strip("  ").split(" ", 1)[1]

    except:
        print('', end='')


    try:
        adc_insalubridade = linha[linha.index("1039 Adc Insalubridade ") + len("1039 Adc Insalubridade "): linha.index(" | ")].strip("  ")

    except:
        print('', end='')

    try:
        gratificacao = linha[linha.index("1043 Gratificação  ") + len("1043 Gratificação "): linha.index(" | ")].strip("  ")

    except:
        print('', end='')

    try:
        adicional_noturno_25 = linha[linha.index("1047 Adicional Noturno 25% ") + len("1047 Adicional Noturno 25% "): linha.index(" | ")].strip("  ")
    except:
        print('', end='')

    try:
        quinquenio =  linha[linha.index("1063 Quinquenio ") + len("1063 Quinquenio "): linha.index(" | ")].strip("  ")

    except:
        print('', end='')

    try:
        hora_extra_com_convencao = linha[linha.index("1426 Hora Extra Dom Convenção ") + len("1426 Hora Extra Dom Convenção "): linha.index(" | ")].strip("  ").split(' ', 1)[1]
    except:
        print('', end='')

    try:
        liquido_ferias_mes_ant = linha[linha.index("167 Liquido Férias Mês Anterior ") + len("167 Liquido Férias Mês Anterior "): len(linha) - 2].strip("  ")
    except:
        print('', end='')
    try:
        emprestimo = linha[linha.index("1335 Empréstimo ") + len("1335 Empréstimo "): len(linha) - 2].strip("  ")
    except:
        print('', end='')

    try:
        assistencia_medica = linha[ linha.index('1363 Assistencia Medica') + len('1363 Assistencia Medica') : len(linha) - 2 ].strip("  ")
    except:
        print('', end='')
    try:
        descontos_energia_agua = linha[linha.index("1023 Desconto Energia/Agua ") + len("1023 Desconto Energia/Agua "): len(linha) - 2].strip("  ")
    except:
        print('', end='')
    try:
        descontos_moradia = linha[ linha.index("1196 Desconto Moradia ") + len("1196 Desconto Moradia ") : len(linha) - 2 ].strip(" ")
    except:
        print('', end='')
    try:
        descontos_aluguel = linha[linha.index("1044 Desconto Aluguel ") + len("1044 Desconto Aluguel "): len(linha) - 2].strip(" ")
    except:
        print('', end='')
    try:
        descontos_aluguel_2 = linha[linha.index("1036 Desconto Aluguel ") + len("1036 Desconto Aluguel "): len(linha) - 2].strip(" ")
    except:
        print('', end='')
    try:
        farmacia =  linha[linha.index("142 Farmácia  ") + len("142 Farmácia  "): len(linha) - 2].strip(" ")
    except:
        print('', end='')
    try:
        irrf_desc_ferias = linha[linha.index("253 IRRF Descontado nas Férias") + len("253 IRRF Descontado nas Férias"): len(linha) - 2].strip(" ")
    except:
        print('', end='')
    try:
        desconto_alimentacao = linha[linha.index("1154 Desconto Alimentação  ") + len("1154 Desconto Alimentação  "): len(linha) - 2].strip("  ")
    except:
        print('', end='')
    try:
        pensao_alimentcia_m = linha[linha.index("1188 Pensão Alimenticia % M ") + len("1188 Pensão Alimenticia % M"): len(linha) - 2 ].strip("  ").split(' ', 1)[1]
    except:
        print('',end='')

    try:
        pensao_alimentcia_liq = linha[linha.index("1277 Pensao Alimenticia %Liq") + len("1277 Pensao Alimenticia %Liq"): len(linha) - 2].strip("  ").split(' ', 1)[1]
    except:
        print('', end='')
    try:
        pensao_alimentcia = linha[linha.index("1250 Pensão Alimenticia ") + len("1250 Pensão Alimenticia "): len(linha) - 2].strip("  ").split(' ', 1)[1]
    except:
        print('', end='')
    try:
        pensao_alimentcia_2 = linha[linha.index("1305 Pensao Alimenticia ") + len("1305 Pensao Alimenticia "): len(linha) - 2].strip("  ").split(' ', 1)[1]
    except:
        print('', end='')
    try:
        contribuicao_confeterativa_v = linha[linha.index("1027 Cont Confederativa (v)") + len("1027 Cont Confederativa (v)"): len(linha) - 2].strip("  ").split(' ', 1)[1]
    except:
        print('', end='')
    try:
        vale_transporte_v = linha[linha.index("1038 Vale Transporte (v)") + len("1038 Vale Transporte (v)"): len(linha) - 2].strip("  ")
    except:
        print('', end='')
    try:
        desc_desp_jardinagem = linha[linha.index("1439 Desc Despesas Jardinagem ") + len("1439 Desc Despesas Jardinagem "): len(linha) - 2].strip("  ")
    except:
        print('', end='')
    try:
        contribuicao_assistencial = linha[linha.index("1407 Contribuição Assistencial ") + len("1407 Contribuição Assistencial "): len(linha) - 2].strip("  ").split(' ', 1)[1]
    except:
        print('', end='')

    try:
        contribuicao_confeterativa = linha[linha.index("1127 Contribuicao Confederativa ") + len("1127 Contribuicao Confederativa "): len(linha) - 2].strip("  ").split(' ', 1)[1]
    except:
        print('', end='')

    try:
        contribuicao_confeterativa_2 = linha[linha.index("1349 Contribuição Confederativa ") + len("1349 Contribuição Confederativa "): len(linha) - 2].strip("  ").split(' ', 1)[1]
    except:
        print('', end='')


    try:
        contribuicao_confeterativa_3 = linha[linha.index("1007 Contribuição Confederativa" ) + len("1007 Contribuição Confederativa "): len(linha) - 2].strip("  ").split(' ', 1)[1]
    except:
        print('', end='')

    try:
        contribuicao_confeterativa_4 = linha[linha.index("32 Contribuição Confederativa " ) + len("32 Contribuição Confederativa "): len(linha) - 2].strip("  ").split(' ', 1)[1]
    except:
        print('', end='')


    try:
        contribuicao_confeterativa_d = linha[linha.index("1339 Cont Confederativa Colhedores ") + len("1339 Cont Confederativa Colhedores  "): len(linha) - 2].strip("  ")
    except:
        print('', end='')
    try:
        contribuicao_confeterativa_l = linha[linha.index("1088 Contribuição Confederativa (L) ") + len("1088 Contribuição Confederativa (L) "): len(linha) - 2].strip("  ").split(' ', 1)[1]
    except:
        print('', end='')

    try:
        adiantamento_anterior = linha[linha.index("12 Adiantamento Anterior ") + len("12 Adiantamento Anterior "): len(linha) - 2].strip("  ")
    except:
        print('', end='')
    try:
        liquido_ferias = linha[linha.index("53 Liquido de Férias ") + len("53 Liquido de Férias "): len(linha) - 2].strip("  ")
    except:
        print('', end='')
    try:
        liquido_rescisao = linha[linha.index("73 Liquido de Rescisão ") + len("73 Liquido de Rescisão "): len(linha) - 2].strip("  ")
    except:
        print('', end='')

    try:
        adiantamento_pagto_ferias = linha[linha.index("1381 Adiantamento Pagto Férias  ") + len("1381 Adiantamento Pagto Férias  "): len(linha) - 2].strip("  ")
    except:
        print('', end='')

    try:

        proventos = linha[linha.index("GProventos:H ") + len("GProventos:H "): linha.index("GDescontos")].strip("  ")
        descontos = linha[linha.index("GDescontos:H ") + len("GDescontos:H "): linha.index("GLiquido")].strip("  ")
        liquido   =linha[linha.index("GLiquido:H ") + len("GLiquido:H "): linha.index(" |")].strip("  ")

        liquido = liquido.replace('.', '')
        liquido = liquido.replace(',', '.')
        adiantamento_anterior = adiantamento_anterior.replace('.', '')
        adiantamento_anterior = adiantamento_anterior.replace(',', '.')
        adiantamento_pagto_ferias = adiantamento_pagto_ferias.replace('.', '')
        adiantamento_pagto_ferias = adiantamento_pagto_ferias.replace(',', '.')
        irrf_s_salario = irrf_s_salario.replace('.', '')
        irrf_s_salario = irrf_s_salario.replace(',', '.')
        irrf_s_salario_rescisao = irrf_s_salario_rescisao.replace('.', '')
        irrf_s_salario_rescisao = irrf_s_salario_rescisao.replace(',', '.')
        irrf_desc_ferias = irrf_desc_ferias.replace('.', '')
        irrf_desc_ferias = irrf_desc_ferias.replace(',', '.')
        liquido_ferias = liquido_ferias.replace('.', '')
        liquido_ferias = liquido_ferias.replace(',', '.')
        liquido_ferias_mes_ant = liquido_ferias_mes_ant.replace('.', '')
        liquido_ferias_mes_ant = liquido_ferias_mes_ant.replace(',', '.')
        liquido_rescisao = liquido_rescisao.replace('.', '')
        liquido_rescisao = liquido_rescisao.replace(',', '.')
        descontos_energia_agua = descontos_energia_agua.replace('.', '')
        descontos_energia_agua = descontos_energia_agua.replace(',', '.')
        desconto_alimentacao = desconto_alimentacao.replace('.', '')
        desconto_alimentacao = desconto_alimentacao.replace(',', '.')
        pensao_alimentcia = pensao_alimentcia.replace('.', '')
        pensao_alimentcia = pensao_alimentcia.replace(',', '.')
        pensao_alimentcia_2 = pensao_alimentcia_2.replace('.', '')
        pensao_alimentcia_2 = pensao_alimentcia_2.replace(',', '.')
        pensao_alimentcia_m = pensao_alimentcia_m.replace('.', '')
        pensao_alimentcia_m = pensao_alimentcia_m.replace(',', '.')
        pensao_alimentcia_liq = pensao_alimentcia_liq.replace('.', '')
        pensao_alimentcia_liq = pensao_alimentcia_liq.replace(',', '.')
        emprestimo = emprestimo.replace('.', '')
        emprestimo = emprestimo.replace(',', '.')
        descontos_moradia = descontos_moradia.replace('.', '')
        descontos_moradia = descontos_moradia.replace(',' ,'.')
        assistencia_medica = assistencia_medica.replace('.', '')
        assistencia_medica = assistencia_medica.replace(',', '.')
        contribuicao_confeterativa = contribuicao_confeterativa.replace('.', '')
        contribuicao_confeterativa = contribuicao_confeterativa.replace(',', '.')
        contribuicao_confeterativa_2 = contribuicao_confeterativa_2.replace('.', '')
        contribuicao_confeterativa_2 = contribuicao_confeterativa_2.replace(',', '.')
        contribuicao_confeterativa_3 = contribuicao_confeterativa_3.replace('.', '')
        contribuicao_confeterativa_3 = contribuicao_confeterativa_3.replace(',', '.')
        contribuicao_confeterativa_4 = contribuicao_confeterativa_4.replace('.', '')
        contribuicao_confeterativa_4 = contribuicao_confeterativa_4.replace(',', '.')
        contribuicao_confeterativa_d = contribuicao_confeterativa_d.replace('.', '')
        contribuicao_confeterativa_d = contribuicao_confeterativa_d.replace(',', '.')

        contribuicao_confeterativa_v = contribuicao_confeterativa_v.replace('.', '')
        contribuicao_confeterativa_v = contribuicao_confeterativa_v.replace(',', '.')
        contribuicao_confeterativa_l = contribuicao_confeterativa_l.replace('.', '')
        contribuicao_confeterativa_l = contribuicao_confeterativa_l.replace(',', '.')

        contribuicao_assistencial = contribuicao_assistencial.replace('.', '')
        contribuicao_assistencial = contribuicao_assistencial.replace(',', '.')
        desc_desp_jardinagem  = desc_desp_jardinagem.replace('.', '')
        desc_desp_jardinagem  = desc_desp_jardinagem.replace(',', '.')
        vale_transporte = vale_transporte.replace('.', '')
        vale_transporte = vale_transporte.replace(',', '.')
        vale_transporte_v = vale_transporte_v.replace('.', '')
        vale_transporte_v = vale_transporte_v.replace(',', '.')
        farmacia = farmacia.replace('.', '')
        farmacia = farmacia.replace(',', '.')
        descontos_aluguel = descontos_aluguel.replace('.', '')
        descontos_aluguel = descontos_aluguel.replace(',', '.')
        descontos_aluguel_2 = descontos_aluguel_2.replace('.', '')
        descontos_aluguel_2 = descontos_aluguel_2.replace(',', '.')


        liquido_value = float( liquido )
        adiantamento_anterior_value = float( adiantamento_anterior )
        adiantamento_pagto_ferias_value = float( adiantamento_pagto_ferias )
        irrf_s_salario_value = float( irrf_s_salario )
        irrf_s_salario_rescisao_value = float( irrf_s_salario_rescisao )
        irrf_desc_ferias_value = float( irrf_desc_ferias )
        liquido_ferias_value = float( liquido_ferias )
        liquido_ferias_mes_ant_value = float( liquido_ferias_mes_ant )
        liquido_rescisao_value = float( liquido_rescisao )
        descontos_energia_agua_value = float( descontos_energia_agua )
        desconto_alimentacao_value = float( desconto_alimentacao )
        pensao_alimentcia_value = float( pensao_alimentcia)
        pensao_alimentcia_2_value = float( pensao_alimentcia_2)
        pensao_alimentcia_m_value = float( pensao_alimentcia_m )
        pensao_alimentcia_liq_value = float( pensao_alimentcia_liq )
        descontos_moradia_value = float( descontos_moradia )
        assistencia_medica_value = float( assistencia_medica )
        contribuicao_confeterativa_value = float( contribuicao_confeterativa )
        contribuicao_confeterativa_2_value = float( contribuicao_confeterativa_2 )
        contribuicao_confeterativa_3_value = float( contribuicao_confeterativa_3 )
        contribuicao_confeterativa_4_value = float( contribuicao_confeterativa_4 )
        contribuicao_confeterativa_d_value = float( contribuicao_confeterativa_d )
        contribuicao_confeterativa_v_value = float( contribuicao_confeterativa_v )
        contribuicao_confeterativa_l_value = float( contribuicao_confeterativa_l )
        contribuicao_assistencial_value = float( contribuicao_assistencial)
        vale_transporte_value = float( vale_transporte )
        vale_transporte__v_value = float( vale_transporte_v )
        farmacia_value = float( farmacia )
        descontos_aluguel_value = float( descontos_aluguel)
        descontos_aluguel_2_value = float( descontos_aluguel_2)
        desc_desp_jardinagem_value = float( desc_desp_jardinagem )
        emprestimo_value = float( emprestimo )
        print(' 1')
        liquido_value += ( emprestimo_value + adiantamento_anterior_value + pensao_alimentcia_liq_value + contribuicao_confeterativa_3_value +contribuicao_confeterativa_4_value +irrf_s_salario_value + liquido_ferias_mes_ant_value + liquido_ferias_value +liquido_rescisao_value +descontos_energia_agua_value + desconto_alimentacao_value + pensao_alimentcia_value + contribuicao_assistencial_value + irrf_desc_ferias_value + adiantamento_pagto_ferias_value + contribuicao_confeterativa_2_value)
        liquido_value += (pensao_alimentcia_m_value + irrf_s_salario_rescisao_value + contribuicao_confeterativa_d_value +descontos_moradia_value + assistencia_medica_value + contribuicao_confeterativa_value + contribuicao_confeterativa_v_value + vale_transporte_value + vale_transporte__v_value + farmacia_value + descontos_aluguel_value + descontos_aluguel_2_value + pensao_alimentcia_2_value + desc_desp_jardinagem_value + contribuicao_confeterativa_l_value)
        print('2')
        #iniciar objheto folha
# def __init(self, salario, horas_extras_50, horas_extras_100, dsr_s_horas_extras, dsr_adicional_noturno, adicional_periculosidade, gratificacao_de_funcao, hora_rod_initinere, adc_insaalubridade, gratificacao, adicional_noturno_25,
 #  quinquenio, hora_extra_com_convencao, inss_s_salario, irrf_s_salario, liquido_ferias_mes_ant, descontos_energia_agua, ferias_pagas_mes_ant, ferias_pagas, proventos, descontos, liquido):

        #folha = Folha(salario, horas_extras_50, horas_extras_100, dsr_s_horas_extras, dsr_adicional_noturno, adicional_periculosidade, gratificacao_de_funcao, hora_rod_initinere, adc_insalubridade,gratificacao, adicional_noturno_25,
         #             quinquenio, hora_extra_com_convencao, inss_s_salario, irrf_s_salario, liquido_ferias_mes_ant, descontos_energia_agua, ferias_pagas_mes_ant, ferias_pagas, proventos, descontos, liquido)

        conta_salario = '04.01.01.002.00001' #REDUZIDA
        conta_credito_salario = '01.01.01.001.00001' #REDUZIDA

        data = '05/'+mes
        sheet.write(0,0, '1')
        sheet.write(0, 2, conta_salario),
        sheet.write(0,1,data)
        sheet.write(0,3, conta_credito_salario)
        sheet.write(0,4, 'PAGTO. DE SALARIO - FOLHA')
        sheet.write(0,5,liquido_value)

        print('ok')
        salario = ""
        horas_extras_50 = ""
        horas_extras_100 = ""
        dsr_s_horas_extras = ""
        dsr_adicional_noturno = ""
        adicional_periculosidade = ""
        gratificacao_de_funcao = ""
        hrora_rod_initinere = ""
        contribuicao_confeterativa = ""
        contribuicao_confeterativa_v = ""
        adc_insalubridade = ""
        gratificacao = ""
        adicional_noturno_25 = ""
        quinquenio = ""
        hora_extra_com_convencao = ""
        inss_s_salario = ""
        farmacia = ""
        inss_s_salario_rescisao = ""
        irrf_s_salario = ""
        vale_transporte_v = ""
        irrf_s_salario_rescisao = ""
        liquido_ferias_mes_ant = ""
        descontos_energia_agua = ""
        irrf_desc_ferias = ""
        ferias_pagas_mes_ant = ""
        pensao_alimentcia = ""
        pensao_alimentcia_m = ""
        ferias_pagas = ""
        proventos = ""
        descontos = ""
        descontos_aluguel = ""
        descontos_aluguel_2 = ""
        vale_transporte = ""
        assistencia_medica = ""
        liquido = ""
        desconto_alimentacao = ""
        liquido_ferias = ""
        liquido_rescisao = ""

    except:
        print('', end='')
    try:

        total_inss = linha[linha.index("Cod. 1066  Total Líquido ") + len("Cod. 1066  Total Líquido "): linha.index(" |")].strip(" ")
        conta_inss = '04.01.01.002.00007'
        data_inss = '20/' + mes
        conta_credito_inss =  '01.01.01.001.00001'
        print('not')
        total_inss = total_inss.replace('.', '')
        total_inss = total_inss.replace(',', '.')
        sheet.write(1, 0, '1')
        sheet.write(1, 2, conta_inss),
        sheet.write(1, 1, data_inss)
        sheet.write(1, 3, conta_credito_inss)
        sheet.write(1, 4, 'PAGTO. DE INSS S/ SALARIO - FOLHA')
        sheet.write(1, 5, float( total_inss ) )
        total_inss = ""
    except:

        print('', end='')
fgts_mensal = str_folha[str_folha.index('FGTS Mensal (Recolhimento SEFIP)') + len('FGTS Mensal (Recolhimento SEFIP') : str_folha.index('F G T S Rescisorio (Recolhimento GRRF)')]

fgts = fgts_mensal[fgts_mensal.index(' F.G.T.S.: ') + 52 : fgts_mensal.index(' C.Social:')].strip(" ")
fgts_13 = fgts_mensal[fgts_mensal.index('F.G.T.S. 13o Salario: ') + len('F.G.T.S. 13o Salario: ') : fgts_mensal.index('F.G.T.S. 13o Salario: ') + 42].strip(" ")
data_fgts = '07/' + mes

fgts_rescisorio = str_folha[str_folha.index('F G T S Rescisorio (Recolhimento GRRF)') + len('F G T S Rescisorio (Recolhimento GRRF)') : str_folha.index('C.Social Multa 10%:')]

fgts_rescisao = fgts_rescisorio[ fgts_rescisorio.index('F.G.T.S.: ') + len('F.G.T.S.: ') : fgts_rescisorio.index('C.Social: ') ].strip(" ")
fgts_13_rescisao = fgts_rescisorio[ fgts_rescisorio.index('F.G.T.S. 13o Salario: ') + len('F.G.T.S. 13o Salario: ') : fgts_rescisorio.index('F.G.T.S. 13o Salario: ') + 42 ].strip(" ")

fgts = fgts.replace('.', '')
fgts = fgts.replace(',', '.')
fgts_13 = fgts_13.replace('.', '')
fgts_13 = fgts_13.replace(',', '.')

fgts_rescisao = fgts_rescisao.replace('.', '')
fgts_rescisao = fgts_rescisao.replace(',', '.')
fgts_13_rescisao = fgts_13_rescisao.replace('.', '')
fgts_13_rescisao = fgts_13_rescisao.replace(',', '.')

fgts_value = float(fgts)
fgts_13_value = float(fgts_13)
fgts_rescisao_value = float(fgts_rescisao)
fgts_13_rescisao_value = float(fgts_13_rescisao)

total_fgts = fgts_value + fgts_13_value + fgts_rescisao_value + fgts_13_rescisao_value

conta_fgts = '04.01.01.002.00007'
conta_credito_fgts =  '01.01.01.001.00001'


sheet.write(2, 0, '1')
sheet.write(2, 2, conta_fgts),
sheet.write(2, 1, data_fgts)
sheet.write(2, 3, conta_credito_fgts)
sheet.write(2, 4, 'PAGTO. DE FGTS S/ SALARIO - FOLHA')
sheet.write(2, 5, total_fgts)


nomearquivo = entrada + ".xls"
planilha.save(nomearquivo)