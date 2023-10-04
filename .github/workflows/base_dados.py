import pprint
from openpyxl import Workbook
import requests

unificada = Workbook()
unificada.create_sheet("Central", 0)
aba = unificada["Central"]
del unificada["Sheet"]

global unidade_alvo, cliente_alvo

unidade_alvo = input("Digite o nome da unidade: ")
cliente_alvo = input("Digite o nome do cliente: ")

# Cria cabeçalhos para o arquivo Excel de upload, conforme padrão do DAP Diel
def cabecalhosExcel():
    # Inserindo cabeçalhos
    aba["A1"] = "Tipo de Solução"
    aba["A2"] = "Unidade"
    aba["B1"] = "Unidade"
    aba["C1"] = "ID da Unidade"
    aba["D1"] = "País"
    aba["D2"] = "Brasil"
    aba["E1"] = "Estado"
    aba["F1"] = "Cidade"
    aba["G1"] = "Fuso Horário (IANA)"
    aba["H1"] = "Status da Unidade"
    aba["H2"] = "Em instalação"
    aba["I1"] = "Latitude e Longitude"
    aba["J1"] = "Endereço"
    aba["K1"] = "Documento 1 Unidade"
    aba["L1"] = "Documento 2 Unidade"
    aba["M1"] = "Documento 3 Unidade"
    aba["N1"] = "Documento 4 Unidade"
    aba["O1"] = "Documento 5 Unidade"
    aba["P1"] = "ID da Máquina"
    aba["Q1"] = "Nome da Máquina"
    aba["R1"] = "Máquina Instalada em"
    aba["S1"] = "Aplicação"
    aba["T1"] = "Tipo de Equipamento"
    aba["U1"] = "Fabricante"
    aba["V1"] = "Fluido"
    aba["W1"] = "Foto 1 Máquina"
    aba["X1"] = "Foto 2 Máquina"
    aba["Y1"] = "Foto 3 Máquina"
    aba["Z1"] = "Foto 4 Máquina"
    aba["AA1"] = "Foto 5 Máquina"
    aba["AB1"] = "Dispositivo de Automação (ID DUT/DAM/DRI)"
    aba["AC1"] = "Posicionamento (INS/AMB/DUO)"
    aba["AD1"] = "DUT DUO (POSIÇÃO SENSORES)"
    aba["AE1"] = "Local de Instalação do DAM"
    aba["AF1"] = "Posicionamento do DAM"
    aba["AG1"] = "Sensor T0 do DAM"
    aba["AH1"] = "Sensor T1 do DAM"
    aba["AI1"] = "Foto 1 Dispositivo de Automação"
    aba["AJ1"] = "Foto 2 Dispositivo de Automação"
    aba["AK1"] = "Foto 3 Dispositivo de Automação"
    aba["AL1"] = "Foto 4 Dispositivo de Automação"
    aba["AM1"] = "Foto 5 Dispositivo de Automação"
    aba["AN1"] = "DUT de Referência (ID DUT)"
    aba["AO1"] = "Foto 1 DUT Referência"
    aba["AP1"] = "Foto 2 DUT Referência"
    aba["AQ1"] = "Foto 3 DUT Referência"
    aba["AR1"] = "Foto 4 DUT Referência"
    aba["AS1"] = "Foto 5 DUT Referência"
    aba["AT1"] = "Nome do Ambiente"
    aba["AU1"] = "Tipo de Ambiente"
    aba["AV1"] = "ID Ativo"
    aba["AW1"] = "Nome do Ativo"
    aba["AX1"] = "Função"
    aba["AY1"] = "Modelo"
    aba["AZ1"] = "Capacidade Frigorífica"
    aba["BA1"] = "COP"
    aba["BB1"] = "Potência nominal [kW]"
    aba["BC1"] = "Modelo da Evaporadora"
    aba["BD1"] = "Corrente Nominal / RLA do Compressor [A]"
    aba["BE1"] = "Alimentação do Equipamento"
    aba["BF1"] = "Foto 1 Ativo"
    aba["BG1"] = "Foto 2 Ativo"
    aba["BH1"] = "Foto 3 Ativo"
    aba["BI1"] = "Foto 4 Ativo"
    aba["BJ1"] = "Foto 5 Ativo"
    aba["BK1"] = "Dispositivo Diel associado ao ativo (ID DAC/DUT)"
    aba["BL1"] = "Comissionado (S/N)"
    aba["BM1"] = "Sensor P0"
    aba["BN1"] = "P0"
    aba["BO1"] = "Sensor P1"
    aba["BP1"] = "P1"
    aba["BQ1"] = "Foto 1 DAC"
    aba["BR1"] = "Foto 2 DAC"
    aba["BS1"] = "Foto 3 DAC"
    aba["BT1"] = "Foto 4 DAC"
    aba["BU1"] = "Foto 5 DAC"
    aba["BV1"] = "ID do Quadro Elétrico"
    aba["BW1"] = "Nome do Quadro Elétrico"
    aba["BX1"] = "ID Dispositivo Energia"
    aba["BY1"] = "ID Med Energia"
    aba["BZ1"] = "N. Série Medidor"
    aba["CA1"] = "Modelo Medidor"
    aba["CB1"] = "Capacidade TC (A)"
    aba["CC1"] = "Tipo de Instalação Elétrica"
    aba["CD1"] = "Intervalo de Envio (s)"
    aba["CE1"] = "Ambiente Monitorado VAV"
    aba["CF1"] = "Fabricante do Termostato VAV"
    aba["CG1"] = "Modelo do Termostato VAV"
    aba["CH1"] = "Fabricante do Atuador VAV"
    aba["CI1"] = "Modelo do Atuador VAV"
    aba["CJ1"] = "Fabricante da Caixa VAV"
    aba["CK1"] = "Modelo da Caixa VAV"
    aba["CL1"] = "Foto 1 DRI"
    aba["CM1"] = "Foto 2 DRI"
    aba["CN1"] = "Foto 3 DRI"
    aba["CO1"] = "Foto 4 DRI"
    aba["CP1"] = "Foto 5 DRI"
    aba["CQ1"] = "ID do DMA"
    aba["CR1"] = "Hidrômetro"
    aba["CS1"] = "Local de Instalação do dispositivo"
    aba["CT1"] = "Data de Instalação do DMA"
    aba["CU1"] = "Capacidade Total dos Reservatórios (L)"
    aba["CV1"] = "Total de Reservatórios"
    aba["CW1"] = "Foto 1 DMA"
    aba["CX1"] = "Foto 2 DMA"
    aba["CY1"] = "Foto 3 DMA"
    aba["CZ1"] = "Foto 4 DMA"
    aba["DA1"] = "Foto 5 DMA"
    aba["DB1"] = "Id do utilitario"
    aba["DC1"] = "Nome do utilitario"
    aba["DD1"] = "Data de Instalação do Utilitário"
    aba["DE1"] = "Fabricante Nobreak"
    aba["DF1"] = "Modelo Nobreak"
    aba["DG1"] = "Tensão de Entrada (VAC)"
    aba["DH1"] = "Tensão de Saída (VAC)"
    aba["DI1"] = "Potência Nominal (VA)"
    aba["DJ1"] = "Autonomia Nominal (min)"
    aba["DK1"] = "Corrente Elétrica de Entrada (A)"
    aba["DL1"] = "Corrente Elétrica de Saída (A)"
    aba["DM1"] = "Capacidade Nominal da Bateria (Ah)"
    aba["DN1"] = "Tensão da Rede (VAC)"
    aba["DO1"] = "Corrente da Rede Elétrica (A)"
    aba["DP1"] = "Dispositivo associado"
    aba["DQ1"] = "Porta do Dispositivo associado"
    aba["DR1"] = "Feedback do DAL ou DMT"
    aba["DS1"] = "Ativo associado"
    aba["DT1"] = "Foto DMT"
    aba["DU1"] = "Foto DAL ou DMT"
    aba["DV1"] = "Foto Utilitário"
# Busca informações dos dados da unidade informada (Nome, Cidade, UF, LAT/LON, ...) e cria a planilha Excel
def arquivo():
    vt_ri_id()
    vt_ri_forms()
    vt_ri_banco_do_brasil()
    cabecalhosExcel()

    global uf, city
    cliente = unidade_nome[:unidade_nome.find("-")]
    nome_unid = unidade_nome[(unidade_nome).find("-") + 2:]
    end_init = int(str(json_princ["items"][unidade - 1]["address"]["formattedAddress"]).find("-", str(json_princ["items"][unidade - 1]["address"]["formattedAddress"]).find("-") + 2))
    endereco = str(json_princ["items"][unidade - 1]["address"]["formattedAddress"])[end_init + 2:]
    uf = str(json_princ["items"][unidade - 1]["address"]["state"])
    city = str(json_princ["items"][unidade - 1]["address"]["city"])
    lat = str(json_princ["items"][unidade - 1]["address"]["coords"]["latitude"])
    long = str(json_princ["items"][unidade - 1]["address"]["coords"]["longitude"])

    aba["B2"] = nome_unid
    aba["E2"] = uf
    aba["F2"] = city
    aba["I2"] = f"{lat}, {long}"
    aba["J2"] = endereco

    # Salvar arquivo Excel
    unificada.save(f"{unidade_nome} - Planilha unificada.xlsx")
# Localiza os formulários de VT e RI com todos os questionários respondidos, da unidade informada
def vt_ri_id():
    pag = 1
    url = 'https://amonamarth.fieldcontrol.com.br/locations?page=' + str(pag) + '&perPage=100'
    headers = {"Content-Type": "application/json;charset=UTF-8", "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
    req = requests.get(url, headers=headers)

    global json_princ, unidade
    json_princ = req.json()
    unidade = 0
    user = 0


    while user == 0:
        global unidade_nome
        unidade_nome = str(json_princ["items"][unidade]["name"]).upper()
        while not (str.upper(unidade_alvo) in unidade_nome and str.upper(cliente_alvo) in unidade_nome):
            unidade_nome = str(json_princ["items"][unidade]["name"]).upper()
            unidade += 1

            if unidade == 100:
                unidade = 0
                pag += 1
                url = 'https://amonamarth.fieldcontrol.com.br/locations?page=' + str(pag) + '&perPage=100'
                headers = {"Content-Type": "application/json;charset=UTF-8",
                           "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
                req = requests.get(url, headers=headers)
                json_princ = req.json()
        else:
            print("Unidade: " + str(unidade_nome))
            user = int(input("Digite [1] se unidade acima é a correta, [0] se incorreta: "))
            if user == 0:
                unidade = unidade + 1

    else:
        global id_local, vt_id, vt_link, vt_list, vt_link_list, ri_id, ri_link, ri_list, ri_link_list
        id_local = json_princ["items"][unidade - 1]["id"]

    pag = 1
    url = 'https://amonamarth.fieldcontrol.com.br/maintenances?page=' + str(pag) + '&perPage=100&locationIdIn=' + str(id_local)
    headers = {"Content-Type": "application/json;charset=UTF-8",
               "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
    req = requests.get(url, headers=headers)
    json = req.json()
    vt = 0
    ri = 0

    vt_list = []
    vt_link_list = []

    while vt <= len(json["items"])-1:

        if "VISTORIA" in str(json["items"][vt]["message"]).upper():
            vt_id = json["items"][vt]["id"]
            vt_list.append(vt_id)
            try:
                vt_link = "Visita Técnica: " + json["items"][vt]["link"]
            except:
                vt_link = "Não há VT preenchida"
            vt_link_list.append(vt_link)
            vt += 1
            continue

        elif "VISITA" in str(json["items"][vt]["message"]).upper():
            vt_id = json["items"][vt]["id"]
            vt_list.append(vt_id)
            try:
                vt_link = "Visita Técnica: " + json["items"][vt]["link"]
            except:
                vt_link = "Não há VT preenchida"
            vt_link_list.append(vt_link)
            vt += 1
            continue

        elif "VT" in str(json["items"][vt]["message"]).upper():
            vt_id = json["items"][vt]["id"]
            vt_list.append(vt_id)
            try:
                vt_link = "Visita Técnica: " + json["items"][vt]["link"]
            except:
                vt_link = "Não há VT preenchida"
            vt_link_list.append(vt_link)
            vt += 1
            continue

        else:
            vt += 1

    ri_list = []
    ri_link_list = []

    while ri <= len(json["items"])-1:


        if "IN" in json["items"][ri]["message"].upper():
            ri_id = json["items"][ri]["id"]
            ri_list.append(ri_id)
            try:
                ri_link = "Relatório de Instalação: " + json["items"][ri]["link"]
            except:
                ri_link = "Não há relatório preenchido"
            ri_link_list.append(ri_link)
            ri += 1
            continue

        elif "RELAT" in json["items"][ri]["message"].upper():
            if "INS" in str(json["items"][ri]["message"]).upper():
                if "GEST" in str(json["items"][ri]["message"]).upper():
                    if "ATIVO" in str(json["items"][ri]["message"]).upper():
                        ri_id = json["items"][ri]["id"]
                        try:
                            ri_link = "Relatório de Instalação: " + json["items"][ri]["link"]
                        except:
                            ri_link = "Não há relatório preenchido"
                        ri_link_list.append(ri_link)
                        ri += 1
                        continue
        else:
            ri += 1

    print("\nNº de VT(s):", str(len(vt_list)))
    print("Nº de RI(s):", str(len(ri_list)))
# Localiza os questionários respondidos para VT e RI, por tipo (Condensadoras, Evaporadoras, Croqui, DUT, DAC, DAM, ...)
def vt_ri_forms():
# Visita Técnica:
    find_form = 0
    for i in range(len(vt_list)):

        pag = 1
        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + vt_list[i] + '/form-answers?page=' + str(pag) + '&perPage=100'
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        global list_cond, list_evap, list_amb, list_croqui, list_cort, list_energia, list_agua, list_self, list_nobreak, vt_select

        if find_form == 0:

            list_cond = []
            list_evap = []
            list_amb = []
            list_energia = []
            list_agua = []
            list_croqui = []
            list_cort = []
            list_self = []
            list_nobreak = []

            i_cond = 0
            i_evap = 0
            i_amb = 0
            i_croqui = 0
            i_cort = 0
            i_energia = 0
            i_agua = 0
            i_self = 0
            i_nobreak = 0

            while i_cond <= len(json["items"])-1:
                if "CONDENSADORAS" in json["items"][i_cond]["name"]:
                    list_cond.append(json["items"][i_cond]["id"])
                i_cond += 1

            while i_evap <= len(json["items"])-1:
                if "EVAPORADORAS" in json["items"][i_evap]["name"]:
                    list_evap.append(json["items"][i_evap]["id"])
                i_evap += 1

            while i_amb <= len(json["items"])-1:
                if "AMBIENTE" in json["items"][i_amb]["name"]:
                    list_amb.append(json["items"][i_amb]["id"])
                i_amb += 1

            while i_croqui <= len(json["items"])-1:
                if "CROQUI" in json["items"][i_croqui]["name"]:
                    list_croqui.append(json["items"][i_croqui]["id"])
                i_croqui += 1

            while i_cort <= len(json["items"])-1:
                if "CORTINA" in json["items"][i_cort]["name"]:
                    list_cort.append(json["items"][i_cort]["id"])
                i_cort += 1

            while i_energia <= len(json["items"])-1:
                if "ENERGIA" in json["items"][i_energia]["name"]:
                    list_energia.append(json["items"][i_energia]["id"])
                i_energia += 1

            while i_agua <= len(json["items"])-1:
                if "ÁGUA" in json["items"][i_agua]["name"]:
                    list_agua.append(json["items"][i_agua]["id"])
                i_agua += 1

            while i_self <= len(json["items"])-1:
                if "SELF" in json["items"][i_self]["name"]:
                    list_self.append(json["items"][i_self]["id"])
                i_self += 1

            while i_nobreak <= len(json["items"]) - 1:
                if "BREAK" in json["items"][i_nobreak]["name"]:
                    list_nobreak.append(json["items"][i_nobreak]["id"])
                i_nobreak += 1

            if len(list_evap) != 0:
                if len(list_cond) != 0:
                    if len(list_amb) != 0:
                        print("\n")
                        print(vt_link_list[i])
                        print("Condensadoras:", len(list_cond))
                        print("Evaporadoras:", len(list_evap))
                        print("Ambientes:", len(list_amb))
                        print("Croqui:", len(list_croqui))
                        print("Cortina de Ar:", len(list_cort))
                        print("Energia:", len(list_energia))
                        print("Água:", len(list_agua))
                        print("Self:", len(list_self))
                        print("NoBreaks:", len(list_nobreak))
                        print("\n")
                        find_form = 1
                        vt_select = vt_list[i]

# Relatório de Instalação:

    for i in range(len(ri_list)):

        pag = 1
        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + ri_list[i] + '/form-answers?page=' + str(pag) + '&perPage=100'
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        global list_dac, list_dut, list_croq, list_dme, list_dma, list_dam, list_dat, list_rede, list_evap_aux, list_cond_aux, list_cort_aux, ri_select

        list_dac = []
        list_dut = []
        list_croq = []
        list_dme = []
        list_dma = []
        list_dam = []
        list_dat = []
        list_rede = []
        list_cond_aux = []
        list_evap_aux = []
        list_cort_aux = []

        i_dac = 0
        i_dut = 0
        i_croq = 0
        i_dme = 0
        i_dma = 0
        i_dam = 0
        i_dat = 0
        i_rede = 0
        i_evap_aux = 0
        i_cond_aux = 0

        while i_dac <= len(json["items"]) - 1:
            if "DAC" in json["items"][i_dac]["name"] and not ("PRÉ" in str(json["items"][i_dac]["name"]).upper()):
                if not ("VALIDA" in json["items"][i_dac]["name"].upper()):
                    list_dac.append(json["items"][i_dac]["id"])
            i_dac += 1

        while i_dut <= len(json["items"]) - 1:
            if "DUT" in json["items"][i_dut]["name"] and not ("PRÉ" in str(json["items"][i_dut]["name"]).upper()):
                if not ("VALIDA" in json["items"][i_dut]["name"].upper()):
                    list_dut.append(json["items"][i_dut]["id"])
            i_dut += 1

        while i_croq <= len(json["items"]) - 1:
            if "CROQUI" in str(json["items"][i_croq]["name"]).upper():
                list_croq.append(json["items"][i_croq]["id"])
            i_croq += 1

        while i_dme <= len(json["items"]) - 1:
            if "DME" in json["items"][i_dme]["name"] and not ("PRÉ" in str(json["items"][i_dme]["name"]).upper()):
                if not ("VALIDA" in json["items"][i_dme]["name"].upper()):
                    list_dme.append(json["items"][i_dme]["id"])
            i_dme += 1

        while i_dma <= len(json["items"]) - 1:
            if "DMA" in json["items"][i_dma]["name"] and not ("PRÉ" in str(json["items"][i_dma]["name"]).upper()):
                if not ("VALIDA" in json["items"][i_dma]["name"].upper()):
                    list_dma.append(json["items"][i_dma]["id"])
            i_dma += 1

        while i_dam <= len(json["items"]) - 1:
            if "DAM" in json["items"][i_dam]["name"] and not ("PRÉ" in str(json["items"][i_dam]["name"]).upper()):
                if not ("VALIDA" in json["items"][i_dam]["name"].upper()):
                    list_dam.append(json["items"][i_dam]["id"])
            i_dam += 1

        while i_dat <= len(json["items"]) - 1:
            if "DAT" in json["items"][i_dat]["name"] and not ("PRÉ" in str(json["items"][i_dat]["name"]).upper()):
                if not ("VALIDA" in json["items"][i_dat]["name"].upper()):
                    list_dat.append(json["items"][i_dat]["id"])
            i_dat += 1

        while i_rede <= len(json["items"]) - 1:
            if "REDE" in str(json["items"][i_rede]["name"]).upper() and not ("PRÉ" in str(json["items"][i_rede]["name"]).upper()):
                list_rede.append(json["items"][i_rede]["id"])
            i_rede += 1

        while i_cond_aux <= len(json["items"]) - 1:
            if "CONDENSADORAS" in json["items"][i_cond_aux]["name"]:
                list_cond_aux.append(json["items"][i_cond_aux]["id"])
            i_cond_aux += 1

        while i_evap_aux <= len(json["items"]) - 1:
            if "EVAPORADORAS" in json["items"][i_evap_aux]["name"]:
                list_evap_aux.append(json["items"][i_evap_aux]["id"])
            i_evap_aux += 1

        if len(list_dut) != 0:
            print(ri_link_list[i])
            print("DAC:", len(list_dac))
            print("DUT:", len(list_dut))
            print("Croqui:", len(list_croq))
            print("Rede:", len(list_rede))
            print("DME:", len(list_dme))
            print("DMA:", len(list_dma))
            print("DAM:", len(list_dam))
            print("DAT:", len(list_dat))
            print("Evap_Aux:", len(list_evap_aux))
            print("Cond_Aux:", len(list_cond_aux))
            print("Cort_Aux:", len(list_cort_aux))
            ri_select = ri_list[i]
            break
# Obtém dados preenchidos pelo instalador, de cada formulário do Banco do Brasil
def vt_ri_banco_do_brasil():

    global evap_list, dat_evap_list, amb_evap_list, tipo_evap_list, andar_evap_list, foto_evap_list, \
    placa_evap_list, marca_evap_list, model_evap_list, cp_evap_list, un_cp_evap_list, fluido_evap_list, othstate_evap_list, \
    volts_evap_list, inverter_evap_list, local_evap_list, condicao_evap_list, correl_evap_list, coment_evap_list

    global cond_list, dat_cond_list, tipo_cond_list, foto_cond_list, marca_cond_list, model_cond_list, \
    cp_cond_list, un_cp_cond_list, fluido_cond_list, volts_cond_list, inv_cond_list, valv_cond_list, \
    zn_cond_list, local_cond_list, condicao_cond_list, othstate_cond_list, sit_cond_list, correl_cond_list, coment_cond_list

    global amb_list, andar_amb_list, foto_amb_list, energia_list, foto_energia_list, andar_energia_list, \
    tipo_energia_list, foto_energia_list, corrent_energia_list, tensao_energia_list, quadro_energia_list

    global hidro_list, foto_hidro_list, andar_hidro_list, capac_hidro_list, diam_hidro_list, oth_diam_hidro_list, croqui_list, \
    andar_croqui_list, cortina_list, dat_cortina_list, local_cortina_list, foto_cortina_list, marca_cortina_list, tensao_cortina_list, \
    andar_cortina_list, cond_cortina_list

    global self_list, amb_self_list, andar_self_list, tipo_self_list, foto_self_list, cf_self_list, unit_cf_self_list, marca_self_list, \
    model_self_list, fluido_self_list, tensao_self_list, inv_self_list, coment_self_list, nobreak_list, amb_nobreak_list, foto_nobreak_list, \
    marca_nobreak_list, model_nobreak_list, serie_nobreak_list, kva_nobreak_list, bateria_nobreak_list, tipo_carga_nobreak_list

# Listas de Evaporadoras
    evap_list = []
    dat_evap_list = []
    amb_evap_list = []
    tipo_evap_list = []
    andar_evap_list = []
    foto_evap_list = []
    placa_evap_list = []
    marca_evap_list = []
    model_evap_list = []
    cp_evap_list = []
    un_cp_evap_list = []
    fluido_evap_list = []
    volts_evap_list = []
    inverter_evap_list = []
    local_evap_list = []
    condicao_evap_list = []
    othstate_evap_list = []
    correl_evap_list = []
    coment_evap_list = []

# Listas das Condensadoras
    cond_list = []
    dat_cond_list = []
    tipo_cond_list = []
    foto_cond_list = []
    marca_cond_list = []
    model_cond_list = []
    cp_cond_list = []
    un_cp_cond_list = []
    fluido_cond_list = []
    volts_cond_list = []
    inv_cond_list = []
    valv_cond_list = []
    zn_cond_list = []
    local_cond_list = []
    condicao_cond_list = []
    othstate_cond_list = []
    sit_cond_list = []
    correl_cond_list = []
    coment_cond_list = []

# Listas dos Ambientes:
    amb_list = []
    andar_amb_list = []
    foto_amb_list = []

# Listas de Energia:
    energia_list = []
    andar_energia_list = []
    tipo_energia_list = []
    corrent_energia_list = []
    tensao_energia_list = []
    quadro_energia_list = []
    foto_energia_list = []

# Listas de Água:
    hidro_list = []
    foto_hidro_list = []
    andar_hidro_list = []
    capac_hidro_list = []
    diam_hidro_list = []
    oth_diam_hidro_list = []

# Listas de Croqui
    croqui_list = []
    andar_croqui_list = []

# Listas de Cortina:
    cortina_list = []
    dat_cortina_list = []
    local_cortina_list = []
    foto_cortina_list = []
    marca_cortina_list = []
    tensao_cortina_list = []
    andar_cortina_list = []
    cond_cortina_list = []

# Listas de Self's:
    self_list = []
    amb_self_list = []
    andar_self_list = []
    tipo_self_list = []
    foto_self_list = []
    cf_self_list = []
    unit_cf_self_list = []
    marca_self_list = []
    model_self_list = []
    fluido_self_list = []
    tensao_self_list = []
    inv_self_list = []
    coment_self_list = []

# Listas de No-Breaks:
    nobreak_list = []
    amb_nobreak_list = []
    foto_nobreak_list = []
    marca_nobreak_list = []
    model_nobreak_list = []
    serie_nobreak_list = []
    kva_nobreak_list = []
    bateria_nobreak_list = []
    tipo_carga_nobreak_list = []

# Evaporadoras:

    for i in range(len(list_evap)):
        opc1 = 0
        opc2 = 0

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + vt_select + '/form-answers/' + list_evap[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            # TAG da Evaporadora:
            for i in range(len(json["questions"])):
                if opc2 == 0:
                    if "CÓDIGO" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            evap_list.append(json["questions"][i]["answer"])
                            opc1 = 1

            # TAG da Evaporadora:
            for i in range(len(json["questions"])):
                if opc1 == 0:
                    if "TAG" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            evap_list.append(json["questions"][i]["answer"])
                            opc2 = 1

            # ID dos DAT's - Evaporadoras:
            for i in range(len(json["questions"])):
                if "DAT" in str(json["questions"][i]["title"]):
                    if not "FOTO" in str(json["questions"][i]["title"]).upper():
                        if "CÓDIGO" in str(json["questions"][i]["title"]).upper():
                            if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                dat_evap_list.append(json["questions"][i]["answer"])
                            else:
                                dat_evap_list.append("None")

            # Ambiente climatizado pelas Evaporadoras:
            for i in range(len(json["questions"])):
                if "AMBIENTE" in str(json["questions"][i]["title"]).upper():
                    if "EQUIPAMENTO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            amb_evap_list.append(json["questions"][i]["answer"])

            # Andar das Evaporadoras:
            for i in range(len(json["questions"])):
                if "ANDAR" in str(json["questions"][i]["title"]).upper():
                    if "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            andar_evap_list.append(json["questions"][i]["answer"])

            # Tipo das Evaporadoras:
            for i in range(len(json["questions"])):
                if "TIPO" in str(json["questions"][i]["title"]).upper():
                    if "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            tipo_evap_list.append(json["questions"][i]["answer"])

            # Fotos das Evaporadoras:
            for i in range(len(json["questions"])):
                if "FOTO" in str(json["questions"][i]["title"]).upper():
                    if "EVAPORADORA" in str(json["questions"][i]["title"]).upper():
                        if "AMBIENTE" in str(json["questions"][i]["title"]).upper():
                            if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                foto_evap_list.append(json["questions"][i]["answer"])

            # Plaquetas de dados das Evaporadoras:
            for i in range(len(json["questions"])):
                if "PLAQUETA" in str(json["questions"][i]["title"]).upper():
                    if "FOTO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            placa_evap_list.append(json["questions"][i]["answer"])

            # Marcas das Evaporadoras:
            for i in range(len(json["questions"])):
                if "MARCA" in str(json["questions"][i]["title"]).upper():
                    if "EQUIPAMENTO" in str(json["questions"][i]["title"]).upper():
                        if not "INSIRA" in str(json["questions"][i]["title"]).upper():
                            if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                marca_evap_list.append(json["questions"][i]["answer"])

            # Modelos das Evaporadoras:
            for i in range(len(json["questions"])):
                if "MODELO" in str(json["questions"][i]["title"]).upper():
                    if "EQUIPAMENTO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            model_evap_list.append(json["questions"][i]["answer"])

            # Capacidades Frigoríficas das Evaporadoras:
            for i in range(len(json["questions"])):
                if "FRIGORÍFICA" in str(json["questions"][i]["title"]).upper():
                    if not "UNIDADE" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            cp_evap_list.append(json["questions"][i]["answer"])

            # Unidades das CP informadas:
            for i in range(len(json["questions"])):
                if "FRIGORÍFICA" in str(json["questions"][i]["title"]).upper():
                    if "UNIDADE" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            un_cp_evap_list.append(json["questions"][i]["answer"])

            # Fluidos Refrigerantes das Evaporadoras:
            for i in range(len(json["questions"])):
                if "GÁS" in str(json["questions"][i]["title"]).upper():
                    if "REFRIGERANTE" in str(json["questions"][i]["title"]).upper():
                        if not "LISTADO" in str(json["questions"][i]["title"]).upper():
                            if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                fluido_evap_list.append(json["questions"][i]["answer"])

            # Tensões de alimentação das Evaporadoras:
            for i in range(len(json["questions"])):
                if "TENSÃO" in str(json["questions"][i]["title"]).upper():
                    if "ALIMENTAÇÃO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            volts_evap_list.append(json["questions"][i]["answer"])

            # Evaporadoras inverter:
            for i in range(len(json["questions"])):
                if "EQUIPAMENTO" in str(json["questions"][i]["title"]).upper():
                    if "INVERTER" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            inverter_evap_list.append(json["questions"][i]["answer"])

            # Locais das Evaporadoras:
            for i in range(len(json["questions"])):
                if "LOCALIZAÇÃO" in str(json["questions"][i]["title"]).upper():
                    if "EVAPORADORA" in str(json["questions"][i]["title"]).upper():
                        local_evap_list.append(json["questions"][i]["answer"])

            # Condições das Evaporadoras:
            for i in range(len(json["questions"])):
                if "CONDIÇÃO" in str(json["questions"][i]["title"]).upper():
                    if "EVAPORADORA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            condicao_evap_list.append(json["questions"][i]["answer"])

            # Outras condições das Evaporadoras:
            for i in range(len(json["questions"])):
                if "OUTRA" in str(json["questions"][i]["title"]).upper():
                    if "CONDIÇÃO" in str(json["questions"][i]["title"]).upper():
                        othstate_evap_list.append(json["questions"][i]["answer"])

            # Correlacionamento das Evaporadoras:
            for i in range(len(json["questions"])):
                if "CORRELAC" in str(json["questions"][i]["title"]).upper():
                    if "CONDENSADORA" in str(json["questions"][i]["title"]).upper():
                        if "q" in str(json["questions"][i]["title"]):
                            if "?" in str(json["questions"][i]["title"]):
                                if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                    correl_evap_list.append(json["questions"][i]["answer"])

            # Comentários Gerais das Evaporadoras:
            for i in range(len(json["questions"])):
                if "COMENTÁRIOS" in str(json["questions"][i]["title"]).upper():
                    if "GERAIS" in str(json["questions"][i]["title"]).upper():
                        coment_evap_list.append(json["questions"][i]["answer"])

# Condensadoras:

    for i in range(len(list_cond)):
        opc1 = 0
        opc2 = 0

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + vt_select + '/form-answers/' + list_cond[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            # TAG da Condensadora:
            for i in range(len(json["questions"])):
                if opc2 == 0:
                    if "CÓDIGO" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            cond_list.append(json["questions"][i]["answer"])
                            opc1 = 1

            # TAG da Condensadora:
            for i in range(len(json["questions"])):
                if opc1 == 0:
                    if "TAG" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            cond_list.append(json["questions"][i]["answer"])
                            opc2 = 1

            # ID dos DAT's - Condensadoras:
            for i in range(len(json["questions"])):
                if "DAT" in str(json["questions"][i]["title"]):
                    if not "FOTO" in str(json["questions"][i]["title"]).upper():
                        if "CÓDIGO" in str(json["questions"][i]["title"]).upper():
                            if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                dat_cond_list.append(json["questions"][i]["answer"])
                            else:
                                dat_cond_list.append("None")

            # Fotos das Condensadoras:
            for i in range(len(json["questions"])):
                if "FOTO" in str(json["questions"][i]["title"]).upper():
                    if "CONDENSADORA" in str(json["questions"][i]["title"]).upper():
                        if "AMPLA" in str(json["questions"][i]["title"]).upper():
                            if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                foto_cond_list.append(json["questions"][i]["answer"])

            # Tipo das Condensadoras:
            for i in range(len(json["questions"])):
                if "TIPO" in str(json["questions"][i]["title"]).upper():
                    if "CONDENSADORA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str( json["questions"][i]["required"]):
                            tipo_cond_list.append(json["questions"][i]["answer"])

            # Marca das Condensadoras:
            for i in range(len(json["questions"])):
                if "MARCA" in str(json["questions"][i]["title"]).upper():
                    if "EQUIPAMENTO" in str(json["questions"][i]["title"]).upper():
                        if "?" in str(json["questions"][i]["title"]).upper():
                            if str(json["questions"][i]["answer"]) != "None" and "True" == str( json["questions"][i]["required"]):
                                marca_cond_list.append(json["questions"][i]["answer"])

            # Modelos das Condensadoras:
            for i in range(len(json["questions"])):
                if "MODELO" in str(json["questions"][i]["title"]).upper():
                    if "EQUIPAMENTO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str( json["questions"][i]["required"]):
                            model_cond_list.append(json["questions"][i]["answer"])

            # Capacidades Frigoríficas das Condensadoras:
            for i in range(len(json["questions"])):
                if "FRIGORÍFICA" in str(json["questions"][i]["title"]).upper():
                    if not "UNIDADE" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str( json["questions"][i]["required"]):
                            cp_cond_list.append(json["questions"][i]["answer"])

            # Unidades das CP's das Condensadoras:
            for i in range(len(json["questions"])):
                if "FRIGORÍFICA" in str(json["questions"][i]["title"]).upper():
                    if "UNIDADE" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str( json["questions"][i]["required"]):
                            un_cp_cond_list.append(json["questions"][i]["answer"])

            # Fluidos Refrigerantes das Condensadoras:
            for i in range(len(json["questions"])):
                if "GÁS" in str(json["questions"][i]["title"]).upper():
                    if "REFRIGERANTE" in str(json["questions"][i]["title"]).upper():
                        if not "NÃO" in str(json["questions"][i]["title"]).upper():
                            if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                fluido_cond_list.append(json["questions"][i]["answer"])

            # Tensões de alimentação das Condensadoras:
            for i in range(len(json["questions"])):
                if "TENSÃO" in str(json["questions"][i]["title"]).upper():
                    if "ALIMENTAÇÃO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str( json["questions"][i]["required"]):
                            volts_cond_list.append(json["questions"][i]["answer"])

            # Condensadoras inverter:
            for i in range(len(json["questions"])):
                if "INVERTER" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                        inv_cond_list.append(json["questions"][i]["answer"])

            # Diâmetros das válvulas:
            for i in range(len(json["questions"])):
                if "DIÂMETRO" in str(json["questions"][i]["title"]).upper():
                    if "VÁLVULA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            valv_cond_list.append(json["questions"][i]["answer"])

            # Zonas técnicas das Condensadoras:
            for i in range(len(json["questions"])):
                if "ZONA" in str(json["questions"][i]["title"]).upper():
                    if "TÉCNICA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            zn_cond_list.append(json["questions"][i]["answer"])

            # Locais das Condensadoras:
            for i in range(len(json["questions"])):
                if "ZONA" in str(json["questions"][i]["title"]).upper() and "CONDENSADORA" in str(json["questions"][i]["title"]).upper():
                        if "TÉCNICA" in str(json["questions"][i]["title"]).upper():
                            if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                local_cond_list.append(json["questions"][i]["answer"])

            # Condições das Condensadoras:
            for i in range(len(json["questions"])):
                if "CONDIÇÃO" in str(json["questions"][i]["title"]).upper():
                    if "CONDENSADORA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            condicao_cond_list.append(json["questions"][i]["answer"])

            # Outras condições das Condensadoras:
            for i in range(len(json["questions"])):
                if "OUTRA" in str(json["questions"][i]["title"]).upper():
                    if "CONDIÇÃO" in str(json["questions"][i]["title"]).upper():
                        othstate_cond_list.append(json["questions"][i]["answer"])

            # Situações das Condensadoras:
            for i in range(len(json["questions"])):
                if "SITUAÇÃO" in str(json["questions"][i]["title"]).upper():
                    if "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str( json["questions"][i]["required"]):
                            sit_cond_list.append(json["questions"][i]["answer"])

            # Correlacionamento das Condensadoras:
            for i in range(len(json["questions"])):
                if "CORRELAC" in str(json["questions"][i]["title"]).upper():
                    if "EVAPORADORA" in str(json["questions"][i]["title"]).upper():
                        if not "(S)" in str(json["questions"][i]["title"]).upper():
                            if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                correl_cond_list.append(json["questions"][i]["answer"])

            # Comentários Gerais das Condensadoras:
            for i in range(len(json["questions"])):
                if "COMENTÁRIOS" in str(json["questions"][i]["title"]).upper():
                    if "GERAIS" in str(json["questions"][i]["title"]).upper():
                        coment_cond_list.append(json["questions"][i]["answer"])

# Ambientes:

    for i in range(len(list_amb)):
        opc1 = 0
        opc2 = 0

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + vt_select + '/form-answers/' + list_amb[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()
        if not "statusCode" in str(json):
            # TAG's do Ambientes:
            for i in range(len(json["questions"])):
                if opc2 == 0:
                    if "CÓDIGO" in str(json["questions"][i]["title"]).upper() and "AMBIENTE" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            amb_list.append(json["questions"][i]["answer"])
                            opc1 = 1

            # TAG's dos Ambientes:
            for i in range(len(json["questions"])):
                if opc1 == 0:
                    if "TAG" in str(json["questions"][i]["title"]).upper() and "AMBIENTE" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            amb_list.append(json["questions"][i]["answer"])
                            opc2 = 1

            # Andares dos Ambientes:
            for i in range(len(json["questions"])):
                if "TIPO" in str(json["questions"][i]["title"]).upper() and "AMBIENTE" in str(json["questions"][i]["title"]).upper():
                    if "CLIMA" in str(json["questions"][i]["answer"]).upper():
                        for item in range(len(json["questions"])):
                            if "ANDAR" in str(json["questions"][item]["title"]).upper():
                                if "AMBIENTE" in str(json["questions"][item]["title"]).upper():
                                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                        andar_amb_list.append(json["questions"][item]["answer"])

            # Fotos dos Ambientes:
            for i in range(len(json["questions"])):
                if "TIPO" in str(json["questions"][i]["title"]).upper() and "AMBIENTE" in str(json["questions"][i]["title"]).upper():
                    if "CLIMA" in str(json["questions"][i]["answer"]).upper():
                        for item in range(len(json["questions"])):
                            if "FOTO" in str(json["questions"][item]["title"]).upper():
                                if "AMBIENTE" in str(json["questions"][item]["title"]).upper():
                                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                        foto_amb_list.append(json["questions"][item]["answer"])

# Energia:

    for i in range(len(list_energia)):

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + vt_select + '/form-answers/' + list_energia[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            # TAG da Energia:
            for i in range(len(json["questions"])):
                    if "TAG" in str(json["questions"][i]["title"]).upper() and "QUADRO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            energia_list.append(json["questions"][i]["answer"])

            # Andares da Energia:
            for i in range(len(json["questions"])):
                    if "ANDAR" in str(json["questions"][i]["title"]).upper() and "QUADRO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            andar_energia_list.append(json["questions"][i]["answer"])

            # Fotos da Energia:
            for i in range(len(json["questions"])):
                    if "FOTO" in str(json["questions"][i]["title"]).upper() and "QUADRO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            foto_energia_list.append(json["questions"][i]["answer"])

            # Quadro de instalação:
            for i in range(len(json["questions"])):
                    if "TIPO" in str(json["questions"][i]["title"]).upper() and "QUADRO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            quadro_energia_list.append(json["questions"][i]["answer"])

            # Tipos de instalação:
            for i in range(len(json["questions"])):
                    if "TIPO" in str(json["questions"][i]["title"]).upper() and "INSTALA" in str(json["questions"][i]["title"]).upper():
                        tipo_energia_list.append(json["questions"][i]["answer"])

            # Tensões de entrada:
            for i in range(len(json["questions"])):
                    if "TENSÃO" in str(json["questions"][i]["title"]).upper() and "ENTRADA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            tensao_energia_list.append(json["questions"][i]["answer"])

            # Correntes dos disjuntores:
            for i in range(len(json["questions"])):
                    if "CORRENTE" in str(json["questions"][i]["title"]).upper() and "DISJUNTOR" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            corrent_energia_list.append(json["questions"][i]["answer"])

# Água:

    for i in range(len(list_agua)):

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + vt_select + '/form-answers/' + list_agua[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            # Hidrômetro:
            for i in range(len(json["questions"])):
                    if "SITUAÇÃO" in str(json["questions"][i]["title"]).upper() and "HIDRÔMETRO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            hidro_list.append(json["questions"][i]["answer"])

            # Fotos dos Hidrômetros:
            for i in range(len(json["questions"])):
                    if "FOTO" in str(json["questions"][i]["title"]).upper() and "HIDRÔMETRO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            foto_hidro_list.append(json["questions"][i]["answer"])

            # Diâmetro da tubulação:
            for i in range(len(json["questions"])):
                    if "DIÂMETRO" in str(json["questions"][i]["title"]).upper() and "TUBULAÇÃO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            if "?" in str(json["questions"][i]["title"]) and not "CAIXA" in str(json["questions"][i]["title"]).upper():
                                diam_hidro_list.append(json["questions"][i]["answer"])

            # Outros diâmetros de tubulação:
            for i in range(len(json["questions"])):
                    if "INFORME" in str(json["questions"][i]["title"]).upper() and "DIÂMETRO" in str(json["questions"][i]["title"]).upper():
                        oth_diam_hidro_list.append(json["questions"][i]["answer"])

            # Local do hidrômetro:
            for i in range(len(json["questions"])):
                    if "LOCALIZADO" in str(json["questions"][i]["title"]).upper() and "MEDIDOR" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            andar_hidro_list.append(json["questions"][i]["answer"])

            # Capacidade do reservatório:
            for i in range(len(json["questions"])):
                    if "CAPACIDADE" in str(json["questions"][i]["title"]).upper() and "RESERVATÓRIO" in str(json["questions"][i]["title"]).upper():
                        capac_hidro_list.append(json["questions"][i]["answer"])

# Croqui:

    for i in range(len(list_croqui)):

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + vt_select + '/form-answers/' + list_croqui[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            # Andares dos Croqui's:
            for i in range(len(json["questions"])):
                    if "ANDAR" in str(json["questions"][i]["title"]).upper() and "CROQUI" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            andar_croqui_list.append(json["questions"][i]["answer"])

            # Fotos dos Croqui's:
            for i in range(len(json["questions"])):
                    if "FOTO" in str(json["questions"][i]["title"]).upper() and "PLANTA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            if len(croqui_list) < 5:
                                croqui_list.append(json["questions"][i]["answer"])

# Cortina de Ar:

    for i in range(len(list_cort)):

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + vt_select + '/form-answers/' + list_cort[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            # Códigos DIEL para Cortinas:
            for i in range(len(json["questions"])):
                    if "DIEL" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            cortina_list.append(json["questions"][i]["answer"])

            # DAT's das Cortinas:
            for i in range(len(json["questions"])):
                    if "CÓDIGO" in str(json["questions"][i]["title"]).upper() and "DAT" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            dat_cortina_list.append(json["questions"][i]["answer"])

            # Locais das Cortinas:
            for i in range(len(json["questions"])):
                    if "PORTA" in str(json["questions"][i]["title"]).upper() and "ATUA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            local_cortina_list.append(json["questions"][i]["answer"])

            # Fotos das Cortinas:
            for i in range(len(json["questions"])):
                    if "FOTO" in str(json["questions"][i]["title"]).upper() and "CORTINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            foto_cortina_list.append(json["questions"][i]["answer"])

            # Marcas das Cortinas:
            for i in range(len(json["questions"])):
                    if "MARCA" in str(json["questions"][i]["title"]).upper() and "EQUIPAMENTO" in str(json["questions"][i]["title"]).upper():
                        marca_cortina_list.append(json["questions"][i]["answer"])

            # Tensões das Cortinas:
            for i in range(len(json["questions"])):
                    if "TENSÃO" in str(json["questions"][i]["title"]).upper() and "ALIMENT" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            tensao_cortina_list.append(json["questions"][i]["answer"])

            # Andares das Cortinas:
            for i in range(len(json["questions"])):
                    if "ANDAR" in str(json["questions"][i]["title"]).upper() and "CORTINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            andar_cortina_list.append(json["questions"][i]["answer"])

            # Condições das Cortinas:
            for i in range(len(json["questions"])):
                    if "CONDIÇÃO" in str(json["questions"][i]["title"]).upper() and "CORTINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            cond_cortina_list.append(json["questions"][i]["answer"])

# Self Contained:

    for i in range(len(list_self)):

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + vt_select + '/form-answers/' + list_self[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            # Códigos DIEL para Cortinas:
            for i in range(len(json["questions"])):
                    if "DIEL" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            self_list.append(json["questions"][i]["answer"])

            # Ambiente climatizado pelo Self:
            for i in range(len(json["questions"])):
                    if "AMBIENTE" in str(json["questions"][i]["title"]).upper() and "CLIMATIZA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            amb_self_list.append(json["questions"][i]["answer"])

            # Andares dos Self's:
            for i in range(len(json["questions"])):
                    if "ANDAR" in str(json["questions"][i]["title"]).upper() and "AMBIENTE" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            andar_self_list.append(json["questions"][i]["answer"])

            # Tipos de Máquinas:
            for i in range(len(json["questions"])):
                    if "TIPO" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            tipo_self_list.append(json["questions"][i]["answer"])

            # Fotos do Ambiente das Máquinas:
            for i in range(len(json["questions"])):
                    if "FOTO" in str(json["questions"][i]["title"]).upper() and "AMBIENTE" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            foto_self_list.append(json["questions"][i]["answer"])

            # Capacidade Frigorífica:
            for i in range(len(json["questions"])):
                    if "FRIGORÍFICA" in str(json["questions"][i]["title"]).upper() and "CAPACIDADE" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            cf_self_list.append(json["questions"][i]["answer"])

            # Unidades da CP's:
            for i in range(len(json["questions"])):
                    if "UNIDADE" in str(json["questions"][i]["title"]).upper() and "FRIGORÍFICA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            unit_cf_self_list.append(json["questions"][i]["answer"])

            # Marca do Equipamento:
            for i in range(len(json["questions"])):
                    if "MARCA" in str(json["questions"][i]["title"]).upper() and "EQUIPAMENTO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            marca_self_list.append(json["questions"][i]["answer"])

            # Modelo do Equipamento:
            for i in range(len(json["questions"])):
                    if "MODELO" in str(json["questions"][i]["title"]).upper() and "EQUIPAMENTO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            model_self_list.append(json["questions"][i]["answer"])

            # Fluidos Refrigerantes:
            for i in range(len(json["questions"])):
                    if "GÁS" in str(json["questions"][i]["title"]).upper() and "REFRIGERANTE" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            fluido_self_list.append(json["questions"][i]["answer"])

            # Tensões de Alimentação:
            for i in range(len(json["questions"])):
                    if "TENSÃO" in str(json["questions"][i]["title"]).upper() and "ALIMENTAÇ" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            tensao_self_list.append(json["questions"][i]["answer"])

            # Inverter?
            for i in range(len(json["questions"])):
                    if "INVERTER" in str(json["questions"][i]["title"]).upper() and "EQUIPAMENTO" in str(json["questions"][i]["title"]).upper():
                        inv_self_list.append(json["questions"][i]["answer"])

            # Comentários Gerais:
            for i in range(len(json["questions"])):
                    if "COMENTÁRIOS" in str(json["questions"][i]["title"]).upper() and "GERAIS" in str(json["questions"][i]["title"]).upper():
                        coment_self_list.append(json["questions"][i]["answer"])

# NoBreak's:

    for i in range(len(list_nobreak)):

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + vt_select + '/form-answers/' + list_nobreak[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            # Códigos DIEL para No-Breaks:
            for i in range(len(json["questions"])):
                    if "TAG" in str(json["questions"][i]["title"]).upper() and "NOBREAK" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            nobreak_list.append(json["questions"][i]["answer"])

            # Ambientes dos No-Breaks:
            for i in range(len(json["questions"])):
                    if "AMBIENTE" in str(json["questions"][i]["title"]).upper() and "NOBREAK" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            amb_nobreak_list.append(json["questions"][i]["answer"])

            # Fotos das Etiquetas Técnicas:
            for i in range(len(json["questions"])):
                    if "FOTO" in str(json["questions"][i]["title"]).upper() and "ETIQUETA" in str(json["questions"][i]["title"]).upper():
                        if not "ALIMENTA" in str(json["questions"][i]["title"]).upper():
                            if not "PATRIMONIAL" in str(json["questions"][i]["title"]).upper():
                                if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                    foto_nobreak_list.append(json["questions"][i]["answer"])

            # Marcas dos No-Breaks:
            for i in range(len(json["questions"])):
                    if "MARCA" in str(json["questions"][i]["title"]).upper():
                        marca_nobreak_list.append(json["questions"][i]["answer"])

            # Modelos dos No-Breaks:
            for i in range(len(json["questions"])):
                    if "MODELO" in str(json["questions"][i]["title"]).upper():
                        model_nobreak_list.append(json["questions"][i]["answer"])

            # Número de Série dos No-Breaks:
            for i in range(len(json["questions"])):
                    if "NÚMERO" in str(json["questions"][i]["title"]).upper() and "SÉRIE" in str(json["questions"][i]["title"]).upper():
                        serie_nobreak_list.append(json["questions"][i]["answer"])

            # Capacidades Nominais (KVA):
            for i in range(len(json["questions"])):
                    if "KVA" in str(json["questions"][i]["title"]).upper() and "CAPACIDADE" in str(json["questions"][i]["title"]).upper():
                        kva_nobreak_list.append(json["questions"][i]["answer"])

            # Capacidades das Baterias (Ah):
            for i in range(len(json["questions"])):
                    if "BATERIA" in str(json["questions"][i]["title"]).upper() and "CAPACIDADE" in str(json["questions"][i]["title"]).upper():
                        bateria_nobreak_list.append(json["questions"][i]["answer"])

            # Tipos de Cargas dos No-Breaks:
            for i in range(len(json["questions"])):
                    if "CARGA" in str(json["questions"][i]["title"]).upper() and "TIPO" in str(json["questions"][i]["title"]).upper():
                        if not "FOTO" in str(json["questions"][i]["title"]).upper():
                            if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                tipo_carga_nobreak_list.append(json["questions"][i]["answer"])
            '''
            print(evap_list)
            print(dat_evap_list)
            print(amb_evap_list)
            print(tipo_evap_list)
            print(andar_evap_list)
            print(foto_evap_list)
            print(placa_evap_list)
            print(marca_evap_list)
            print(model_evap_list)
            print(cp_evap_list)
            print(un_cp_evap_list)
            print(fluido_evap_list)
            print(volts_evap_list)
            print(inverter_evap_list)
            print(local_evap_list)
            print(condicao_evap_list)
            print(othstate_evap_list)
            print(correl_evap_list)
            print(coment_evap_list)
            print("\n")
            print("\n")
            print(cond_list)
            print(dat_cond_list)
            print(tipo_cond_list)
            print(foto_cond_list)
            print(marca_cond_list)
            print(model_cond_list)
            print(cp_cond_list)
            print(un_cp_cond_list)
            print(fluido_cond_list)
            print(volts_cond_list)
            print(inv_cond_list)
            print(valv_cond_list)
            print(zn_cond_list)
            print(local_cond_list)
            print(condicao_cond_list)
            print(othstate_cond_list)
            print(sit_cond_list)
            print(correl_cond_list)
            print(coment_cond_list)
            print("\n")
            print("\n")
            print(amb_list)
            print(andar_amb_list)
            print(foto_amb_list)
            print("\n")
            print("\n")
            print(energia_list)
            print(andar_energia_list)
            print(tipo_energia_list)
            print(corrent_energia_list)
            print(tensao_energia_list)
            print(quadro_energia_list)
            print(foto_energia_list)
            print("\n")
            print("\n")
            print(hidro_list)
            print(foto_hidro_list)
            print(andar_hidro_list)
            print(capac_hidro_list)
            print(diam_hidro_list)
            print(oth_diam_hidro_list)
            print("\n")
            print("\n")
            print(croqui_list)
            print(andar_croqui_list)
            print("\n")
            print("\n")
            print(cortina_list)
            print(dat_cortina_list)
            print(local_cortina_list)
            print(foto_cortina_list)
            print(marca_cortina_list)
            print(tensao_cortina_list)
            print(andar_cortina_list)
            print(cond_cortina_list)
            print("\n")
            print("\n")
            print(self_list)
            print(amb_self_list)
            print(andar_self_list)
            print(tipo_self_list)
            print(foto_self_list)
            print(cf_self_list)
            print(unit_cf_self_list)
            print(marca_self_list)
            print(model_self_list)
            print(fluido_self_list)
            print(tensao_self_list)
            print(inv_self_list)
            print(coment_self_list)
            print("\n")
            print("\n")
            print(nobreak_list)
            print(amb_nobreak_list)
            print(foto_nobreak_list)
            print(marca_nobreak_list)
            print(model_nobreak_list)
            print(serie_nobreak_list)
            print(kva_nobreak_list)
            print(bateria_nobreak_list)
            print(tipo_carga_nobreak_list)
            '''

    global tag_dac_list, maq_dac_list, id_dac_list, foto_id_dac_list, valv_dac_list, foto_dac_list , coment_dac_list, \
    tag_dut_list, tipo_dut_list, id_dut_list, foto_id_dut_list, foto_dut_list, amb_dut_list, coment_dut_list, tag_ref_list, \
    tag_ref_list, tipo_ref_list, id_ref_list, foto_id_ref_list, foto_ref_list, amb_ref_list, and_croqui_list, foto_croqui_list, \
    tipo_dme_list, id_dme_list, foto_id_dme_list, inst_dme_list, foto_dme_list, ampere_dme_list, tc_dme_list, \
    quadro_dme_list, tipo_dma_list, local_dma_list, id_dma_list, foto_id_dma_list, foto_dma_list, tag_dam_list, tipo_dam_list, \
    id_dam_list, foto_id_dam_list, foto_dam_list, id_dat_list, foto_id_dat_list, ativo_dat_list, foto_dat_list, id_rede_list, \
    tipo_rede_list, amb_rede_list, foto_rede_list, foto_id_rede_list, dat_dac_list, dat_dut_list, dat_dam_list, evap_aux_list, cond_aux_list,\
    cort_aux_list, tag_evap_aux_list, tag_cond_aux_list, tag_cort_aux_list

# Listas para DAC:
    dat_dac_list = []
    tag_dac_list = []
    maq_dac_list = []
    id_dac_list = []
    foto_id_dac_list = []
    valv_dac_list = []
    foto_dac_list = []
    coment_dac_list = []

# Listas para DUT:
    dat_dut_list = []
    tag_dut_list = []
    tipo_dut_list = []
    id_dut_list = []
    foto_id_dut_list = []
    foto_dut_list = []
    amb_dut_list = []
    coment_dut_list = []

# Listas para REF:
    tag_ref_list = []
    tipo_ref_list = []
    id_ref_list = []
    foto_id_ref_list = []
    foto_ref_list = []
    amb_ref_list = []

# Listas para Croqui:
    and_croqui_list = []
    foto_croqui_list = []

# Listas para DME:
    tipo_dme_list = []
    id_dme_list = []
    foto_id_dme_list = []
    inst_dme_list = []
    foto_dme_list = []
    ampere_dme_list = []
    tc_dme_list = []
    quadro_dme_list = []

# Listas para DMA:
    tipo_dma_list = []
    local_dma_list = []
    id_dma_list = []
    foto_id_dma_list = []
    foto_dma_list = []

# Listas para DAM:
    dat_dam_list = []
    tag_dam_list = []
    tipo_dam_list = []
    id_dam_list = []
    foto_id_dam_list = []
    foto_dam_list = []

# Listas para DAT:
    id_dat_list = []
    foto_id_dat_list = []
    ativo_dat_list = []
    foto_dat_list = []

# Listas para Rede:
    id_rede_list = []
    foto_id_rede_list = []
    tipo_rede_list = []
    amb_rede_list = []
    foto_rede_list = []

# Listas para Forms de Evaporadoras, Condensadoras e Cortinas Auxiliares:
    evap_aux_list = []
    tag_evap_aux_list = []
    cond_aux_list = []
    tag_cond_aux_list = []
    cort_aux_list = []
    tag_cort_aux_list = []

# Dispositivos DAC's:

    for i in range(len(list_dac)):

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + ri_select + '/form-answers/' + list_dac[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            # Códigos DIEL para Máquinas:
            for i in range(len(json["questions"])):
                if "CÓDIGO" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                        tag_dac_list.append(json["questions"][i]["answer"])

            # Tipos de Máquinas dos DAC's:
            for i in range(len(json["questions"])):
                if "TIPO" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper() and "DAC" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                        maq_dac_list.append(json["questions"][i]["answer"])

            # ID's dos DAC's instalados:
            for i in range(len(json["questions"])):
                if "ID" in str(json["questions"][i]["title"]) and "DAC" in str(json["questions"][i]["title"]):
                    if not "FOTO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            id_dac_list.append(json["questions"][i]["answer"])

            # Fotos dos ID's dos DAC's:
            for i in range(len(json["questions"])):
                if "ID" in str(json["questions"][i]["title"]) and "DAC" in str(json["questions"][i]["title"]):
                    if "FOTO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            foto_id_dac_list.append(json["questions"][i]["answer"])

            # Diâmetros dos Engates de Pressão:
            for i in range(len(json["questions"])):
                if "ENGATE" in str(json["questions"][i]["title"]).upper() and "PRESSÃO" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                        valv_dac_list.append(json["questions"][i]["answer"])

            # Fotos das máquinas com DAC's:
            for i in range(len(json["questions"])):
                if "FOTO" in str(json["questions"][i]["title"]).upper() and "AMPLA" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                        foto_dac_list.append(json["questions"][i]["answer"])

            # Comentários Gerais das Condensadoras:
            for i in range(len(json["questions"])):
                if "COMENTÁRIOS" in str(json["questions"][i]["title"]).upper() and "INST" in str(json["questions"][i]["title"]).upper():
                    coment_dac_list.append(json["questions"][i]["answer"])

            # DAT(s) informado(s) na instalação:
            for i in range(len(json["questions"])):
                if "CÓDIGO" in str(json["questions"][i]["title"]).upper() and "DAT" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                        dat_dac_list.append(json["questions"][i]["answer"])

        '''print(tag_dac_list)
        print(maq_dac_list)
        print(id_dac_list)
        print(foto_id_dac_list)
        print(valv_dac_list)
        print(foto_dac_list)
        print(coment_dac_list)'''

# Dispositivos DUT's:

    for i in range(len(list_dut)):

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + ri_select + '/form-answers/' + list_dut[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            for i in range(len(json["questions"])):
                if "TIPO" in str(json["questions"][i]["title"]).upper() and "DUT" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and not "Refer" in str(json["questions"][i]["answer"]):

                        # Códigos DIEL para Máquinas:
                        for i in range(len(json["questions"])):
                            if "CÓDIGO" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                                if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                    tag_dut_list.append(json["questions"][i]["answer"])

                        # Tipos de DUT's:
                        for i in range(len(json["questions"])):
                            if "TIPO" in str(json["questions"][i]["title"]).upper() and "DUT" in str(json["questions"][i]["title"]).upper():
                                if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                    tipo_dut_list.append(json["questions"][i]["answer"])

                        # ID's dos DUT's:
                        for i in range(len(json["questions"])):
                            if "ID" in str(json["questions"][i]["title"]) and "DUT" in str(json["questions"][i]["title"]):
                                if not "FOTO" in str(json["questions"][i]["title"]).upper():
                                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                        id_dut_list.append(json["questions"][i]["answer"])

                        # Fotos dos ID's dos DUT's:
                        for i in range(len(json["questions"])):
                            if "ID" in str(json["questions"][i]["title"]) and "DUT" in str(json["questions"][i]["title"]) and "FOTO" in str(json["questions"][i]["title"]).upper():
                                if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                    foto_id_dut_list.append(json["questions"][i]["answer"])

                        # Fotos das Evaporadoras com DUT:
                        for i in range(len(json["questions"])):
                            if "MÁQUINA" in str(json["questions"][i]["title"]).upper() and "AMPLO" in str(json["questions"][i]["title"]).upper() and "FOTO" in str(json["questions"][i]["title"]).upper():
                                if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                    foto_dut_list.append(json["questions"][i]["answer"])

                        # Ambientes dos DUT's:
                        for i in range(len(json["questions"])):
                            if "AMBIENTE" in str(json["questions"][i]["title"]).upper() and "DUT" in str(json["questions"][i]["title"]):
                                if not "FOTO" in str(json["questions"][i]["title"]).upper():
                                    amb_dut_list.append(json["questions"][i]["answer"])

                        # Comentários Gerais das Evaporadoras:
                        for i in range(len(json["questions"])):
                            if "COMENT" in str(json["questions"][i]["title"]).upper() and "INST" in str(json["questions"][i]["title"]).upper():
                                coment_dut_list.append(json["questions"][i]["answer"])

                        # DAT(s) inserido(s) na instalação:
                        for i in range(len(json["questions"])):
                            if "CÓDIGO" in str(json["questions"][i]["title"]).upper() and "DAT" in str(json["questions"][i]["title"]).upper():
                                if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                    dat_dut_list.append(json["questions"][i]["answer"])

        '''print(tag_dut_list)
        print(tipo_dut_list)
        print(id_dut_list)
        print(foto_id_dut_list)
        print(foto_dut_list)
        print(amb_dut_list)
        print(coment_dut_list)'''

# Dispositivos REF's:

    for i in range(len(list_dut)):

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + ri_select + '/form-answers/' + list_dut[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            for i in range(len(json["questions"])):
                if "TIPO" in str(json["questions"][i]["title"]).upper() and "DUT" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and "Refer" in str(json["questions"][i]["answer"]):

                        # Códigos DIEL para Máquinas:
                        for i in range(len(json["questions"])):
                            if "CÓDIGO" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                                if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                    tag_ref_list.append(json["questions"][i]["answer"])

                        # Tipos de DUT's:
                        for i in range(len(json["questions"])):
                            if "TIPO" in str(json["questions"][i]["title"]).upper() and "DUT" in str(json["questions"][i]["title"]).upper():
                                if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                    tipo_ref_list.append(json["questions"][i]["answer"])

                        # ID's dos DUT's:
                        for i in range(len(json["questions"])):
                            if "ID" in str(json["questions"][i]["title"]) and "DUT" in str(json["questions"][i]["title"]):
                                if not "FOTO" in str(json["questions"][i]["title"]).upper():
                                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                        id_ref_list.append(json["questions"][i]["answer"])

                        # Fotos dos ID's dos DUT's:
                        for i in range(len(json["questions"])):
                            if "ID" in str(json["questions"][i]["title"]) and "DUT" in str(json["questions"][i]["title"]) and "FOTO" in str(json["questions"][i]["title"]).upper():
                                if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                    foto_id_ref_list.append(json["questions"][i]["answer"])

                        # Fotos dos Ambientes dos DUT's:
                        for i in range(len(json["questions"])):
                            if "MONIT" in str(json["questions"][i]["title"]).upper() and "AMBIENTE" in str(json["questions"][i]["title"]).upper() and "FOTO" in str(json["questions"][i]["title"]).upper():
                                if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                    foto_ref_list.append(json["questions"][i]["answer"])

                        # Ambientes dos DUT's:
                        for i in range(len(json["questions"])):
                            if "AMBIENTE" in str(json["questions"][i]["title"]).upper() and "DUT" in str(json["questions"][i]["title"]):
                                if not "FOTO" in str(json["questions"][i]["title"]).upper():
                                    amb_ref_list.append(json["questions"][i]["answer"])

        for i in range(len(list_croq)):

            url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + ri_select + '/form-answers/' + list_croq[i]
            headers = {"Content-Type": "application/json;charset=UTF-8",
                       "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
            req = requests.get(url, headers=headers)
            json = req.json()

            if not "statusCode" in str(json):
                for i in range(len(json["questions"])):
                    if "ANDAR" in str(json["questions"][i]["title"]).upper() and "CROQUI" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            and_croqui_list.append(json["questions"][i]["answer"])

                for i in range(len(json["questions"])):
                    if "CROQUI" in str(json["questions"][i]["title"]).upper() and "FOTO" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            if len(foto_croqui_list) < 5:
                                foto_croqui_list.append(json["questions"][i]["answer"])

            ''' print(andar_croqui_list)
            print(foto_croqui_list) '''

    for i in range(len(list_dme)):

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + ri_select + '/form-answers/' + list_dme[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            for i in range(len(json["questions"])):
                if "DISPOSIT" in str(json["questions"][i]["title"]).upper() and "A SER" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                        tipo_dme_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "ID" in str(json["questions"][i]["title"]) and "DME" in str(json["questions"][i]["title"]):
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and not "FOTO" in str(json["questions"][i]["title"]).upper():
                        id_dme_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "ID" in str(json["questions"][i]["title"]) and "DME" in str(json["questions"][i]["title"]):
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and "FOTO" in str(json["questions"][i]["title"]).upper():
                        foto_id_dme_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "TIPO" in str(json["questions"][i]["title"]).upper() and "ELÉTRICA" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                        inst_dme_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "DISPOSITIVO" in str(json["questions"][i]["title"]).upper() and "AMPL" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and "FOTO" in str(json["questions"][i]["title"]).upper():
                        foto_dme_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "CORRENTE" in str(json["questions"][i]["title"]).upper() and "TC" in str(json["questions"][i]["title"]):
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and not "FOTO" in str(json["questions"][i]["title"]).upper():
                        ampere_dme_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "SIMPLES" in str(json["questions"][i]["title"]).upper() and "TC" in str(json["questions"][i]["title"]):
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and not "FOTO" in str(json["questions"][i]["title"]).upper():
                        tc_dme_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "QUADRO" in str(json["questions"][i]["title"]).upper() and "TIPO" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and not "FOTO" in str(json["questions"][i]["title"]).upper():
                        quadro_dme_list.append(json["questions"][i]["answer"])

        '''print(tipo_dme_list)
        print(id_dme_list)
        print(quadro_dme_list)
        print(foto_id_dme_list)
        print(inst_dme_list)
        print(foto_dme_list)
        print(ampere_dme_list)
        print(tc_dme_list)'''

    for i in range(len(list_dma)):

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + ri_select + '/form-answers/' + list_dma[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            for i in range(len(json["questions"])):
                if "TIPO" in str(json["questions"][i]["title"]).upper() and "INSTAL" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and not "FOTO" in str(json["questions"][i]["title"]).upper():
                        tipo_dma_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "SOLUÇÃO" in str(json["questions"][i]["title"]).upper() and "INSTAL" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and not "FOTO" in str(json["questions"][i]["title"]).upper():
                        local_dma_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "ID" in str(json["questions"][i]["title"]) and "DMA" in str(json["questions"][i]["title"]):
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and not "FOTO" in str(json["questions"][i]["title"]).upper():
                        id_dma_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "FOTO" in str(json["questions"][i]["title"]).upper() and "ID" in str(json["questions"][i]["title"]):
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                        foto_id_dma_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "INSTAL" in str(json["questions"][i]["title"]).upper() and "AMPL" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and "FOTO" in str(json["questions"][i]["title"]).upper():
                        foto_dma_list.append(json["questions"][i]["answer"])

        '''print(tipo_dma_list)
        print(local_dma_list)
        print(id_dma_list)
        print(foto_id_dma_list)
        print(foto_dma_list)'''

    for i in range(len(list_dam)):

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + ri_select + '/form-answers/' + list_dam[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            for i in range(len(json["questions"])):
                if "CÓDIGO" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and not "FOTO" in str(json["questions"][i]["title"]).upper():
                        tag_dam_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "TIPO" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and "DAM" in str(json["questions"][i]["title"]):
                        tipo_dam_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "DAM" in str(json["questions"][i]["title"]) and "ID" in str(json["questions"][i]["title"]):
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and not "FOTO" in str(json["questions"][i]["title"]).upper():
                        id_dam_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "DAM" in str(json["questions"][i]["title"]) and "ID" in str(json["questions"][i]["title"]):
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and "FOTO" in str(json["questions"][i]["title"]).upper():
                        foto_id_dam_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "DAM" in str(json["questions"][i]["title"]) and "INSTAL" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and "FOTO" in str(json["questions"][i]["title"]).upper():
                        foto_dam_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "CÓDIGO" in str(json["questions"][i]["title"]).upper() and "DAT" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                        dat_dam_list.append(json["questions"][i]["answer"])

        '''print(tag_dam_list)
        print(tipo_dam_list)
        print(id_dam_list)
        print(foto_id_dam_list)
        print(foto_dam_list)'''

    for i in range(len(list_dat)):

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + ri_select + '/form-answers/' + list_dat[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            for i in range(len(json["questions"])):
                if "ID" in str(json["questions"][i]["title"]) and "DAT" in str(json["questions"][i]["title"]):
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and not "FOTO" in str(json["questions"][i]["title"]).upper():
                        id_dat_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "ID" in str(json["questions"][i]["title"]) and "DAT" in str(json["questions"][i]["title"]):
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and "FOTO" in str(json["questions"][i]["title"]).upper():
                        foto_id_dat_list.append(json["questions"][i]["answer"])
                    else:
                        foto_id_dat_list.append("None")

            for i in range(len(json["questions"])):
                if "TAG" in str(json["questions"][i]["title"]) and "ATIVO" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and not "FOTO" in str(json["questions"][i]["title"]).upper():
                        ativo_dat_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "AMPL" in str(json["questions"][i]["title"]).upper() and "DAT" in str(json["questions"][i]["title"]):
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and "FOTO" in str(json["questions"][i]["title"]).upper():
                        foto_dat_list.append(json["questions"][i]["answer"])

        '''print(tag_dat_list)
        print(foto_id_dat_list)
        print(ativo_dat_list)
        print(foto_dat_list)'''

    for i in range(len(list_rede)):

        url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + ri_select + '/form-answers/' + list_rede[i]
        headers = {"Content-Type": "application/json;charset=UTF-8",
                   "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
        req = requests.get(url, headers=headers)
        json = req.json()

        if not "statusCode" in str(json):
            for i in range(len(json["questions"])):
                if "ID" in str(json["questions"][i]["title"]) and "SIMCARD" in str(json["questions"][i]["title"]):
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and not "FOTO" in str(json["questions"][i]["title"]).upper():
                        id_rede_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "ID" in str(json["questions"][i]["title"]) and "SIMCARD" in str(json["questions"][i]["title"]):
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and "FOTO" in str(json["questions"][i]["title"]).upper():
                        foto_id_rede_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "TIPO" in str(json["questions"][i]["title"]).upper() and "DISPOSITIVO" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and not "FOTO" in str(json["questions"][i]["title"]).upper():
                        tipo_rede_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "AMBIENTE" in str(json["questions"][i]["title"]).upper() and "INSTALADO" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and not "FOTO" in str(json["questions"][i]["title"]).upper():
                        amb_rede_list.append(json["questions"][i]["answer"])

            for i in range(len(json["questions"])):
                if "DISPOSITIVO" in str(json["questions"][i]["title"]).upper() and "AMBIENTE" in str(json["questions"][i]["title"]).upper():
                    if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]) and "FOTO" in str(json["questions"][i]["title"]).upper():
                        foto_rede_list.append(json["questions"][i]["answer"])

        '''print(id_rede_list)
        print(foto_id_rede_list)
        print(tipo_rede_list)
        print(amb_rede_list)
        print(foto_rede_list)'''

        for i in range(len(list_evap_aux)):
            opc1 = 0
            opc2 = 0

            url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + ri_select + '/form-answers/' + list_evap_aux[i]
            headers = {"Content-Type": "application/json;charset=UTF-8",
                       "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
            req = requests.get(url, headers=headers)
            json = req.json()

            for i in range(len(json["questions"])):
                if "DAT" in str(json["questions"][i]["title"]):
                    if not "FOTO" in str(json["questions"][i]["title"]).upper():
                        if "CÓDIGO" in str(json["questions"][i]["title"]).upper():
                            if str(json["questions"][i]["answer"]) != "None" and "True" == str(
                                    json["questions"][i]["required"]):
                                evap_aux_list.append(json["questions"][i]["answer"])
                            else:
                                evap_aux_list.append("None")

            for i in range(len(json["questions"])):
                if opc2 == 0:
                    if "CÓDIGO" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            tag_evap_aux_list.append(json["questions"][i]["answer"])
                            opc1 = 1

            for i in range(len(json["questions"])):
                if opc1 == 0:
                    if "TAG" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            tag_evap_aux_list.append(json["questions"][i]["answer"])
                            opc2 = 1

        '''print(evap_aux_list)
        print(tag_evap_aux_list)'''

        for i in range(len(list_cond_aux)):
            opc1 = 1
            opc2 = 0

            url = 'https://amonamarth.fieldcontrol.com.br/maintenances/' + ri_select + '/form-answers/' + list_cond_aux[i]
            headers = {"Content-Type": "application/json;charset=UTF-8",
                       "X-Api-Key": "bWZTTmlQUVJ0cERYSHNwV3BFSnpxYmlZY2Z4UGhqaFg6MTc="}
            req = requests.get(url, headers=headers)
            json = req.json()

            for i in range(len(json["questions"])):
                if "DAT" in str(json["questions"][i]["title"]):
                    if not "FOTO" in str(json["questions"][i]["title"]).upper():
                        if "CÓDIGO" in str(json["questions"][i]["title"]).upper():
                            if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                                cond_aux_list.append(json["questions"][i]["answer"])
                            else:
                                cond_aux_list.append("None")

            for i in range(len(json["questions"])):
                if opc2 == 0:
                    if "CÓDIGO" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            tag_cond_aux_list.append(json["questions"][i]["answer"])
                            opc1 = 1

            for i in range(len(json["questions"])):
                if opc1 == 0:
                    if "TAG" in str(json["questions"][i]["title"]).upper() and "MÁQUINA" in str(json["questions"][i]["title"]).upper():
                        if str(json["questions"][i]["answer"]) != "None" and "True" == str(json["questions"][i]["required"]):
                            tag_cond_aux_list.append(json["questions"][i]["answer"])
                            opc2 = 1

        '''print(cond_aux_list)
        print(tag_cond_aux_list)'''

def planilha_uni():
    arquivo()

# Fuso Horário da Unidade - Noronha, Demais UFs, Amazonas e Acre:
    fusos_list = ["America/Noronha", "America/Sao_Paulo", "America/Toronto", "America/Chicago"]

    if "FERNANDO" in city.upper():
        if "NORONHA" in city.upper():
            aba["G2"] = fusos_list[0]
    elif "AM" in uf.upper():
        aba["G2"] = fusos_list[2]
    elif "AC" in uf.upper():
        aba["G2"] = fusos_list[3]
    else:
        aba["G2"] = fusos_list[1]

    vt_ri_banco_do_brasil()

# Fotos do Croqui do Relatório de Instalação:

    if len(foto_croqui_list) == 1:
        if str(foto_croqui_list[0]).find("g&") != -1:
            e_comercial = str(foto_croqui_list[0]).find("g&", 1)
            aba["K2"] = str(foto_croqui_list[0])[0:e_comercial + 1]
        else:
            aba["K2"] = str(foto_croqui_list[0])

    elif len(foto_croqui_list) == 2:
        if str(foto_croqui_list[0]).find("g&") != -1:
            e_comercial = str(foto_croqui_list[0]).find("g&", 1)
            aba["K2"] = str(foto_croqui_list[0])[0:e_comercial + 1]
        else:
            aba["K2"] = str(foto_croqui_list[0])

        if str(foto_croqui_list[1]).find("g&") != -1:
            e_comercial = str(foto_croqui_list[1]).find("g&", 1)
            e_comercial2 = str(foto_croqui_list[1]).find("g&", e_comercial)
            aba["L2"] = str(foto_croqui_list[1])[e_comercial + 2: e_comercial2 + 1]
        else:
            aba["L2"] = foto_croqui_list[1]

    elif len(foto_croqui_list) == 3:
        if str(foto_croqui_list[0]).find("g&") != -1:
            e_comercial = str(foto_croqui_list[0]).find("g&", 1)
            aba["K2"] = str(foto_croqui_list[0])[0:e_comercial + 1]
        else:
            aba["K2"] = str(foto_croqui_list[0])

        if str(foto_croqui_list[1]).find("g&") != -1:
            e_comercial = str(foto_croqui_list[1]).find("g&", 1)
            e_comercial2 = str(foto_croqui_list[1]).find("g&", e_comercial)
            aba["L2"] = str(foto_croqui_list[1])[e_comercial + 2: e_comercial2 + 1]
        else:
            aba["L2"] = foto_croqui_list[1]

        if str(foto_croqui_list[2]).find("g&") != -1:
            e_comercial = str(foto_croqui_list[2]).find("g&", 1)
            e_comercial2 = str(foto_croqui_list[2]).find("g&", e_comercial)
            e_comercial3 = str(foto_croqui_list[2]).find("g&", e_comercial2)
            aba["M2"] = str(foto_croqui_list[2])[e_comercial2 + 2: e_comercial3 + 1]
        else:
            aba["M2"] = foto_croqui_list[2]

    elif len(foto_croqui_list) == 4:
        if str(foto_croqui_list[0]).find("g&") != -1:
            e_comercial = str(foto_croqui_list[0]).find("g&", 1)
            aba["K2"] = str(foto_croqui_list[0])[0:e_comercial + 1]
        else:
            aba["K2"] = str(foto_croqui_list[0])

        if str(foto_croqui_list[1]).find("g&") != -1:
            e_comercial = str(foto_croqui_list[1]).find("g&", 1)
            e_comercial2 = str(foto_croqui_list[1]).find("g&", e_comercial)
            aba["L2"] = str(foto_croqui_list[1])[e_comercial + 2: e_comercial2 + 1]
        else:
            aba["L2"] = foto_croqui_list[1]

        if str(foto_croqui_list[2]).find("g&") != -1:
            e_comercial = str(foto_croqui_list[2]).find("g&", 1)
            e_comercial2 = str(foto_croqui_list[2]).find("g&", e_comercial)
            e_comercial3 = str(foto_croqui_list[2]).find("g&", e_comercial2)
            aba["M2"] = str(foto_croqui_list[2])[e_comercial2 + 2: e_comercial3 + 1]
        else:
            aba["M2"] = foto_croqui_list[2]

        if str(foto_croqui_list[3]).find("g&") != -1:
            e_comercial = str(foto_croqui_list[3]).find("g&", 1)
            e_comercial2 = str(foto_croqui_list[2]).find("g&", e_comercial)
            e_comercial3 = str(foto_croqui_list[2]).find("g&", e_comercial2)
            e_comercial4 = str(foto_croqui_list[2]).find("g&", e_comercial3)
            aba["N2"] = str(foto_croqui_list[3])[e_comercial3 + 2: e_comercial4 + 1]
        else:
            aba["N2"] = foto_croqui_list[3]

    elif len(foto_croqui_list) == 5:
        if str(foto_croqui_list[0]).find("g&") != -1:
            e_comercial = str(foto_croqui_list[0]).find("g&", 1)
            aba["K2"] = str(foto_croqui_list[0])[0:e_comercial + 1]
        else:
            aba["K2"] = str(foto_croqui_list[0])

        if str(foto_croqui_list[1]).find("g&") != -1:
            e_comercial = str(foto_croqui_list[1]).find("g&", 1)
            e_comercial2 = str(foto_croqui_list[1]).find("g&", e_comercial)
            aba["L2"] = str(foto_croqui_list[1])[e_comercial + 2: e_comercial2 + 1]
        else:
            aba["L2"] = foto_croqui_list[1]

        if str(foto_croqui_list[2]).find("g&") != -1:
            e_comercial = str(foto_croqui_list[2]).find("g&", 1)
            e_comercial2 = str(foto_croqui_list[2]).find("g&", e_comercial)
            e_comercial3 = str(foto_croqui_list[2]).find("g&", e_comercial2)
            aba["M2"] = str(foto_croqui_list[2])[e_comercial2 + 2: e_comercial3 + 1]
        else:
            aba["M2"] = foto_croqui_list[2]

        if str(foto_croqui_list[3]).find("g&") != -1:
            e_comercial = str(foto_croqui_list[3]).find("g&", 1)
            e_comercial2 = str(foto_croqui_list[2]).find("g&", e_comercial)
            e_comercial3 = str(foto_croqui_list[2]).find("g&", e_comercial2)
            e_comercial4 = str(foto_croqui_list[2]).find("g&", e_comercial3)
            aba["N2"] = str(foto_croqui_list[3])[e_comercial3 + 2: e_comercial4 + 1]
        else:
            aba["N2"] = foto_croqui_list[3]

        if str(foto_croqui_list[4]).find("g&") != -1:
            e_comercial = str(foto_croqui_list[4]).find("g&", 1)
            e_comercial2 = str(foto_croqui_list[2]).find("g&", e_comercial)
            e_comercial3 = str(foto_croqui_list[2]).find("g&", e_comercial2)
            e_comercial4 = str(foto_croqui_list[2]).find("g&", e_comercial3)
            e_comercial5 = str(foto_croqui_list[2]).find("g&", e_comercial4)
            aba["O2"] = str(foto_croqui_list[4])[e_comercial4 + 2: e_comercial5 + 1]
        else:
            aba["O2"] = foto_croqui_list[4]

# Fotos do Croqui da Visita Técnica:
    if len(foto_croqui_list) == 0:
        if len(croqui_list) == 1:
            if str(croqui_list[0]).find("g&") != -1:
                e_comercial = str(croqui_list[0]).find("g&", 1)
                aba["K2"] = str(croqui_list[0])[0:e_comercial + 1]
            else:
                aba["K2"] = str(croqui_list[0])

        elif len(croqui_list) == 2:
            if str(croqui_list[0]).find("g&") != -1:
                e_comercial = str(croqui_list[0]).find("g&", 1)
                aba["K2"] = str(croqui_list[0])[0:e_comercial + 1]
            else:
                aba["K2"] = str(croqui_list[0])

            if str(croqui_list[1]).find("g&") != -1:
                e_comercial = str(croqui_list[1]).find("g&", 1)
                aba["L2"] = str(croqui_list[1])[0:e_comercial + 1]
            else:
                aba["L2"] = str(croqui_list[1])

        elif len(croqui_list) == 3:
            if str(croqui_list[0]).find("g&") != -1:
                e_comercial = str(croqui_list[0]).find("g&", 1)
                aba["K2"] = str(croqui_list[0])[0:e_comercial + 1]
            else:
                aba["K2"] = str(croqui_list[0])

            if str(croqui_list[1]).find("g&") != -1:
                e_comercial = str(croqui_list[1]).find("g&", 1)
                aba["L2"] = str(croqui_list[1])[0:e_comercial + 1]
            else:
                aba["L2"] = str(croqui_list[1])

            if str(croqui_list[2]).find("g&") != -1:
                e_comercial = str(croqui_list[2]).find("g&", 1)
                aba["M2"] = str(croqui_list[2])[0:e_comercial + 1]
            else:
                aba["M2"] = str(croqui_list[2])

        elif len(croqui_list) == 4:
            if str(croqui_list[0]).find("g&") != -1:
                e_comercial = str(croqui_list[0]).find("g&", 1)
                aba["K2"] = str(croqui_list[0])[0:e_comercial + 1]
            else:
                aba["K2"] = str(croqui_list[0])

            if str(croqui_list[1]).find("g&") != -1:
                e_comercial = str(croqui_list[1]).find("g&", 1)
                aba["L2"] = str(croqui_list[1])[0:e_comercial + 1]
            else:
                aba["L2"] = str(croqui_list[1])

            if str(croqui_list[2]).find("g&") != -1:
                e_comercial = str(croqui_list[2]).find("g&", 1)
                aba["M2"] = str(croqui_list[2])[0:e_comercial + 1]
            else:
                aba["M2"] = str(croqui_list[2])

            if str(croqui_list[3]).find("g&") != -1:
                e_comercial = str(croqui_list[3]).find("g&", 1)
                aba["N2"] = str(croqui_list[3])[0:e_comercial + 1]
            else:
                aba["N2"] = str(croqui_list[3])

        elif len(croqui_list) == 5:
            if str(croqui_list[0]).find("g&") != -1:
                e_comercial = str(croqui_list[0]).find("g&", 1)
                aba["K2"] = str(croqui_list[0])[0:e_comercial + 1]
            else:
                aba["K2"] = str(croqui_list[0])

            if str(croqui_list[1]).find("g&") != -1:
                e_comercial = str(croqui_list[1]).find("g&", 1)
                aba["L2"] = str(croqui_list[1])[0:e_comercial + 1]
            else:
                aba["L2"] = str(croqui_list[1])

            if str(croqui_list[2]).find("g&") != -1:
                e_comercial = str(croqui_list[2]).find("g&", 1)
                aba["M2"] = str(croqui_list[2])[0:e_comercial + 1]
            else:
                aba["M2"] = str(croqui_list[2])

            if str(croqui_list[3]).find("g&") != -1:
                e_comercial = str(croqui_list[3]).find("g&", 1)
                aba["N2"] = str(croqui_list[3])[0:e_comercial + 1]
            else:
                aba["N2"] = str(croqui_list[3])

            if str(croqui_list[4]).find("g&") != -1:
                e_comercial = str(croqui_list[4]).find("g&", 1)
                aba["O2"] = str(croqui_list[4])[0:e_comercial + 1]
            else:
                aba["O2"] = str(croqui_list[4])

    global linha
    linha = 3

    for i in range(len(cond_list)):
        linha = 3 + i

        # Coluna de Tipo de Solução (A):
        aba["A"+str(linha)] = "Máquina"
        # Coluna do Nome da Unidade (B):
        aba["B"+str(linha)] = unidade_nome[(unidade_nome).find("-")+2:]
        # Coluna do Fabricante (U):
        aba["U" + str(linha)] = str(marca_cond_list[i])
        # Coluna de Fluido (V):
        aba["V" + str(linha)] = str(fluido_cond_list[i])
        # Coluna do Nome do Ativo (AS):
        aba["AW" + str(linha)] = str(cond_list[i]) + " - " + str(local_cond_list[i])
        # Coluna da Função (AT):
        aba["AX" + str(linha)] = "Condensadora"
        if i <= (len(model_cond_list) - 1) != 0:
            # Coluna do Modelo (AU):
            aba["AY" + str(linha)] = str(model_cond_list[i])
        # Coluna da Capacidade Frigorífica (AV):
        if float(cp_cond_list[i]) >= 1000:
            cap_frig = float(cp_cond_list[i])
            aba["AZ" + str(linha)] = cap_frig / 12000
            # Coluna da Potência Nominal [kW] (AX):
            aba["BB" + str(linha)] = (cap_frig / 12000) + ((cap_frig / 12000) * 0.1)
        else:
            cap_frig = float(cp_cond_list[i])
            aba["AZ" + str(linha)] = cap_frig
            # Coluna da Potência Nominal [kW] (AX):
            aba["BB" + str(linha)] = cap_frig + (cap_frig * 0.1)

        # Coluna do Nome da Máquina no DAP (Q):
        if len(evap_list) != 0:
            for item in range(len(evap_list)):
                # Coluna de Aplicação (S):
                aplicacoes_list = ["Ar-Condicionado", "Câmara Fria", "Trocador de Calor"]
                if str(tipo_evap_list[item])[0:5] == "Split":
                    aba["S" + str(linha)] = aplicacoes_list[0]
                elif str(tipo_evap_list[item])[0:6] == "Câmara":
                    aba["S" + str(linha)] = aplicacoes_list[1]
                else:
                    aba["S" + str(linha)] = aplicacoes_list[2]

                # Coluna de Nome da Máquina (Q):
                if str(correl_cond_list[i]) in str(evap_list[item]):
                    if str(un_cp_cond_list[i]).upper() == "BTU/H":
                        conversao = float(cp_cond_list[i])/12000
                        nome_maq = str(tipo_evap_list[item]) + " " + str(correl_cond_list[i]) + " (" + str(conversao) + " TR) - " + str(amb_evap_list[item])
                        aba["Q" + str(linha)] = nome_maq.replace(".", ",")
                        # Coluna do Tipo de Equipamento (T):
                        aba["T" + str(linha)] = str(tipo_evap_list[item])

            # Coluna de Foto 2 Máquina (X):
            if str(foto_cond_list[i]).find("g&") != -1:
                e_comercial = str(foto_cond_list[i]).find("g&", 1)
                aba["X" + str(linha)] = str(foto_cond_list[i])[0: e_comercial + 1]
                aba["BF" + str(linha)] = str(foto_cond_list[i])[0: e_comercial + 1]
            else:
                aba["X" + str(linha)] = str(foto_cond_list[i])
                aba["BF" + str(linha)] = str(foto_cond_list[i])

        for tag in range(len(tag_dac_list)):
            if str(cond_list[i]) == str(tag_dac_list[tag]):
                if len(str(id_dac_list[tag])) == 9:
                    # Coluna de Dispositivo Diel associado ao ativo (BD):
                    aba["BK" + str(linha)] = "DAC" + str(id_dac_list[tag])
                    aba["BL" + str(linha)] = "S"

                # Coluna de Foto 1 Máquina (W):
                if str(foto_dac_list[tag]).find("g&") != -1:
                    e_comercial = str(foto_dac_list[tag]).find("g&", 1)
                    aba["W" + str(linha)] = str(foto_dac_list[tag])[0: e_comercial + 1]
                    # Coluna Foto 1 DAC (BJ):
                    aba["BQ" + str(linha)] = str(foto_dac_list[tag])[0: e_comercial + 1]
                else:
                    aba["W" + str(linha)] = str(foto_dac_list[tag])
                    # Coluna Foto 1 DAC (BJ):
                    aba["BQ" + str(linha)] = str(foto_dac_list[tag])

        if i <= (len(dat_cond_list)-1) != 0:
            if len(dat_cond_list[i]) == 7:
                # Coluna de ID Ativo (AR):
                aba["AV" + str(linha)] = "DAT00" + str(dat_cond_list[i])
            elif len(dat_cond_list[i]) == 9:
                # Coluna de ID Ativo (AR):
                aba["AV" + str(linha)] = "DAT" + str(dat_cond_list[i])

        if len(id_dat_list) != 0:
            for dat in range(len(id_dat_list)):
                if i <= (len(cond_list)-1) and dat <= (len(ativo_dat_list)-1):

                    if str(ativo_dat_list[dat]).replace(" ", "").upper() in str(cond_list[i]).upper():
                        if len(id_dat_list[dat]) == 9:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT" + str(id_dat_list[dat])
                        elif len(id_dat_list[dat]) == 7:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT00" + str(id_dat_list[dat])

                    elif str(ativo_dat_list[dat]).upper() in str(cond_list[i]).upper():
                        if len(id_dat_list[dat]) == 9:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT" + str(id_dat_list[dat])
                        elif len(id_dat_list[dat]) == 7:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT00" + str(id_dat_list[dat])

        if len(dat_dac_list) != 0:
            for dat in range(len(dat_dac_list)):
                if i <= (len(cond_list)-1) and dat <= (len(tag_dac_list)-1):

                    if str(tag_dac_list[dat]).replace(" ", "").upper() in str(cond_list[i]).upper():
                        if len(dat_dac_list[dat]) == 9:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT" + str(dat_dac_list[dat])
                        elif len(dat_dac_list[dat]) == 7:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT00" + str(dat_dac_list[dat])

                    elif str(tag_dac_list[dat]).upper() in str(cond_list[i]).upper():
                        if len(dat_dac_list[dat]) == 9:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT" + str(dat_dac_list[dat])
                        elif len(dat_dac_list[dat]) == 7:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT00" + str(dat_dac_list[dat])

        if len(cond_aux_list) != 0:
            for dat in range(len(cond_aux_list)):
                if i <= (len(cond_list)-1) and dat <= (len(cond_aux_list)-1):

                    if str(tag_cond_aux_list[dat]).replace(" ", "").upper() in str(cond_list[i]).upper():
                        if len(cond_aux_list[dat]) == 9:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT" + str(cond_aux_list[dat])
                        elif len(cond_aux_list[dat]) == 7:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT00" + str(cond_aux_list[dat])

                    elif str(tag_cond_aux_list[dat]).upper() in str(cond_list[i]).upper():
                        if len(cond_aux_list[dat]) == 9:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT" + str(cond_aux_list[dat])
                        elif len(cond_aux_list[dat]) == 7:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT00" + str(cond_aux_list[dat])

    for i in range(len(evap_list)):
        linha = linha + 1

        # Coluna de Tipo de Solução (A):
        aba["A" + str(linha)] = "Máquina"
        # Coluna do Nome da Unidade (B):
        aba["B" + str(linha)] = unidade_nome[(unidade_nome).find("-") + 2:]
        # Coluna do Nome da Máquina no DAP (Q):
        if len(evap_list) != 0:
            aplicacoes_list = ["Ar-Condicionado", "Câmara Fria", "Trocador de Calor"]
            if str(tipo_evap_list[i])[0:5] == "Split":
                aba["S" + str(linha)] = aplicacoes_list[0]
            elif str(tipo_evap_list[i])[0:6] == "Câmara":
                aba["S" + str(linha)] = aplicacoes_list[1]
            else:
                aba["S" + str(linha)] = aplicacoes_list[2]

            # Coluna do Tipo de Equipamento (T):
            aba["T" + str(linha)] = str(tipo_evap_list[i])
            # Coluna do Fabricante (U):
            aba["U" + str(linha)] = str(marca_evap_list[i])
            # Coluna de Fluido (V):
            aba["V" + str(linha)] = str(fluido_evap_list[i])
            # Coluna do Nome do Ativo (AS):
            aba["AW" + str(linha)] = str(evap_list[i]) + " - " + str(amb_evap_list[i])
            # Coluna da Função (AT):
            aba["AX" + str(linha)] = "Evaporadora"

            # Coluna do Modelo (AU):
            if i <= (len(model_evap_list) - 1):
                aba["BC" + str(linha)] = str(model_evap_list[i])
                aba["AY" + str(linha)] = str(model_evap_list[i])

            # Coluna da Capacidade Frigorífica (AV):
            if float(cp_evap_list[i]) >= 1000:
                cap_frig = float(cp_evap_list[i])
                aba["AZ" + str(linha)] = cap_frig / 12000
                # Coluna da Potência Nominal [kW] (AX):
                aba["BB" + str(linha)] = (cap_frig / 12000) + ((cap_frig / 12000) * 0.1)
            else:
                cap_frig = float(cp_evap_list[i])
                aba["AZ" + str(linha)] = cap_frig
                # Coluna da Potência Nominal [kW] (AX):
                aba["BB" + str(linha)] = cap_frig + (cap_frig * 0.1)

            # Coluna de Foto 1 Máquina (W):
            for tag in range(len(tag_dut_list)):
                if str(evap_list[i]) == str(tag_dut_list[tag]):
                    if str(foto_dut_list[tag]).find("g&") != -1:
                        e_comercial = str(foto_dut_list[tag]).find("g&", 1)
                        aba["W" + str(linha)] = str(foto_dut_list[tag])[0: e_comercial + 1]
                    else:
                        aba["W" + str(linha)] = str(foto_dut_list[tag])

            # Coluna de Foto 2 Máquina (X):
            if str(foto_evap_list[i]).find("g&") != -1:
                e_comercial = str(foto_evap_list[i]).find("g&", 1)
                aba["X" + str(linha)] = str(foto_evap_list[i])[0: e_comercial + 1]
                aba["BF" + str(linha)] = str(foto_evap_list[i])[0: e_comercial + 1]
            else:
                aba["X" + str(linha)] = str(foto_evap_list[i])
                aba["BF" + str(linha)] = str(foto_evap_list[i])

            for tag in range(len(tag_dut_list)):
                if str(evap_list[i]) == str(tag_dut_list[tag]):
                    if len(str(id_dut_list[tag])) == 9:
                        # Coluna de Dispositivo de Automação (AB):
                        aba["AB" + str(linha)] = "DUT" + id_dut_list[tag]
                        # Coluna do Nome do Ambiente (AP):
                        aba["AT" + str(linha)] = str(evap_list[i]) + " - " + str(amb_evap_list[i])

                        # Coluna de Posicionamento (INS/AMB/DUO) (AC):
                        if str(tipo_dut_list[tag]).upper() == "DUO":
                            aba["AC" + str(linha)] = "DUO"
                            # Coluna de DUT DUO (POSIÇÃO SENSORES) (AD):
                            aba["AD" + str(linha)] = "Retorno, Insuflação"
                        elif "AUTO" in str(tipo_dut_list[tag]).upper():
                            aba["AC" + str(linha)] = "AMB"

                    # Coluna de Foto 1 Dispositivo de Automação (AE):
                    if str(foto_dut_list[tag]).find("g&") != -1:
                        e_comercial = str(foto_dut_list[tag]).find("g&", 1)
                        aba["AI" + str(linha)] = str(foto_dut_list[tag])[0: e_comercial + 1]
                    else:
                        aba["AI" + str(linha)] = str(foto_dut_list[tag])

            for maq in range(len(tag_dam_list)):
                if str(evap_list[i]) in str(tag_dam_list[maq]):
                    if len(str(id_dam_list[maq])) == 9:
                        # Coluna de Dispositivo de Automação (AB):
                        aba["AB" + str(linha)] = "DAM" + id_dam_list[maq]
                        # Coluna do Nome do Ambiente (AP):
                        aba["AT" + str(linha)] = str(evap_list[i]) + " - " + str(amb_evap_list[i])
                        # Local de Instalação do DAM:
                        aba["AE" + str(linha)] = str(amb_evap_list[i])
                        # Coluna de Posicionamento (INS/AMB/DUO) (AC):
                        if "DUO" in tipo_dam_list[maq].upper():
                            aba["AF" + str(linha)] = "DUO"
                            aba["AG" + str(linha)] = "Retorno"
                            aba["AH" + str(linha)] = "Insuflação"
                        else:
                            aba["AF" + str(linha)] = "Retorno"

            for item in range(len(cond_list)):
                # Coluna do Nome da Máquina (Q):
                if str(evap_list[i]) in str(correl_cond_list[item]):
                    if str(un_cp_cond_list[item]).upper() == "BTU/H":
                        conversao = float(cp_cond_list[item])/12000
                        nome_maq = str(tipo_evap_list[i]) + " " + str(correl_cond_list[item]) + " (" + str(conversao) + " TR) - " + str(amb_evap_list[i])
                        aba["Q" + str(linha)] = nome_maq.replace(".", ",")

        if i <= (len(dat_evap_list)-1) != 0:
            if len(dat_evap_list[i]) == 7:
                # Coluna de ID Ativo (AR):
                aba["AV" + str(linha)] = "DAT00" + str(dat_evap_list[i])
            elif len(dat_evap_list[i]) == 9:
                # Coluna de ID Ativo (AR):
                aba["AV" + str(linha)] = "DAT" + str(dat_evap_list[i])

        if len(id_dat_list) != 0:
            for dat in range(len(id_dat_list)):
                if i <= (len(evap_list)-1) and dat <= (len(ativo_dat_list)-1):

                    if str(ativo_dat_list[dat]).replace(" ", "").upper() in str(evap_list[i]).upper():
                        if len(id_dat_list[dat]) == 9:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT" + str(id_dat_list[dat])
                        elif len(id_dat_list[dat]) == 7:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT00" + str(id_dat_list[dat])

                    elif str(ativo_dat_list[dat]).upper() in str(evap_list[i]).upper():
                        if len(id_dat_list[dat]) == 9:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT" + str(id_dat_list[dat])
                        elif len(id_dat_list[dat]) == 7:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT00" + str(id_dat_list[dat])

        if len(dat_dut_list) != 0:
            for dat in range(len(dat_dut_list)):
                if i <= (len(evap_list)-1) and dat <= (len(tag_dut_list)-1):

                    if str(dat_dut_list[dat]).replace(" ", "").upper() in str(evap_list[i]).upper():
                        if len(dat_dut_list[dat]) == 9:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT" + str(dat_dut_list[dat])
                        elif len(dat_dut_list[dat]) == 7:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT00" + str(dat_dut_list[dat])

                    elif str(tag_dut_list[dat]).upper() in str(evap_list[i]).upper():
                        if len(dat_dut_list[dat]) == 9:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT" + str(dat_dut_list[dat])
                        elif len(dat_dut_list[dat]) == 7:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT00" + str(dat_dut_list[dat])

        if len(evap_aux_list) != 0:
            for dat in range(len(evap_aux_list)):
                if i <= (len(evap_list)-1) and dat <= (len(evap_aux_list)-1):

                    if str(tag_evap_aux_list[dat]).replace(" ", "").upper() in str(evap_list[i]).upper():
                        if len(evap_aux_list[dat]) == 9:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT" + str(evap_aux_list[dat])
                        elif len(evap_aux_list[dat]) == 7:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT00" + str(evap_aux_list[dat])

                    elif str(tag_evap_aux_list[dat]).upper() in str(evap_list[i]).upper():
                        if len(evap_aux_list[dat]) == 9:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT" + str(evap_aux_list[dat])
                        elif len(evap_aux_list[dat]) == 7:
                            # Coluna de ID Ativo (AR):
                            aba["AV" + str(linha)] = "DAT00" + str(evap_aux_list[dat])

    for i in range(len(id_dme_list)):
        linha = linha + 1

        # Coluna de Tipo de Solução (A):
        aba["A" + str(linha)] = "Energia"
        # Coluna do Nome da Unidade (B):
        aba["B" + str(linha)] = unidade_nome[(unidade_nome).find("-") + 2:]
        # Coluna de Aplicação (S):
        aba["S" + str(linha)] = "Medidor de Energia"
        # Coluna de Tipo de Quadro (BW):
        aba["BW" + str(linha)] = quadro_dme_list[i]
        if len(str(id_dme_list[i])) == 9:
            # Coluna de ID Med Energia (AJ):
            aba["BY" + str(linha)] = "DRI" + id_dme_list[i]
        elif len(str(id_dme_list[i])) == 7:
            # Coluna de ID Med Energia (AJ):
            aba["BY" + str(linha)] = "DRI00" + id_dme_list[i]

        # Coluna de Tipo de Dispositivo (CA):
        if "DME FULL" in str(tipo_dme_list).upper():
            aba["CA" + str(linha)] = "ET330"

        # Coluna de Capacidade TC (A):
        aba["CB" + str(linha)] = ampere_dme_list[i]
        # Coluna de Tipo de Instalação Elétrica (CC):
        aba["CC" + str(linha)] = inst_dme_list[i]
        # Coluna de Intervalo de Envio (s) (CD):
        aba["CD" + str(linha)] = 30

        # Coluna de Foto 1 DRI (CL):
        if str(foto_dme_list[i]).find("g&") != -1:
            e_comercial = str(foto_dme_list[i]).find("g&", 1)
            e_comercial2 = str(foto_dme_list[i]).find("g&", e_comercial + 1)
            aba["CL" + str(linha)] = str(foto_dme_list[i])[e_comercial + 2: e_comercial2 + 1]
        else:
            aba["CL" + str(linha)] = str(foto_dme_list[i])

    for i in range(len(id_ref_list)):
        linha = linha + 1

        # Coluna de Tipo de Solução (A):
        aba["A" + str(linha)] = "Ambiente"
        # Coluna do Nome da Unidade (B):
        aba["B" + str(linha)] = unidade_nome[(unidade_nome).find("-") + 2:]
        # Coluna de DUT de Referência (AJ):
        aba["AN" + str(linha)] = "DUT" + id_ref_list[i]
        # Coluna de Foto 1 DUT Referência (AK):
        if str(foto_ref_list[i]).find("g&") != -1:
            e_comercial = str(foto_ref_list[i]).find("g&", 1)
            aba["AO" + str(linha)] = foto_ref_list[i][0:e_comercial + 1]
        else:
            aba["AO" + str(linha)] = foto_ref_list[i]
        # Coluna de Nome do Ambiente (AP):
        if "VENDAS" in str(amb_ref_list[i]).upper():
            aba["AT" + str(linha)] = "Salão de Vendas"
        else:
            if " D" in str(amb_ref_list[i]).upper():
                aba["AT" + str(linha)] = str(amb_ref_list[i]).title().replace(" D", " d")
            else:
                aba["AT" + str(linha)] = str(amb_ref_list[i]).title()

    # Salvar arquivo Excel
    unificada.save(f"{unidade_nome} - Planilha unificada.xlsx")

planilha_uni()
