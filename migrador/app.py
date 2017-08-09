# -*- coding: utf-8 -*-
#------------------------------------------------------------------------------
# Author:      Elder Sanitá Trevisan
# Este sistema foi desenvolvido com a finalidade de migrar os dados de cadastro
# de condomínios e também as inadimplências desses condomínios do sistem SCC
# para o sistema da Superlógica
# Copyright(©) 2017 Elder Sanitá Trevisan
# Licence:     GPL
#------------------------------------------------------------------------------
from bottle import redirect, route, static_file, error, request, run,\
	jinja2_template as template
import html, json, os, csv, re, time
from datetime import datetime
from pyexcel_xls import get_data
from pycpfcnpj import cpfcnpj as doc

@route('/')
@route('/index')
def index():
	return template('index.html', title='Index')

@route('/cad_condominio', method=['GET', 'POST'])
def cad_condominio():
	if request.method == "POST":
		nome_condominio = html.escape(request.forms.get('nome_condominio'))
		arquivo_xls = request.POST['arquivo_xls']
		
		data = get_data(arquivo_xls.file)
		data_json = json.loads(json.dumps(data, ensure_ascii=False))
		
		with open("arquivos/CADASTRO/"+nome_condominio+"_CADASTRO.csv", 'w', newline='') as csv_file:
			field_names = [
							'unidade',\
							'bloco',\
							'fração',\
							'area',\
							'abatimento',\
							'proprietário_nome',\
							'proprietário_telefone',\
							'proprietário_celular',\
							'proprietário_forma_de_entrega',\
							'proprietário_cpf/cnpj',\
							'proprietário_rg',\
							'proprietário_email',\
							'proprietário_endereço',\
							'proprietário_complemento',\
							'proprietário_cep',\
							'proprietário_cidade',\
							'proprietário_bairro',\
							'proprietário_estado',\
							'inquilino_nome',\
							'inquilino_telefone',\
							'inquilino_celular',\
							'inquilino_forma_de_entrega',\
							'inquilino_cpf/cnpj',\
							'inquilino_rg',\
							'inquilino_email'
							]
			writer = csv.DictWriter(csv_file, fieldnames=field_names, delimiter=';')
			writer.writeheader()
			data_json['R90001'].pop(0)
			for i in data_json['R90001']:
				i.insert(len(i),None)
				writer.writerow({
								'unidade': i[3],\
								'bloco': str(i[2]),\
								'fração': i[4],\
								'proprietário_nome': i[5],\
								'proprietário_telefone': testa_telefone(i[11], i[12]),\
								'proprietário_celular':  i[13],\
								'proprietário_cpf/cnpj': testa_cpfcnpj(i[29]),\
								'proprietário_rg': re.sub(r'^RG:\s|RG:' ,"", i[27]),\
								'proprietário_email': i[14].lower(),\
								'proprietário_endereço': i[6],\
								'proprietário_cep': i[8],\
								'proprietário_cidade': i[9],\
								'proprietário_bairro': i[7],\
								'proprietário_estado': i[10],\
								'inquilino_nome': i[15],\
								'inquilino_telefone': i[21],\
								'inquilino_celular': i[23],\
								'inquilino_cpf/cnpj': testa_cpfcnpj(i[30]),\
								'inquilino_rg': re.sub(r'^RG:\s|RG:' ,"", i[28]),\
								'inquilino_email': i[25].lower()
								})
		
		redirect('/cad_condominio')
	else:
		return template('cadastro.html', title='Cadastro do condomínio')

def testa_cpfcnpj(documento):
	documento = re.sub(r'\D', "",str(documento))
	if documento != None:		
		if doc.validate(documento) == True:
			return documento
		else:
			return ""
	else:
		return ""

def testa_telefone(foneC, foneR):
	if len(foneC) == 0:
		return foneR
	elif len(foneR) == 0:
		return foneC
	else:
		return foneC+"-"+foneR

@route('/cad_inadimplente', method=['GET', 'POST'])
def cad_inadimplente():
	if request.method == "POST":
		nome_condominio = html.escape(request.forms.get('nome_condominio'))
		arquivo_xls = request.POST['arquivo_xls']
		subseq_check = request.forms.get('subsequente')
		
		data = get_data(arquivo_xls.file)
		data_json = json.loads(json.dumps(data, ensure_ascii=False))
		
		with open("arquivos/INADIMPLÊNCIA/"+nome_condominio+"_INADIMPLENTES.csv", 'w', newline='') as csv_file:
			field_names = [
							'unidade',\
							'bloco',\
							'vencimento',\
							'conta_bancária',\
							'nosso_numero',\
							'conta_categoria',\
							'complemento',\
							'valor',\
							'data_de_competência',\
							'taxa_de_juros_(%)',\
							'taxa_de_multa_(%)',\
							'taxa_de_desconto_(%)',\
							'cobrança_extraordinária',\
							'data_crédito',\
							'data_liquidação',\
							'valor_pago'
							]
			writer = csv.DictWriter(csv_file, fieldnames=field_names, delimiter=';')
			writer.writeheader()
			data_json['DEV101C'].pop(0)
			for i in data_json['DEV101C']:
				#i.insert(len(i),None)
				venc = time.strptime(str(i[3]), '%d-%b-%y')
				comp = testa_data_competencia(i[3], i[14], subseq_check)
				writer.writerow({
								'unidade': i[1],\
								'vencimento': str(venc[2])+"/"+str(venc[1])+"/"+str(venc[0]),\
								'nosso_numero': str(i[13]),\
								'conta_categoria': conta_categoria(i[5]),\
								'complemento': i[2],\
								'valor': i[8],\
								'data_de_competência': str(comp),\
								'taxa_de_juros_(%)': i[38],\
								'taxa_de_multa_(%)': i[37]
								})				
		redirect('/cad_inadimplente')
	else:
		return template('inadimplente.html', title='Cadastro de inadimplentes')

def conta_categoria(cc):
	contas = {
				"MENSAL": "1.1",
				"(AVULSO)": "1.21",
				"AVCB": "1.25",
				"FDO OBRAS": "1.18",
				"MULTA": "1.8",
				"TX EXTRA": "1.19",
				"GARAGEM": "1.23",
				"ALUGUEL": "1.28",
				"PINTURA": "1.30",
				"IPTU": "1.29",
				"ACORDO": "1.20"
			 }
	for chave,valor in contas.items():
		if cc == chave:
			return valor

def testa_data_competencia(venc, comp, subseq_check):
	venc = time.strptime(str(venc), '%d-%b-%y')
	if subseq_check == "checked":
		if comp == '':
			return str(venc[2])+"/"+str(venc[1])+"/"+str(venc[0])
		else:
			comp = time.strptime(str(comp), '%m%Y')
			return str(venc[2])+"/"+str(comp[1])+"/"+str(comp[0])
	else:
		if comp == '':
			if venc[1] == 1:
				return str(venc[2])+"/"+str(12)+"/"+str(venc[0]-1)
			else:
				return str(venc[2])+"/"+str(venc[1]-1)+"/"+str(venc[0])
		else:
			comp = time.strptime(str(comp), '%m%Y')
			return str(venc[2])+"/"+str(comp[1])+"/"+str(comp[0])

@route('/static/<filename:path>')
def server_static(filename):
	return static_file(filename, root='static/')

@error(404)
def error404(error):
	return "Ops, página não encontrada!"

@error(500)
def error500(error):
	return template('error.html', title="OPS, HOUVE UM ERRO...")

if __name__ == "__main__":
	run(reloader=True, debug=True, host='0.0.0.0', port=9090)