import time
from boto import kinesis
import json
import datetime
from collections import namedtuple
import xlwt
'''
    Antes de iniciar, executar o seguinte comando para atribuirmos os valores de acesso 
    $ aws configure --profile=parsely
   
    AWS Access Key ID AKIAIMYC23Z4HZ4BVJVA
    AWS Secret Access Key  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
    Default region name: us-east-1
    Default output format [None]: json
'''

'''
   Variavel de controle para quantos registros com dados vamos pegar para analisar
'''
rodar = 10

'''
       Iniciando Arquivo excel
'''
wb = xlwt.Workbook()
ws = wb.add_sheet('kinesis')

i = 0

'''
      Criando o cabeçalho dos dados que vamos armazenar
'''
ws.write(i, 0, "data")
ws.write(i, 1, "enviado")
ws.write(i, 2, "aberto")
ws.write(i, 3, "clicou")

'''
      informando region que vamos conectar
'''
kinesis = kinesis.connect() .connect_to_region("us-east-1")

'''
      iniciando shard_id
'''
shard_id = 'shardId-000000000000'

'''
      Informando de onde vamos pegar os dados e como vamos iterar 
'''
shard_it = kinesis.get_shard_iterator("big-data-analytics-desafio", shard_id, "LATEST")["ShardIterator"]

result = {}
t = 0

'''
      iniciando while que vai parar de rodar apenas quando pegarmos a quantidade de registros validos informado no inicio
'''
while i == 0:
    '''
          atribuindo uma variavel com valores coletados e ja buscando o proximo registro
    '''
    out = kinesis.get_records(shard_it)
    shard_it = out["NextShardIterator"]

    '''
          Verifica se existe informações para processar 
    '''
    if len(out["Records"]) != 0:
        '''
              Controle de quantos registros validos pegamos
        '''
        t += 1

        '''
              Iteramos sobre o Record encontrado
        '''
        for resp in out["Records"]:

            '''
                 Pegando o conteudo do Data e convertendo em um dicionario
            '''
            respJson = json.loads(resp["Data"], object_hook=lambda d: namedtuple('X', d.keys())(*d.values()))

            '''
                 Convertendo data
            '''
            date = datetime.datetime.fromtimestamp(respJson.datetime / 1e3).date()

            '''
                 Zerando contador
            '''
            sent = 0
            opened = 0
            clicked = 0

            '''
                Identificando o tipo de evento que ocorreu
            '''
            if respJson.event_type == 'sent':
                sent = 1
            elif respJson.event_type == 'opened':
                opened = 1
            elif respJson.event_type == 'clicked':
                clicked = 1

            '''
                Procurando se ja temos a data armazenada e caso tenha somamos com o valor ja existente
            '''
            if date in result:
                sent += result[date][0]
                opened += result[date][1]
                clicked += result[date][2]

            '''
                Grava resultado utilizando da data como indice 
            '''
            result[date] = [sent, opened, clicked]

    '''
       Verifica se ja rodou a quantidade de vezes solicitada no inicio, por padrao essa quantidade esta em 10
    '''
    if t >= rodar:
        i = 1

    '''
       Esperamos 2 sec para pedir mais dados epenas para teste
    '''
    time.sleep(0.2)

'''
    Itera sobre o resultado coletado para gravar no arquivo
'''
for k, v in result.items():
    print("FOR.:", k)
    ws.write(i, 0, str(k))
    ws.write(i, 1, v[0])
    ws.write(i, 2, v[1])
    ws.write(i, 3, v[2])
    i += 1

'''
    Gravamos o arquivo no diretorio do projeto
'''
wb.save('ResultKinesis.xls')