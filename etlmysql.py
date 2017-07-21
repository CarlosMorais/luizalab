import pymysql
import xlwt

''' Teste utilizando Classe
 class Orders:
    def __init__(self,category,sub_category,brand,status,qtd_vendas,precoVendaBruto,precoVendaLiquido,total,precoEnvio,dataPedido,dataUltimaAtualizacao):
        self.category = category
        self.sub_category = sub_category
        self.brand = brand
        self.status = status
        self.qtd_vendas = qtd_vendas
        self.precoVendaBruto = precoVendaBruto
        self.precoVendaLiquido = precoVendaLiquido
        self.total = total
        self.precoEnvio = precoEnvio
        self.dataPedido = dataPedido
        self.dataUltimaAtualizacao = dataUltimaAtualizacao
    '''
try:

    '''
       Definição de variavel estilo que vamos utilizar na gravação do arquivo de modelo excel
    '''
    style1 = xlwt.easyxf(num_format_str='#,##0.00')

    '''
        Iniciando Arquivo excel
    '''
    wb = xlwt.Workbook()
    ws = wb.add_sheet('dados')

    i = 0
    '''
       Criando o cabeçalho dos dados que vamos armazenar
    '''
    ws.write(i, 0, "category")
    ws.write(i, 1, "sub_category")
    ws.write(i, 2, "brand")
    ws.write(i, 3, "status")
    ws.write(i, 4, "qtd_price")
    ws.write(i, 5, "selling_price_total")
    ws.write(i, 6, "net_sales_price")
    ws.write(i, 7, "total")
    ws.write(i, 8, "shipping_price")
    ws.write(i, 9, "order_date")
    ws.write(i, 10, "last_update")

    '''
        Iniciando conexão com banco de dados mysql 
    '''
    db = pymysql.connect(host='big-data-analytics-test.cd2vfjltihkr.us-east-1.rds.amazonaws.com', user='desafio',
                         passwd='1i450U', db='bigdata_desafio')

    cursor = db.cursor()

    '''
       Executando consulta inicial onde vamos listar as Marcas inicialmente 
    '''
    cursor.execute("Select brand From products group by brand")

    '''
       pegando todos os dados retornados da consulta e atribuindo a uma variavel
    '''
    data = cursor.fetchall()

    '''
        Iniciando loop dos registros de Marcas
    '''
    for row in data:
        print("MARCAS", row[0])
        '''
            Montando a consulta que vamos extrair por marcas, de forma que fique mais leve e seu resultado atenda as demandas solicitadas
        '''
        sql = """Select p.category 
                      , p.sub_category 
                      , p.brand  
                      , o.status 
                      , count(o.order_id)                      as qtd_vendas 
                      , round(sum(i.selling_price))            as selling_price_total
                      , round(sum(i.selling_price) - 
                        sum(o.shipping_price))                 as net_sales_price
                      , round(sum(o.total))                    as total
                      , round(sum(o.shipping_price))           as shipping_price
                      , DATE_FORMAT(o.order_date, '%d/%m/%Y')  as order_date
                      , DATE_FORMAT(o.last_update, '%d/%m/%Y') as last_update
                   From products p inner join orderitem i on p.product_id = i.product_id
                                   inner join orders    o on i.order_id = o.order_id 
                  Where p.brand = '""" + row[0] + """'
                  Group by p.category
                         , p.sub_category
                         , p.brand
                         , o.status
                         , DATE_FORMAT(o.order_date, '%d/%m/%Y') 
                         , DATE_FORMAT(o.last_update, '%d/%m/%Y')"""

        '''
           Executando consulta que vamos exportar
        '''
        cursor.execute(sql)

        '''
           pegando todos os dados retornados
        '''
        resp = cursor.fetchall()

        '''
            Iniciando loop dos registros de finais
        '''
        for result in resp:
            i += 1
            j = 0
            '''
                Iniciando loop dos registros de cada linha para armazenar no arquivo
            '''
            while j <= 10:

                '''
                    Caso registro for decimal, aplicar mascara de formatação
                '''
                if j >= 5 and j <= 8:
                    ws.write(i, j, result[j], style1)
                else:
                    ws.write(i, j, result[j])

                '''
                    Incrementa variavel para percorrer proxima coluna
                '''
                j += 1
    '''
        Fecha conexão com banco
    '''
    db.close()

    '''
        Salva o arquivo xls na pasta do projeto
    '''
    wb.save('resultado.xls')

except pymysql as err:
    '''
       Caso tenha algum erro de baanco de dados print na tela
    '''
    print("erro .:", err)

    '''
        #Fecha conexão com banco
    '''
    db.close()
