import pandas as pd
import fdb

dst_path = r'MTK:C:/Microsys/MsysIndustrial/Dados/MSYSDADOS.FDB'
excel_path1 = r'C:/Users/Gabriel/Desktop/Dados.xlsx'



TABLE_NAME1 = 'PEDIDOS_VENDAS_ITENS'
TABLE_NAME2 = 'PRODUTOS'
TABLE_NAME3 = 'PEDIDOS_VENDAS'
########################################################################
SELECT1 = 'select PVI_NUMERO, PVI_PRO_CODIGO, PVI_ITEM, PVI_QUANTIDADE, PVI_UNITARIO from %s ' \
          % (TABLE_NAME1)
SELECT2 = 'select  PRO_RESUMO, PRO_CODIGO, PRO_NIVEL3 from %s' % (TABLE_NAME2)
SELECT3 = 'select PDV_NUMERO, PDV_DATA, PDV_REP_CODIGO, PDV_VALORPRODUTOS, PDV_PSI_CODIGO, PDV_TIPOPAGAMENTO' \
          ' from %s' % (TABLE_NAME3)
########################################################################
con = fdb.connect(dsn=dst_path, user='SYSDBA', password='masterkey', charset='UTF8')
cur = con.cursor()
########################################################################
cur.execute(SELECT1)
table_rows1 = cur.fetchall()
########################################################################
cur.execute(SELECT2)
table_rows2 = cur.fetchall()
########################################################################
cur.execute(SELECT3)
table_rows3 = cur.fetchall()
########################################################################
df1 = pd.DataFrame(table_rows1)
########################################################################
df2 = pd.DataFrame(table_rows2)
########################################################################
df3 = pd.DataFrame(table_rows3)

df3 = df3[df3[4] != "CC"]
df3 = df3[df3[5] != "G"]

for y in df3.loc[2]:
    df3[2] = df3[2].replace([1,2,3,4,5,6,7,12],["Leid","Castilho","Loja","Site","Samuel", "Chico", "Zefs",
                                                "Michele"])
########################################################################
df4 = df3.drop(columns=3)
df4 = df4.drop(columns=4)
df4 = df4.drop(columns=1)
df4 = pd.merge(df1,df4, how='left', on=[0])
df4 = pd.merge(df4,df2, how='left', on=[1])
for x in df4.loc[2]:
    df4[2] = df4[2].replace([1, 2, 3, 4, 5, 6, 7, 8, 9, 29,
    1532, 1174, 161, 28, 27, 26, 25, 24, 22, 21, 20, 19, 18, 17, 16, 15, 13, 12, 11, 10, 0, None],
    ["Chave", "Chave", "Chave", "Chave", "Chave", "Chave", "Chave", "Chave", "Chave", "Chave","Ferragem",
     "Ferragem", "Ferragem", "Ferragem", "Ferragem", "Ferragem", "Ferragem", "Ferragem", "Ferragem",
     "Ferragem", "Ferragem", "Ferragem", "Ferragem", "Ferragem", "Ferragem", "Ferragem", "Ferragem",
     "Ferragem", "Ferragem", "Ferragem", "Ferragem", "Ferragem"])

data = df4[3]*df4[4]
df4[8] = data

########################################################################
with pd.ExcelWriter(excel_path1) as writer:
     df4.to_excel(writer, index=False, sheet_name='PEDIDOS_VENDAS_ITENS')
     df2.to_excel(writer, index=False, sheet_name='PRODUTOS')
     df3.to_excel(writer, index=False, sheet_name='PEDIDOS_VENDAS')