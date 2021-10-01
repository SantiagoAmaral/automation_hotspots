# import libraries

import pandas as pd
import numpy as np
import geopandas as gpd
import requests
import io
import plotly.express as px
import plotly.graph_objects as go
import datetime
import time
import matplotlib.pyplot as plt
from datetime import date
import six
import xlsxwriter

import win32com.client as win32


# Shapefiles
municipios = gpd.read_file('./shp/municipios_2019.shp')
UC = gpd.read_file('./shp/UCs.shp')

#Function to organize data acquisition date. Always 1 day before the current one.

def datestdtojd (stddate):
  fmt='%Y-%m-%d'
  sdtdate = datetime.datetime.strptime(stddate, fmt)
  sdtdate = sdtdate.timetuple()
  jdate = sdtdate.tm_yday
  return(jdate)

data_hj = datetime.date.today().strftime("%Y-%m-%d")
data_arquivo = str(date.today().year) + str(datestdtojd(data_hj))

#Nasa Authorization
my_headers = {'Authorization' : 'Bearer {YWxpc3Nvbi5hbWFyYWw6WVd4cGMzTnZiaTVoYldGeVlXeEFhVzVsYldFdVltRXVaMjkyTG1KeToxNjI3MzA4NjEzOjhiZDhkOGM3OTlmMWUxMzVmMGM5ZGY1NjkxN2JhYmNhNzc3OGIzNGY}'}

#Nasa_links
noaa_link = 'https://nrt3.modaps.eosdis.nasa.gov/api/v2/content/archives/FIRMS/noaa-20-viirs-c2/South_America/J1_VIIRS_C2_South_America_VJ114IMGTDL_NRT_' + data_arquivo + '.txt'
aqte_link = 'https://nrt3.modaps.eosdis.nasa.gov/api/v2/content/archives/FIRMS/modis-c6.1/South_America/MODIS_C6_1_South_America_MCD14DL_NRT_' + data_arquivo + '.txt'
npp_link = 'https://nrt3.modaps.eosdis.nasa.gov/api/v2/content/archives/FIRMS/suomi-npp-viirs-c2/South_America/SUOMI_VIIRS_C2_South_America_VNP14IMGTDL_NRT_' + data_arquivo + '.txt'

print('Realizando os requests...')

#requests and dataframe
 #--NOAA
noaa_rqt = requests.get(noaa_link, headers=my_headers).content
data_noaa = pd.read_csv(io.StringIO(noaa_rqt.decode('utf-8')))
 #--AQUA & TERRA
aqte_rqt = requests.get(aqte_link, headers=my_headers).content
data_aqte = pd.read_csv(io.StringIO(aqte_rqt.decode('utf-8')))
 #--NPP
npp_rqt = requests.get(npp_link, headers=my_headers).content
data_npp = pd.read_csv(io.StringIO(npp_rqt.decode('utf-8')))

print('Realizando operações com o Shapefile...')
# Intersection of municipalities and conservation units
 #--NOAA
data_noaa = gpd.GeoDataFrame(data_noaa, geometry=gpd.points_from_xy(data_noaa.longitude,data_noaa.latitude, crs='EPSG:4674'))
data_noaa = gpd.overlay(data_noaa, municipios, how='intersection')

data_noaa_UC = gpd.overlay(data_noaa,UC,how = 'intersection')
data_noaa_NonUC = gpd.overlay(data_noaa,UC,how = 'difference')

data_noaa = pd.concat([data_noaa_UC, data_noaa_NonUC])

dict_sat_noaa = {1:'NOAA'}
data_noaa['satellite'] = data_noaa['satellite'].map(dict_sat_noaa)
data_noaa = data_noaa.loc[:,['municipios', 'territorio', 'regiao_cli', 'latitude', 'longitude', 'acq_date','acq_time','daynight',
                             'satellite', 'frp', 'NOME_UC', 'GRUPO', 'Dominio']]
 #--AQUA & TERRA
data_aqte = gpd.GeoDataFrame(data_aqte, geometry=gpd.points_from_xy(data_aqte.longitude,data_aqte.latitude, crs='EPSG:4674'))
data_aqte = gpd.overlay(data_aqte, municipios, how='intersection')

data_aqte_UC = gpd.overlay(data_aqte, UC, how = 'intersection')
data_aqte_NonUC = gpd.overlay(data_aqte, UC, how = 'difference')

data_aqte = pd.concat([data_aqte_UC, data_aqte_NonUC])

dict_sat_aqte = {'A':'AQUA','T':'TERRA'}
data_aqte['satellite'] = data_aqte['satellite'].map(dict_sat_aqte)
data_aqte = data_aqte.loc[:,['municipios', 'territorio', 'regiao_cli', 'latitude', 'longitude', 'acq_date','acq_time','daynight',
                             'satellite', 'frp', 'NOME_UC', 'GRUPO', 'Dominio']]
 #--NPP-375
data_npp = gpd.GeoDataFrame(data_npp, geometry=gpd.points_from_xy(data_npp.longitude,data_npp.latitude, crs='EPSG:4674'))
data_npp = gpd.overlay(data_npp, municipios, how='intersection')

data_npp_UC = gpd.overlay(data_npp, UC, how = 'intersection')
data_npp_NonUC = gpd.overlay(data_npp, UC, how = 'difference')

data_npp = pd.concat([data_npp_UC, data_npp_NonUC])

dict_sat_npp = {'N':'NPP'}
data_npp['satellite'] = data_npp['satellite'].map(dict_sat_npp)
data_npp = data_npp.loc[:,['municipios', 'territorio', 'regiao_cli', 'latitude', 'longitude', 'acq_date','acq_time','daynight',
                           'satellite', 'frp', 'NOME_UC', 'GRUPO', 'Dominio']]

print('Preparando Dataframe...')
#Concat and filter the Dataframes
data_final = pd.concat([data_noaa,data_aqte,data_npp])
 #--Data for All satellites
data_final2 = data_final.loc[:,['acq_date','acq_time','satellite','municipios','latitude', 
                                'longitude','NOME_UC']].sort_values(by='municipios', ascending=True)
data_final2 = data_final2.rename(columns={'acq_date': 'DATA','acq_time': 'HORÁRIO','satellite': 'SATÉLITE', 'municipios': ' MUNICÍPIOS',
                                          'latitude': 'LATITUDE','longitude': 'LONGITUDE','NOME_UC': 'UNIDADE DE CONSERVAÇÃO' })
 #--Counts hotspots per city - AQUA
filtro_aqua = data_final[(data_final['satellite'] == 'AQUA') & (data_final['daynight'] == 'D')].copy()
filtro_aqua['quantidade'] = [1 for i in list(filtro_aqua['municipios'])]

tabela_aqua = filtro_aqua[['municipios','quantidade']].groupby(['municipios'],as_index=False).count().sort_values(by='quantidade', ascending=False)

#Add Sum of hotspots
linha_total = pd.DataFrame()
aqua_quantidade = tabela_aqua.quantidade.sum()
linha_total['Total'] = [aqua_quantidade]
linha_total = linha_total.T.reset_index()
linha_total.columns = ['municipios','quantidade']

tabela_aqua = pd.concat([tabela_aqua,linha_total])
tabela_aqua = tabela_aqua.rename(columns={'municipios': 'MUNICÍPIOS','quantidade': 'QUANTIDADE'})


#Date and time
date_corr = datetime.datetime.strptime(str(data_arquivo[2:4]) + str(data_arquivo[4:]), '%y%j').date()
date_corr = date_corr.strftime('%d-%m-%Y')

time_is_it = datetime.datetime.now().strftime('%H%M')


#Save the Datas
 #--All Satellites
excel_data_final = pd.ExcelWriter('Y:/METEOROLOGIA/Boletins Meteorológicos/Boletim Diário de Focos de Calor/Dados_diarios/Geral/FIRES' + '_-_' + date_corr + '_-_' + time_is_it + '_-_' +  "AQUA,TERRA,NOAA,NPP" + ".xlsx", engine='xlsxwriter')
data_final2.to_excel(excel_data_final, sheet_name='Focos de Calor Total', index=False)
excel_data_final.save()
 #--Number of hot spots per city - AQUA
excel_tabela_aqua = pd.ExcelWriter("Y:/METEOROLOGIA/Boletins Meteorológicos/Boletim Diário de Focos de Calor/Dados_diarios/Satelite_referencia/TOTAL" + '_-_' + "AQUA" + '_-_' + date_corr + '_-_' + time_is_it + ".xlsx", engine='xlsxwriter')
tabela_aqua.to_excel(excel_tabela_aqua, sheet_name='Focos de Calor AQUA', index=False)
excel_tabela_aqua.save()

print('Enviando e-mail...')
#Email

# criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')
# criar um email
email = outlook.CreateItem(0)
# configurar as informações do seu e-mail
email.To = "alisson.santiago123@gmail.com"
email.Subject = "Focos de Calor na Bahia"
email.HTMLBody = f"""
<p>Prezados, boa tarde!</p>

<p>Estamos enviando uma lista com o total de focos de calor registrados pelo <strong>satélite de referência AQUA no dia {date_corr}</strong>.<br />
Este satélite registrou <strong>{aqua_quantidade} focos de calor na Bahia </strong> no decorrer deste período (tabela por município em anexo).</p>

<p>Qualquer dúvida, estamos à disposição.</p>

<p>Atenciosamente,</p>

<table width="500" border="0" cellspacing="0" cellpadding="0" class="" style="font-family:&quot;Times New Roman&quot;; table-layout:fixed">
  <tbody>
    <tr class="" style="font-size:16px"><td height="112" rowspan="4" class="" style="font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif; width:172px; margin:0px"><img data-imagetype="AttachmentByCid" originalsrc="cid:8f1c1be1-1cf9-4797-a201-f273ecf2e628" data-custom="AAMkAGM3NzU5NzBjLWE2NDMtNDhkNS1iZTk4LTE3YWZlOWIxNmVmYgBGAAAAAAB%2B1RbOkSloS5OXKhzoIYysBwAtxFGzRJyTSpiMUcOYxGAeAAAAAAEJAAAtxFGzRJyTSpiMUcOYxGAeAAWj7nxPAAABEgAQANLBwqAqIHBNkv6EbITAvKc%3D" naturalheight="0" naturalwidth="0" src="https://attachments.office.net/owa/meteorologia%40inema.ba.gov.br/service.svc/s/GetAttachmentThumbnail?id=AAMkAGM3NzU5NzBjLWE2NDMtNDhkNS1iZTk4LTE3YWZlOWIxNmVmYgBGAAAAAAB%2B1RbOkSloS5OXKhzoIYysBwAtxFGzRJyTSpiMUcOYxGAeAAAAAAEJAAAtxFGzRJyTSpiMUcOYxGAeAAWj7nxPAAABEgAQANLBwqAqIHBNkv6EbITAvKc%3D&amp;thumbnailType=2&amp;token=eyJhbGciOiJSUzI1NiIsImtpZCI6IjMwODE3OUNFNUY0QjUyRTc4QjJEQjg5NjZCQUY0RUNDMzcyN0FFRUUiLCJ0eXAiOiJKV1QiLCJ4NXQiOiJNSUY1emw5TFV1ZUxMYmlXYTY5T3pEY25ydTQifQ.eyJvcmlnaW4iOiJodHRwczovL291dGxvb2sub2ZmaWNlLmNvbSIsInVjIjoiM2YzMTA4MjhhZjc0NDk2YThkMTk1NTFhMjViOGVhYzAiLCJ2ZXIiOiJFeGNoYW5nZS5DYWxsYmFjay5WMSIsImFwcGN0eHNlbmRlciI6Ik93YURvd25sb2FkQDRjZDgzNWY0LTU0NDAtNDA4Zi05M2EzLTk3NWZjMTdjMzg0YSIsImlzc3JpbmciOiJXVyIsImFwcGN0eCI6IntcIm1zZXhjaHByb3RcIjpcIm93YVwiLFwicHVpZFwiOlwiMTE1MzkwNjY2MDkxMzQ3NTg3OFwiLFwic2NvcGVcIjpcIk93YURvd25sb2FkXCIsXCJvaWRcIjpcIjk1MzU5ZDcxLTgyZGEtNDgxYS04NjVhLWFhOWQ4OTRlMjA4NlwiLFwicHJpbWFyeXNpZFwiOlwiUy0xLTUtMjEtMzUzODUxOTQ3My05MDExMzkwODQtMzM5MDE4ODgwMy03MTcxNDc1XCJ9IiwibmJmIjoxNjMzMDk2NTAwLCJleHAiOjE2MzMwOTcxMDAsImlzcyI6IjAwMDAwMDAyLTAwMDAtMGZmMS1jZTAwLTAwMDAwMDAwMDAwMEA0Y2Q4MzVmNC01NDQwLTQwOGYtOTNhMy05NzVmYzE3YzM4NGEiLCJhdWQiOiIwMDAwMDAwMi0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDAvYXR0YWNobWVudHMub2ZmaWNlLm5ldEA0Y2Q4MzVmNC01NDQwLTQwOGYtOTNhMy05NzVmYzE3YzM4NGEiLCJoYXBwIjoib3dhIn0.LmK5Y8lJR-WKjmOJLkJXyh-7UFjog14Iak_ZS1JRHBIz7N7xplbxbIuN1YiB_vcAbY_WvA9cx00_adzzbnLqbjwYoO6XFOUmO1-MTShbkcELYlcFm6kgS2KwHqf-Btgj3mFaB18uObMVN5hrrVYqQJnEK9tL52sCtz_-zhCH5b-yEfonus4VLW8ilkSr9j59eB0WLUuCPPvO4aTVA0wWuShl4mLRAh6hNmtmvyDM5jlIFU97iqZApLGn1lqwRqTEXRc8wkdQIR9ajF4TGQfOwmERbKSrdsSbwRdbMHZ31Ui0WVwQsiuJ5Q6XzDLyo3XJApZ5JlrAGhZ712SQhTC6bQ&amp;X-OWA-CANARY=RGrhWQ_m2k2t4bfGiJ3ghlDTFxDjhNkYXPzk6R0HZBXdIQK8iiAVtBU3vNf6yhI93UEsbKk0hYw.&amp;owa=outlook.office.com&amp;scriptVer=20210927003.07&amp;animation=true" class="x_x_EmojiInsert" data-outlook-trace="F:2|T:2" style="cursor: pointer; max-width: 100%; height: auto;" crossorigin="use-credentials"><br></td>
    <td height="112" rowspan="4" class="" style="font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif; width:23px; margin:0px"><br><img data-imagetype="External" src="https://ci6.googleusercontent.com/proxy/n1qHxeXlChWZikRMLC3qdBJC_mBdx_FymR1uscY3_f0j862ZQSWUmroPT09mf4aXvvqb-d0ZBCI6TF83bdo7xovGLrWN5hd4foi5AfBFx84neV2y2oMsyPpQ5J2jLBaVGtYNNFaIOw=s0-d-e1-ft#http://www.somarmeteorologia.com.br/assinatura/images/somar-assinatura-modelo_02.png" width="39" height="112" style="display:block"></td><td height="18" class="" style="font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif; width:322px; margin:0px"></td></tr><tr class="" style="font-size:16px"><td class="" style="font-family:Arial,sans-serif; vertical-align:text-top; margin:0px; line-height:17px"><font style="color:rgb(25,60,80)"><b>
    <br>Equipe de Meteorologia</b></font><br><span style="color:rgb(25,60,80); font-size:14px; font-family:Arial,sans-serif,serif,EmojiFont"><br>Diretoria de Recursos Hídricos e Monitoramento Ambiental<br>Coordenação de Estudos do Clima e Projetos Especiais&nbsp;- COCEP</span><br><font style="color:rgb(25,60,80); font-size:14px">Tel.: (71) 3118-4163 / 4162<br><br></font></td></tr><tr class="" style="font-size:16px">
    <td rowspan="1" class="" style="font-family:Arial,sans-serif; vertical-align:text-top; margin:0px; line-height:17px"><br></td></tr>
    </tbody>
</table>

"""

email.Attachments.Add("Y:/METEOROLOGIA/Boletins Meteorológicos/Boletim Diário de Focos de Calor/Dados_diarios/Satelite_referencia/TOTAL" + '_-_' + "AQUA" + '_-_' + date_corr + '_-_' + time_is_it + ".xlsx")
email.Send()
print("Email Enviado")