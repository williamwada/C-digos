import xlsxwriter
import logging
import pyodbc
import re
from xlsxwriter.utility import xl_rowcol_to_cell
#from datetime import datetime

class GerarExcel:

    def writeExcel(self, listAbaExcel_Anual, cursor,conexao, nome_arquivo, fonte, nota):

        try:
            db = 'BACEN'
            # Criacao do arquivo
            workbook = xlsxwriter.Workbook(str(nome_arquivo) + '.xlsx')
            workbook.formats[0].set_font_size(8)  # default cell format to size 8
            workbook.formats[0].set_font_name('Arial')  # default cell Arial
            # Set Main_Titulo
            titulo_format = workbook.add_format(
                {'font_name': 'Arial', 'bold': True, 'italic': False, 'size': 18, 'indent': 1})
            inf_format = workbook.add_format(
                {'font_name': 'Arial', 'bold': True, 'italic': False, 'size': 8, 'font_color': '#808080', 'indent': 2})
            alerta_format = workbook.add_format(
                {'font_name': 'Arial', 'bold': True, 'italic': False, 'size': 8, 'font_color': '#808080'})
            desc_b_format = workbook.add_format({'font_name': 'Arial', 'bold': True, 'italic': False, 'size': 8})
            desc_format = workbook.add_format({'font_name': 'Arial', 'bold': False, 'italic': False, 'size': 8})
            column_format = workbook.add_format(
                {'align': 'center', 'font_name': 'Arial', 'bold': True, 'size': 8, 'font_color': '#FFFFFF',
                 'bg_color': '#003366', 'border': 7})
            Per_col_format = workbook.add_format(
                {'align': 'center', 'font_name': 'Arial', 'bold': True, 'size': 8, 'font_color': '#FFFFFF',
                 'bg_color': '#003366', 'border': 0 },)
            row_format = workbook.add_format({'align': 'center', 'font_name': 'Arial', 'bold': False, 'size': 8, 'border': 7})
          #  date_format = workbook.add_format({'num_format': 'mmm/yy', 'align': 'center', 'font_name': 'Arial', 'bold': True, 'size': 8,
           #                                    'font_color': '#FFFFFF',  'bg_color': '#003366', 'border': 7})

            # Criacao das abas
            for aba in listAbaExcel_Anual:
                ###
                ### HEADER
                ###
                worksheet = workbook.add_worksheet(str(aba[1]))
                worksheet.freeze_panes('B13')  # Freeze the row and columns
                worksheet.set_default_row(11.25)  # default row heigh
                worksheet.hide_gridlines(2)  # hide the grid lines.
                worksheet.set_row(0, 11.25)  # rows tamanho
                worksheet.set_row(1, 23.25)  # rows tamanho
                worksheet.set_row(2, 11.25)  # rows tamanho
                worksheet.set_row(3, 5)  # rows tamanho
                for row in range(4, 11):
                    worksheet.set_row(row, 11.25)  # rows tamanho
                # Column
                worksheet.set_column(0, 0, 11.83)

                # set Logo
                worksheet.insert_image('A1', 'logo.jpg', {'x_offset': 15, 'y_offset': 5})

                worksheet.write('B2', str(aba[2]), titulo_format)
                worksheet.write('B3', 'Para mais informações sobre as séries, clique no ícone "+" à esquerda',
                                inf_format)
                string_serie = [desc_b_format, 'Série: ', desc_format, str(aba[2])]
                string_fonte = [desc_b_format, 'Fonte: ', desc_format, str(fonte)]
                string_uni = [desc_b_format, 'Unidade: ', desc_format, str(aba[3])]
                string_nota = [desc_b_format, 'Notas: ', desc_format, str(nota)]
                worksheet.write_rich_string('C5', *string_serie)
                worksheet.write_rich_string('C6', *string_fonte)
                worksheet.write_rich_string('C7', *string_uni)
                worksheet.write_rich_string('C8', *string_nota)
                worksheet.write('C9',
                                '*Dados organizados exclusivamente para clientes da LCA. A reprodução desta tabela é proibida.',
                                alerta_format)

                # Set Grouping
                worksheet.set_row(4, None, None, {'level': 1, 'hidden': True})
                worksheet.set_row(5, None, None, {'level': 1, 'hidden': True})
                worksheet.set_row(6, None, None, {'level': 1, 'hidden': True})
                worksheet.set_row(7, None, None, {'level': 1, 'hidden': True})
                worksheet.set_row(8, None, None, {'level': 1, 'hidden': True})
                worksheet.set_row(9, None, None, {'collapsed': True})

                ###
                ### BODY
                ###

                # Importando e tratando os dados do Sql server
                # Criando uma nova tabela auxiliar para poder deletar as colunas desnecessarias. Devido as colunas serem dinamicas, nao eh possivel aponta-las no select
                query = "USE " + db + " IF object_id('TB_AUX_MERC_EXPORT_FINAL') is not null DROP TABLE " + db + ".dbo.TB_AUX_MERC_EXPORT_FINAL;"
                cursor.execute(query)
                conexao.commit()
                if (aba[4] != None):
                    query = "select * into " + db + ".dbo.TB_AUX_MERC_EXPORT_FINAL from " + db + ".dbo.TB_AUX_MERC_EXPORT WHERE indicador =" + str(aba[0]) + " and indicadorDetalhe='" + str(aba[4]) + "';"
                else:
                    query = "select * into " + db + ".dbo.TB_AUX_MERC_EXPORT_FINAL from " + db + ".dbo.TB_AUX_MERC_EXPORT WHERE indicador =" + str(aba[0]) + ";"
                cursor.execute(query)
                conexao.commit()
                # deletando as colunas que nao devem aparecer no excel
                if (nome_arquivo =='Focus-BCB_Anual_Mediana') or (nome_arquivo=='Focus-BCB_Anual_Media'):
                    query = "alter table TB_AUX_MERC_EXPORT_FINAL drop column indicador, indicadorDetalhe"
                if (nome_arquivo =='Focus-BCB_Trimestral_Mediana') or (nome_arquivo=='Focus-BCB_Trimestral_Media'):
                    query = "alter table TB_AUX_MERC_EXPORT_FINAL drop column indicador"
                cursor.execute(query)
                conexao.commit()

                query = "select * from "+db+".dbo.TB_AUX_MERC_EXPORT_FINAL order by convert(date,periodo)"
                mysel = cursor.execute(query)

                # column names
                cols = [i[0] for i in cursor.description]
                for c, col in enumerate(cols):

                    worksheet.write(10, c, '', column_format)
                    if c == 0:  # para a coluna "periodo"
                        worksheet.write(10, c, col, Per_col_format)
                        worksheet.write(11, c, '', Per_col_format)
                    else:
                        #if (nome_arquivo == 'Focus-BCB_Anual_Mediana') or (nome_arquivo == 'Focus-BCB_Anual_Media'):
                            worksheet.write(11, c, col, column_format)
                        #else:
                        #    date_time = datetime.strptime(col, '%d/%m/%Y')
                         #   worksheet.write_datetime(11, c, date_time, date_format)

                    tabela_format = workbook.add_format(
                        {'align': 'center_across', 'font_name': 'Arial', 'bold': True, 'italic': False, 'size': 8,
                         'font_color': '#FFFFFF', 'bg_color': '#003366'})
                    found = re.search(' (.*)', str(aba[2])).group(0)
                    worksheet.write('B11', str(found), tabela_format)
                    # alinhamento centralizando selecao de acordo com o numero de colunas
                    if c != 0:
                        cell = xl_rowcol_to_cell(10, c)
                        worksheet.write(cell, '', tabela_format)
                        worksheet.set_column("B12:"+cell,6.7) # tamanho das colunas onde se encontram os valores dos indicadores
                # dados
                for i, row in enumerate(mysel):
                    for j, value in enumerate(row):
                        worksheet.write(i + 12, j, value, row_format)
                worksheet.set_selection(i+12,0,i+12,0)
            workbook.close()

            #AUTOFIT COLUMNS
            #faz com que as colunas fiquem do tamanho da string
            #Porem tem um custo maior de processamento ( tempo aumenta cerca de
           # import win32com.client as win32
           # excel = win32.gencache.EnsureDispatch("Excel.Application")
          #  wb = excel.Workbooks.Open(r"C:/Users/william.wada/Desktop/OneDrive - LCA Consultores/Tarefas/Bacen/API/"+nome_arquivo+".xlsx")

           # for aba in listAbaExcel_Anual:
          #      ws = wb.Worksheets(str(aba[1]))
           #     ws.Range("B12:"+cell).Columns.AutoFit()
          #      wb.Save()
          #  excel.Application.Quit()
            print("Excel gerado com sucesso")
        except Exception as e:
            print(e)
            print()
            print("Erro na geracao do excel.")





