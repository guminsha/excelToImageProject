"""
Pegar os dados da planilha 
Tipo nome do curso, nome participante, tipo de participação, data do inicio, data do final, carga horária, data da emissão do certificado

Transferir os dados da planilha para a imagem do certificado
"""
import openpyxl

workbook_alunos = openpyxl.load_workbook("src\\assets\\planilha_alunos.xlsx")

print(workbook_alunos)

