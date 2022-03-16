
import win32com.client
import pythoncom
import os
from csv import reader

with open('chamada.csv',  encoding='utf-8') as read_obj:
    csv_reader = reader(read_obj)
    estudantes = []
    for row in csv_reader:
        estudantes.append(row)
        print(row)

pythoncom.CoInitialize()
psApp = win32com.client.Dispatch("Photoshop.Application")

psApp.Open('C:/Users/iago-/Documents/ensino médio/estágio/programa 1/certificadominepsd.psd')
doc = psApp.Application.ActiveDocument
for estudante in estudantes:
    layernome = doc.ArtLayers["nome"]
    textoflayernome = layernome.TextItem
    textoflayernome.contents = estudante[0]

    folder = "C:/Users/iago-/Documents/ensino médio/estágio/programa 1/certificados"
    filename = estudante[0] + '.jpeg'
    full_path = os.path.join(folder, filename)
    options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")
    options.Format = 6
    options.Quality = 12

    doc.Export(ExportIn=full_path, ExportAs = 2, Options = options)