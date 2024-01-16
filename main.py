import ifcopenshell
import ifcopenshell.geom
import ifcopenshell.util.shape
import pandas as pd
import PySimpleGUI as sg

''' ________________________________________________________________________________________________________
Criação da interface gráfica 
____________________________________________________________________________________________________________'''

sg.theme('Reddit')
layout = [
    [sg.Text('Selecione o arquivo ifc: '), sg.FileBrowse(key='file1')],
    [sg.Button('Gerar')]
]

window = sg.Window('Contador de polígonos', layout)

while True:
    event, values = window.read()
    # Quando a janela for fechada
    if event == sg.WIN_CLOSED:
        break
    if event == 'Gerar':
        break

window.close()

file1 = values['file1']

'''-----------------------------------------------------------------------------------------------------------------
Abrindo arquivos
-------------------------------------------------------------------------------------------------------------------'''

ifc_file = ifcopenshell.open(file1)

elements = ifc_file.by_type('IfcElement')

lista = []

for element in elements:
    settings = ifcopenshell.geom.settings()
    try:
        shape = ifcopenshell.geom.create_shape(settings, element)
        grouped_verts = ifcopenshell.util.shape.get_vertices(shape.geometry)
        grouped_edges = ifcopenshell.util.shape.get_edges(shape.geometry)
        grouped_faces = ifcopenshell.util.shape.get_faces(shape.geometry)

        dic = {
            'Global id': element.GlobalId,
            'Classe Ifc': element.is_a(),
            'Vértices': len(grouped_verts),
            'Edges': len(grouped_edges),
            'Faces': len(grouped_faces)
        }

        lista.append(dic)
    except:

        dic = {
            'Global id': element.GlobalId,
            'Classe Ifc': element.is_a(),
            'Vértices': 0,
            'Edges': 0,
            'Faces': 0
        }

        lista.append(dic)

df1 = pd.DataFrame(lista)
print(df1)


# Especificar o nome do arquivo Excel de saída
nome_arquivo_excel = "saida_excel.xlsx"

# Salvar o DataFrame como um arquivo Excel
df1.to_excel(nome_arquivo_excel, index=False)
