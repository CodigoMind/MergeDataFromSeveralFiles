import os
import re
from io import open
import pandas as pd

# function to get all the full paths
def get_full_paths(list_directories):
    patter = re.compile(".*\.txt")
    full_paths = list()
    for i in range(1,len(list_directories)):
        file = [x for x in list_directories[i][2] if re.match(patter, x)]
        full_paths.append(os.path.join(list_directories[i][0], file[0]))
    return full_paths

# get the data of each file in a excel file
def convert_files_to_excel(full_paths, col_names, patterns):
    # create the dataframe
    data = pd.DataFrame(columns = col_names)

    for i in range(0,len(full_paths)):
    #for i in range(0, 2):
        # get the content
        file_point = open(full_paths[i], "r")
        content = file_point.read()
        file_point.close()
        # get the data
        row = list()
        for j in range(0, len(col_names)):    
            try:
                found = re.search(patterns[j], content, flags=re.DOTALL).group(1).strip().replace("\n", " ")
            except AttributeError:
                found = "NA" # apply your error handling
            row.append(found)
        # add the data to the dataframe
        data.loc[i] = row

    # write the new xlsx file
    os.chdir("..")
    data.to_excel(os.path.join(os.getcwd(), "dataMerged.xlsx"), index = False)
    print(os.getcwd())

if __name__ == "__main__":

    path_directory_father = input("Escriba el nombre de la carpeta que contiene el conjunto de carpeetas a escanear:")
    #path_directory_father = "OneDrive-2022-03-07"
    os.chdir(path_directory_father)
    os.getcwd()

    # get the name of the directories and files
    list_directories = list(os.walk("."))

    # get the paths of each file 
    full_paths = get_full_paths(list_directories = list_directories)

    # set columns name and the patterns to find the data
    col_names = ("Ortofoto_Digital", "Fuente", "Procesamiento", 
                "Proyeccion", "Datum", "Elipsoide",
                "ImagenDim_cols", "ImagenDim_rows",
                "ZonaUTM", "CoorEsqNoroeste_este",
                "CoorEsqNoroeste_norte", "CoorEsqSureste_este",
            "CoorEsqSureste_norte", "Dim_Pixel_XY", "Formato")

    patterns = ("ORTOFOTO DIGITAL:(.+?)\nFUENTE", "FUENTE:(.+?)\nPROCESAMIENTO", 
            "PROCESAMIENTO:(.+?)\nPROYECCION", "PROYECCION:(.+?)\nDATUM", 
            "DATUM:(.+?)\nELIPSOIDE", "ELIPSOIDE:(.+?)\nDIMENSIONES", 
            "DIMENSIONES DE LA IMAGEN:.*Columnas :(.+?)\n.*Renglones", 
            "DIMENSIONES DE LA IMAGEN:.*Columnas :.*\n.*Renglones:(.+?)\nZONA",
            "ZONA UTM:(.+?)\nCOORDENADAS", 
            "COORDENADAS DE LA ESQUINA NOROESTE:.*Este:(.+?)\n.*Norte",
            "COORDENADAS DE LA ESQUINA NOROESTE:.*Este:.*\n.*Norte:(.+?)\nCOORDENADAS",
            "COORDENADAS DE LA ESQUINA SURESTE:.*Este:(.+?)\n.*Norte",
            "COORDENADAS DE LA ESQUINA SURESTE:.*Este:.*\n.*Norte:(.+?)\nDIMENSIONES",
            "DIMENSIONES DEL PIXEL X,Y:(.+?)\nFORMATO", 
            "FORMATO:(.+?)\n")

    convert_files_to_excel(full_paths = full_paths, col_names = col_names, patterns = patterns)