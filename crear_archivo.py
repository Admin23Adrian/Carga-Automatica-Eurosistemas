from os import listdir, path
import shutil


class CrearArhivoTxt:
    def __init__(self, src, dst):
        self.src = src
        self.dst = dst
    
    def crearTxt(self, nombreTxt):
        try:
            with open(self.src + nombreTxt, "w") as archivo:
                archivo.write("He creado esta linea.")
            return "Se ha creado el txt.", self.src + nombreTxt
        except:
            return "No se ha podido crear el txt."
    
    def escribirTxt(self):
        pass

        
    
ruta_guardado = "C:/Users/aalarcon/Desktop/"

# Creamos el objeto de la clase CrearArchivoTxt
mi_txt = CrearArhivoTxt(ruta_guardado, "Aca mi destino.")
resultado_crear_archivo, ruta_txt = mi_txt.crearTxt("Prueba.txt")
print(ruta_txt)



