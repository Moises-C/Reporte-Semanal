# autopep8 --in-place --aggressive --aggressive <file>

import fdb
from tkinter import messagebox


class ConexionDB:

    conexion = ""
    conexionAduana = ""
    usuario = "admin"
    contrasenia = "admin"
    querySentencia = ""
    estadoConexion = False
    
    def aplicarCambios(self):
        
        self.conexion.commit()
    
    def establecerConexion(self, aduana):

        conexiones = {

            '160': '192.168.1.20:d:/casawin/CTRAWIN/Datos/CASA_2017.GDB',
            '240': '192.168.240.157:c:/sys_1/CTRAWIN/DATOS/CASA_2016.GDB',
            '430': '192.168.3.4:C:/Google Drive/CASAWIN/BDATOS/CASA.GDB',
            '470': '192.168.90.7:c:/CASAWIN/CARBD7/CASA.GDB'
        }

        self.conexionAduana = conexiones[aduana]
        self.conexionABD()

    def colocarEjecutarQuery(self, query):

        self.setQuery(query)
        self.ejecutarQuery()

    def ejecutarQuery(self):

        self.query.execute(self.querySentencia)

    def conexionABD(self):

        try:

            self.conexion = fdb.connect(
                dsn=self.conexionAduana,
                user=self.usuario,
                password=self.contrasenia)

            self.query = self.conexion.cursor()
            self.setEstadoConexion(True)

        except BaseException:

            self.setEstadoConexion(False)
            messagebox.showerror("Sin conexión",
                                 "No se pudo establecer conexión con la BD.")

    def cerrarConexion(self):

        self.conexion.close()
        self.query.close()

    def getEjecucionQuery(self):

        return self.query

    def getEstadoConexion(self):

        return self.estadoConexion

    def setEstadoConexion(self, estado):

        self.estadoConexion = estado

    def getQuery(self):

        return self.querySentencia

    def setUsuario(self, usuario):

        self.usuario = usuario

    def setContrasenia(self, contrasenia):

        self.contrasenia = contrasenia

    def setQuery(self, query):

        self.querySentencia = query

    def setConexion(self, conexion):

        self.conexion = conexion

    def obtenerUnRegistro(self, columna):

        bandera = False

        for registro in self.getEjecucionQuery():

            return (registro[columna])

        return bandera
