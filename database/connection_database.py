import sys
import os
import mysql.connector
# import psycopg2

database_data = {}
dict_words = []


# def resource_path(relative_path):
#     if hasattr(sys, '_MEIPASS'):
#         return os.path.join(sys._MEIPASS, relative_path)
#     return os.path.join(os.path.abspath("."), relative_path)


def get_db_info():
    # fh = open(resource_path('info/DB_info.txt'))
    fh = open('C:\\Users\\usuario\\PycharmProjects\\SGD\\info\\DB_info.txt')
    for line in fh:
        words = line.rstrip().split()
        dict_words.extend(words)
    database_data['user'] = dict_words[0]  # root
    if dict_words[1] == '-':
        dict_words[1] = ''
    database_data['password'] = dict_words[1]  # -
    database_data['host'] = dict_words[2]  # localhost
    database_data['database'] = dict_words[3]  #sistemasgd
    return database_data


class DatabaseConnection:
    def __init__(self):
        try:
            self.database_data = get_db_info()
            # print(self.database_data)
            self.error = False
        except:
            msg = 'El archivo con las credenciales de conexión a la base de datos ha sido corrompido.'
            self.error = True
        try:
            self.connection = mysql.connector.connect(user=self.database_data['user'],
                                                      password=self.database_data['password'],
                                                      host=self.database_data['host'],
                                                      database=self.database_data['database'])
            msg = 'Se estableció la conexión a la base de datos.'
            self.cursor = self.connection.cursor()
        except:
            msg = 'No se pudo establecer conexión a la base de datos.'
            self.error = True
        self.close_connection()
        # print(msg)

    def close_connection(self):
        if not self.error:
            self.cursor.close()
            self.connection.close()
            self.error = True

    def open_connection(self):
        self.error = False
        try:
            self.connection = mysql.connector.connect(user=self.database_data['user'],
                                                      password=self.database_data['password'],
                                                      host=self.database_data['host'],
                                                      database=self.database_data['database'])
            self.cursor = self.connection.cursor()
        except:
            self.error = True

    # def get_user_info(self, id_user):
    #     error_msg = ''
    #     self.open_connection()
    #     if not self.error:
    #         try:
    #             select_query = """select * from usuarios where idUsuario = '{0}'""".format(id_user)
    #             self.cursor.execute(select_query)
    #             mobile_records = self.cursor.fetchall()
    #         except (Exception, psycopg2.Error) as error:
    #             error_msg = 'No se pudo obtener la información del usuario. ' + str(error)
    #     self.close_connection()
    #     return mobile_records, error_msg
    #
    # def get_document_info(self, id_document):
    #     error_msg = ''
    #     self.open_connection()
    #     if not self.error:
    #         try:
    #             select_query = """select * from documentos where idDocumento = '{0}'""".format(id_document)
    #             self.cursor.execute(select_query)
    #             mobile_records = self.cursor.fetchall()
    #         except (Exception, psycopg2.Error) as error:
    #             error_msg = 'No se pudo obtener los usuarios. ' + str(error)
    #     self.close_connection()
    #     return mobile_records, error_msg

    def get_user_info(self, id_user):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select * from usuario where codigo_usuario = '{0}'""".format(id_user)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, mysql.connector.Error) as error:
                error_msg = 'No se pudo obtener la información del usuario. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_introduccion(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select text1 from informe where cod_inf = '{0}'""".format(id_document)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, mysql.connector.Error) as error:
                error_msg = 'No se pudo obtener el texto de la Introducción. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_antecedentes(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select text2 from informe where cod_inf = '{0}'""".format(id_document)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, mysql.connector.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Antecedentes. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    # def get_objectivos(self, id_document):
    #     mobile_records = []
    #     error_msg = ''
    #     self.open_connection()
    #     if not self.error:
    #         try:
    #             select_query = """select text3 from informe where cod_inf = '{0}'""".format(id_document)
    #             self.cursor.execute(select_query)
    #             mobile_records = self.cursor.fetchall()
    #         except (Exception, mysql.connector.Error) as error:
    #             error_msg = 'No se pudo obtener el texto de los Objetivos. ' + str(error)
    #     self.close_connection()
    #     return mobile_records, error_msg

    def get_objectivo_general(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select text3 from informe where cod_inf = '{0}'""".format(id_document)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, mysql.connector.Error) as error:
                error_msg = 'No se pudo obtener el texto del Objetivo General. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_objectivos_especificos(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select text4 from informe where cod_inf = '{0}'""".format(id_document)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, mysql.connector.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Objetivos Específicos. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    # def get_recursos(self, id_document):
    #     mobile_records = []
    #     error_msg = ''
    #     self.open_connection()
    #     if not self.error:
    #         try:
    #             select_query = """select text6 from informe where cod_inf = '{0}'""".format(id_document)
    #             self.cursor.execute(select_query)
    #             mobile_records = self.cursor.fetchall()
    #         except (Exception, mysql.connector.Error) as error:
    #             error_msg = 'No se pudo obtener el texto de los Recursos. ' + str(error)
    #     self.close_connection()
    #     return mobile_records, error_msg

    def get_recursos_humanos(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select text5 from informe where cod_inf = '{0}'""".format(id_document)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, mysql.connector.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Recursos Humanos. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_recursos_materiales(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select text6 from informe where cod_inf = '{0}'""".format(id_document)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, mysql.connector.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Recursos Materiales. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_otros_recursos(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select text7 from informe where cod_inf = '{0}'""".format(id_document)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, mysql.connector.Error) as error:
                error_msg = 'No se pudo obtener el texto de Otros Recursos. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_actividades_desarrolladas(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select text8 from informe where cod_inf = '{0}'""".format(id_document)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, mysql.connector.Error) as error:
                error_msg = 'No se pudo obtener el texto de las Actividades Desarrolladas. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_resultados(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select text9 from informe where cod_inf = '{0}'""".format(id_document)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, mysql.connector.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Resultados. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_comentarios(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select text10 from informe where cod_inf = '{0}'""".format(id_document)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, mysql.connector.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Comentarios. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_conclusiones(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select text11 from informe where cod_inf = '{0}'""".format(id_document)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, mysql.connector.Error) as error:
                error_msg = 'No se pudo obtener el texto de las Conclusiones. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_referencias(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select text12 from informe where cod_inf = '{0}'""".format(id_document)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, mysql.connector.Error) as error:
                error_msg = 'No se pudo obtener el texto de las Referencias. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_apendices(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select text13 from informe where cod_inf = '{0}'""".format(id_document)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, mysql.connector.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Apéndices. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_anexos(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select text14 from informe where cod_inf = '{0}'""".format(id_document)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, mysql.connector.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Anexos. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg
