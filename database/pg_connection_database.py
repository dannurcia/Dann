
import psycopg2

database_data = {}
dict_words = []


def get_db_info():
    # fh = open(resource_path('info/DB_info.txt'))
    fh = open('C:\\Users\\user\\PycharmProjects\\Dann\\DB_info.txt')
    # fh = open('/DB_info.txt')
    for line in fh:
        words = line.rstrip().split()
        dict_words.extend(words)
    database_data['user'] = dict_words[0]  # root
    database_data['password'] = dict_words[1]  # Inictel123.
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
            print(msg)
        try:
            self.connection = psycopg2.connect(host=self.database_data['host'], database=self.database_data['database'],
                                               user=self.database_data['user'], password=self.database_data['password'])
            msg = 'Se estableció la conexión a la base de datos.'
            self.cursor = self.connection.cursor()
            print(msg)
        except:
            msg = 'No se pudo establecer conexión a la base de datos.'
            self.error = True
            print(msg)
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
            self.connection = psycopg2.connect(user=self.database_data['user'],
                                               password=self.database_data['password'],
                                               host=self.database_data['host'],
                                               database=self.database_data['database'])
            self.cursor = self.connection.cursor()
        except:
            self.error = True

    def get_documento(self, codigo_ev):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select codigo_doc from documento where codigo_evi = '{0}'""".format(codigo_ev)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el código del documento. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_user_info(self, id_user):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select * from usuario where codigo_usu = '{0}'""".format(id_user)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener la información del usuario. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    # def get_cod_evidencia(self, id_evidencia):
    #     mobile_records = []
    #     error_msg = ''
    #     self.open_connection()
    #     if not self.error:
    #         try:
    #             select_query = """select codigo_evi from documento where codigo_doc = '{0}'""".format(id_evidencia)
    #             self.cursor.execute(select_query)
    #             mobile_records = self.cursor.fetchall()
    #         except (Exception, psycopg2.Error) as error:
    #             error_msg = 'No se pudo obtener el código de la evidencia. ' + str(error)
    #     self.close_connection()
    #     return mobile_records, error_msg

    def get_introduccion(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select introduccion_doc from documento where codigo_doc = '{0}'""".format(id_document[0][0][0])
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de la Introducción. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_antecedentes(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select antecendentes_doc from documento where codigo_doc = '{0}'""".format(id_document[0][0][0])
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Antecedentes. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_objectivo_general(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select objetivogeneral_doc from documento where codigo_doc = '{0}'""".format(id_document[0][0][0])
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto del Objetivo General. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_objectivos_especificos(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select objetivosespecificos_doc from documento where codigo_doc = '{0}'""".format(id_document[0][0][0])
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Objetivos Específicos. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_recursos_humanos(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select recursos_humanos_doc from documento where codigo_doc = '{0}'""".format(id_document[0][0][0])
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Recursos Humanos. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_recursos_materiales(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select recursos_materiales_doc from documento where codigo_doc = '{0}'""".format(id_document[0][0][0])
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Recursos Materiales. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_otros_recursos(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select otros_recursos_doc from documento where codigo_doc = '{0}'""".format(id_document[0][0][0])
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de Otros Recursos. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_cuerpo_evidencia(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select cuerpo_evidencia_doc from documento where codigo_doc = '{0}'""".format(id_document[0][0][0])
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de las Actividades Desarrolladas. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_resultados(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select resultados_doc from documento where codigo_doc = '{0}'""".format(id_document[0][0][0])
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Resultados. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_comentarios(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select comentarios_doc from documento where codigo_doc = '{0}'""".format(id_document[0][0][0])
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Comentarios. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_conclusiones(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select otros1_doc from documento where codigo_doc = '{0}'""".format(id_document[0][0][0])
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de las Conclusiones. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_referencias(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select referencias_doc from documento where codigo_doc = '{0}'""".format(id_document[0][0][0])
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de las Referencias. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_apendices(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select apendices_doc from documento where codigo_doc = '{0}'""".format(id_document[0][0][0])
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Apéndices. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_anexos(self, id_document):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select anexos_doc from documento where codigo_doc = '{0}'""".format(id_document[0][0][0])
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de los Anexos. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_descripcion_eviGdR(self, codigo_evi):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select descripcion_evi from evidenciagdr where codigo_evi = '{0}'""".format(codigo_evi)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de la Evidencia. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_prioridadGdR(self, codigo_evi):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                # select_query = """select * from prioridadgdr where codigo_pri = '{0}'""".format(codigo_pri)
                select_query = """select codigo_pri from indicadorgdr inner join evidenciagdr ON evidenciagdr.codigo_ind = indicadorgdr.codigo_ind where evidenciagdr.codigo_evi = '{0}'""".format(codigo_evi)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de la Prioridad. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

    def get_descripcion_indGdR(self, codigo_ind):
        mobile_records = []
        error_msg = ''
        self.open_connection()
        if not self.error:
            try:
                select_query = """select descripcion_ind from indicadorgdr inner join evidenciagdr ON evidenciagdr.codigo_ind = indicadorgdr.codigo_ind where evidenciagdr.codigo_evi = '{0}'""".format(codigo_ind)
                self.cursor.execute(select_query)
                mobile_records = self.cursor.fetchall()
            except (Exception, psycopg2.Error) as error:
                error_msg = 'No se pudo obtener el texto de la Evidencia. ' + str(error)
        self.close_connection()
        return mobile_records, error_msg

