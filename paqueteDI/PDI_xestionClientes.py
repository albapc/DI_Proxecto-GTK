"""
PDI_xestionClientes.py
====================================
Arquivo que contén a interfaz do menú de produtos/servizos
"""

import gi
import sqlite3

from reportlab.lib.styles import getSampleStyleSheet

gi.require_version('Gtk', '3.0')
from gi.repository import Gtk

from reportlab.platypus import (SimpleDocTemplate, Spacer, Table, Paragraph)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
import datetime


class VentaClientes(Gtk.Window):
    """
    Clase que conten as funcións do menú para xestionar os clientes dispoñibles
    """

    def __init__(self):
        """
        Función que se executa cada vez que se inicia a clase e inicializa os elementos contidos nela (elementos físicos,
        base de datos...)
        """
        Gtk.Window.__init__(self, title="Xestión de clientes")
        self.set_border_width(3)

        self.set_default_size(830, 200)

        self.connFlag = True

        self.conn = sqlite3.connect('lista_clientes.db')
        self.cursor = self.conn.cursor()

        self.cursor.execute('''CREATE TABLE if not exists clientes
                     (dni text, nome text, apelidos text, enderezo text, telefono text, sexo text, datan text)''')

        self.cursor.execute("""INSERT INTO clientes(dni, nome, apelidos, enderezo, telefono, sexo, datan) 
SELECT '11111111A', 'Carlos', 'Fernandez Rodriguez', 'Pizarro 94', '19841919', 'Home', '2020-3-17'
WHERE NOT EXISTS(SELECT 1 FROM clientes WHERE dni = '11111111A')""")

        self.cursor.execute("""INSERT INTO clientes(dni, nome, apelidos, enderezo, telefono, sexo, datan) 
        SELECT '22222222B', 'Isabel', 'Iglesias Roncero', 'Calvario 8', '987987987', 'Muller', '2020-3-3'
        WHERE NOT EXISTS(SELECT 1 FROM clientes WHERE dni = '22222222B')""")

        self.cursor.execute("""INSERT INTO clientes(dni, nome, apelidos, enderezo, telefono, sexo, datan) 
        SELECT '38585626H', 'Alba', 'Perez', 'Calle Falsa 123', '432423423423', 'Muller', '1993-10-15'
        WHERE NOT EXISTS(SELECT 1 FROM clientes WHERE dni = '38585626H')""")

        self.cursor.execute("""INSERT INTO clientes(dni, nome, apelidos, enderezo, telefono, sexo, datan) 
        SELECT '33333333C', 'Antonio', 'Perez Reverte', 'RAE 35', '123456789', 'Home', '2020-3-25'
        WHERE NOT EXISTS(SELECT 1 FROM clientes WHERE dni = '33333333C')""")

        self.cursor.execute("""INSERT INTO clientes(dni, nome, apelidos, enderezo, telefono, sexo, datan) 
        SELECT '4444444D', 'Pepe', 'Navarro', 'Avenida de Madrid 76', '78484848948', 'Home', '1967-7-19'
        WHERE NOT EXISTS(SELECT 1 FROM clientes WHERE dni = '4444444D')""")

        self.cursor.execute("""INSERT INTO clientes(dni, nome, apelidos, enderezo, telefono, sexo, datan) 
        SELECT '123456789Z', 'Carlos', 'Rodriguez Francisco', 'Colombia 53', '321313123', 'Home', '2020-3-3'
        WHERE NOT EXISTS(SELECT 1 FROM clientes WHERE dni = '123456789Z')""")

        self.cursor.execute("""INSERT INTO clientes(dni, nome, apelidos, enderezo, telefono, sexo, datan) 
        SELECT '987654321G', 'Maria', 'Torres Pardo', 'Mexico 123', '942342342', 'Muller', '2020-3-20'
        WHERE NOT EXISTS(SELECT 1 FROM clientes WHERE dni = '987654321G')""")

        self.conn.commit()

        self.v = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        self.notebook = Gtk.Notebook()
        self.v.pack_start(self.notebook, False, False, 0)

        self.v2 = Gtk.Box(spacing=6)
        self.bVoltar = Gtk.Button(label="Voltar ao menú principal")
        self.bVoltar.set_property("width-request", 400)
        self.bVoltar.connect("clicked", self.on_bVoltar_clicked)
        self.v2.pack_start(self.bVoltar, True, False, 0)
        self.v.pack_start(self.v2, True, True, 0)

        self.add(self.v)

        self.box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)

        self.pInsertar = Gtk.Box(spacing=6)
        self.box.set_border_width(10)
        self.box.pack_start(self.pInsertar, False, False, 0)
        self.pInsertar.pack_start(Gtk.Label('DNI: '), True, True, 0)
        self.eDni = Gtk.Entry()
        self.pInsertar.pack_start(self.eDni, True, True, 0)

        self.pInsertar.pack_start(Gtk.Label('Nome: '), True, True, 0)
        self.eNome = Gtk.Entry()
        self.pInsertar.pack_start(self.eNome, True, True, 0)

        self.pInsertar.pack_start(Gtk.Label('Apelidos: '), True, True, 0)
        self.eApelidos = Gtk.Entry()
        self.pInsertar.pack_start(self.eApelidos, True, True, 0)

        self.pInsertar2 = Gtk.Box(spacing=6)
        self.pInsertar2.pack_start(Gtk.Label('Enderezo: '), True, True, 0)
        self.eEnderezo = Gtk.Entry()
        self.eEnderezo.set_width_chars(70)
        self.pInsertar2.pack_start(self.eEnderezo, True, True, 0)

        self.box.pack_start(self.pInsertar2, False, False, 0)

        self.pInsertar3 = Gtk.Box(spacing=6)
        self.pInsertar3.pack_start(Gtk.Label('Teléfono: '), True, True, 0)
        self.eTelefono = Gtk.Entry()
        self.pInsertar3.pack_start(self.eTelefono, True, True, 0)
        self.pInsertar3.pack_start(Gtk.Label('Sexo: '), True, True, 0)
        self.cmbSexo = Gtk.ComboBoxText()
        self.sexos = ["Home", "Muller", "Non especificado"]
        self.cmbSexo.set_entry_text_column(0)
        for s in self.sexos:
            self.cmbSexo.append_text(s)
        self.pInsertar3.pack_start(self.cmbSexo, True, True, 0)

        self.box.pack_start(self.pInsertar3, False, False, 0)

        self.pInsertar4 = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        self.pInsertar4.pack_start(Gtk.Label('Data de nacemento: '), False, False, 0)

        self.calendario = Gtk.Calendar()
        self.pInsertar4.pack_start(self.calendario, True, True, 0)

        self.box.pack_start(self.pInsertar4, True, True, 0)

        self.pInsertar5 = Gtk.Box(spacing=6)
        self.bAceptar = Gtk.Button(label="Aceptar")
        self.bAceptar.connect("clicked", self.on_bAceptar_clicked)
        self.pInsertar5.pack_start(self.bAceptar, True, True, 0)
        self.bBorrar = Gtk.Button(label="Borrar campos")
        self.bBorrar.connect("clicked", self.on_bBorrar_clicked)
        self.pInsertar5.pack_start(self.bBorrar, True, True, 0)

        self.box.pack_start(self.pInsertar5, False, False, 0)

        self.notebook.append_page(self.box, Gtk.Label('Insertar cliente'))

        self.page2 = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.page2.set_border_width(10)

        self.pConsultar1 = Gtk.Box(spacing=6)
        self.pConsultar1.pack_start(Gtk.Label('Buscar por: '), False, False, 0)
        self.cmbFiltrar = Gtk.ComboBoxText()
        self.opcionesBusqueda = ["DNI", "Nome", "Apelidos", "Enderezo", "Telefono", "Sexo", "Data de nacemento"]
        for s in self.opcionesBusqueda:
            self.cmbFiltrar.append_text(s)
        self.pConsultar1.pack_start(self.cmbFiltrar, False, False, 0)
        self.busqueda = Gtk.Entry()
        self.bBuscar = Gtk.Button(label="Buscar")
        self.bBuscar.connect("clicked", self.on_bBuscar_clicked)

        self.pConsultar1.pack_start(self.busqueda, True, True, 0)
        self.pConsultar1.pack_start(self.bBuscar, True, True, 0)
        self.page2.pack_start(self.pConsultar1, False, False, 10)

        self.pBorrar = Gtk.Box(spacing=6)
        self.bBorrarCampos = Gtk.Button(label="Eliminar cliente(s) seleccionado(s)")
        self.bBorrarCampos.connect("clicked", self.on_bBorrarCampos_clicked)
        self.bBorrarCampos.set_sensitive(False)
        self.bVerTodo = Gtk.Button(label="Ver tódolos clientes")
        self.bVerTodo.connect("clicked", self.on_bVerTodo_clicked)
        self.pBorrar.pack_start(self.bBorrarCampos, True, True, 0)
        self.pBorrar.pack_start(self.bVerTodo, True, True, 0)
        self.page2.pack_start(self.pBorrar, False, False, 10)

        self.pConsultar2 = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.labelExplicacion1 = Gtk.Label()
        self.labelExplicacion1.set_markup("<b>Faga clic nas cabeceiras da táboa para ordear por categoría.</b>")
        self.pConsultar2.pack_start(self.labelExplicacion1, False, False, 0)

        self.labelExplicacion2 = Gtk.Label()
        self.labelExplicacion2.set_markup("<b>Pode seleccionar varios clientes facendo clic + Ctrl nas filas que "
                                          "desexe</b>")
        self.pConsultar2.pack_start(self.labelExplicacion2, False, False, 0)
        self.modelo = Gtk.ListStore(str, str, str, str, str, str, str)

        self.select_clientes()

        self.treeview = Gtk.TreeView(model=self.modelo)

        renderer_text = Gtk.CellRendererText()
        column_text = Gtk.TreeViewColumn("DNI", renderer_text, text=0)
        column_text.set_sort_column_id(0)
        self.treeview.append_column(column_text)

        renderer_text = Gtk.CellRendererText()
        column_text = Gtk.TreeViewColumn("Nome", renderer_text, text=1)
        column_text.set_sort_column_id(1)
        self.treeview.append_column(column_text)

        renderer_text = Gtk.CellRendererText()
        column_text = Gtk.TreeViewColumn("Apelidos", renderer_text, text=2)
        column_text.set_sort_column_id(2)
        self.treeview.append_column(column_text)

        renderer_text = Gtk.CellRendererText()
        column_text = Gtk.TreeViewColumn("Enderezo", renderer_text, text=3)
        column_text.set_sort_column_id(3)
        self.treeview.append_column(column_text)

        renderer_text = Gtk.CellRendererText()
        column_text = Gtk.TreeViewColumn("Teléfono", renderer_text, text=4)
        column_text.set_sort_column_id(4)
        self.treeview.append_column(column_text)

        renderer_text = Gtk.CellRendererText()
        column_text = Gtk.TreeViewColumn("Sexo", renderer_text, text=5)
        column_text.set_sort_column_id(5)
        self.treeview.append_column(column_text)

        renderer_text = Gtk.CellRendererText()
        column_text = Gtk.TreeViewColumn("Data de nacemento", renderer_text, text=6)
        column_text.set_sort_column_id(6)
        self.treeview.append_column(column_text)

        self.tree_selection = self.treeview.get_selection()
        self.tree_selection.set_mode(Gtk.SelectionMode.MULTIPLE)
        self.tree_selection.connect("changed", self.onSelectionChanged)

        scrollTree = Gtk.ScrolledWindow()
        scrollTree.add(self.treeview)
        self.pConsultar2.pack_start(scrollTree, True, True, 0)
        self.page2.pack_start(self.pConsultar2, True, True, 0)

        self.pConsultar3 = Gtk.Box()
        self.bXerarUnCliente = Gtk.Button(label="Xerar informe cliente(s) seleccionado(s)")
        self.bXerarUnCliente.set_sensitive(False)
        self.bXerarUnCliente.connect("clicked", self.on_bXerarUnCliente_clicked)
        self.pConsultar3.pack_start(self.bXerarUnCliente, True, True, 10)
        self.bXerarTodolosClientes = Gtk.Button(label="Xerar informe de tódolos clientes")
        self.bXerarTodolosClientes.connect("clicked", self.on_bXerarTodolosClientes_clicked)
        self.pConsultar3.pack_start(self.bXerarTodolosClientes, True, True, 10)
        self.bBorrarSeleccion = Gtk.Button(label="Desfacer selección")
        self.bBorrarSeleccion.connect("clicked", self.on_bBorrarSeleccion_clicked)
        self.pConsultar3.pack_start(self.bBorrarSeleccion, True, True, 10)
        self.page2.pack_start(self.pConsultar3, False, False, 6)

        self.notebook.connect("switch-page", self.on_notebook_switch_page)

        self.notebook.append_page(self.page2, Gtk.Label('Consultar/Eliminar cliente'))

    def check_connection(self):
        """
        Comproba a conexión da base de datos. Se non está conectada, realiza a conexión

        :return:
        """
        if not self.connFlag:
            self.conn = sqlite3.connect('lista_clientes.db')
            self.cursor = self.conn.cursor()
            self.connFlag = True

    def select_clientes(self):
        """
        Recolle todolos produtos contidos na taboa clientes da base de datos e os inserta no treeview

        :return:
        """
        for row in self.cursor.execute("select * from clientes"):
            self.modelo.append([row[0], row[1], row[2], row[3], row[4], row[5], row[6]])

    def on_bAceptar_clicked(self, widget):
        """
        Función que se executa ao pulsar o botón "Aceptar" que inserta os datos introducidos polo usuario na base de
        datos e borra os datos introducidos nas caixas de texto

        :param widget:
        :return:
        """
        tupla = self.calendario.get_date()
        anho = str(tupla[0])
        mes = str(tupla[1] + 1)
        dia = str(tupla[2])

        fecha = str(anho + "-" + mes + "-" + dia)

        self.check_connection()

        t = (self.eDni.get_text(), self.eNome.get_text(), self.eApelidos.get_text(), self.eEnderezo.get_text(),
             self.eTelefono.get_text(), self.cmbSexo.get_active_text(), fecha)

        self.cursor.execute('INSERT INTO clientes VALUES (?,?,?,?,?,?,?)', t)
        self.conn.commit()
        self.conn.close()
        self.connFlag = False

        now = datetime.datetime.now()
        self.calendario.select_day(now.day)
        self.calendario.select_month(now.month - 1, now.year)  # pone un mes de más por defecto en el calendario, aunque
        # imprime el mes correctamente, por ello se resta

        self.eDni.set_text("")
        self.eNome.set_text("")
        self.eApelidos.set_text("")
        self.eEnderezo.set_text("")
        self.eTelefono.set_text("")
        self.cmbSexo.set_active(0)

    def on_bBorrar_clicked(self, widget):
        """
        Función que se executa ao pulsar o botón "Borrar" que resetea o formulario aos valores iniciais

        :param widget:
        :return:
        """
        now = datetime.datetime.now()
        self.calendario.select_day(now.day)
        self.calendario.select_month(now.month - 1, now.year)  # pone un mes de más por defecto en el calendario, aunque
        # imprime el mes correctamente, por ello se resta

        self.eDni.set_text("")
        self.eNome.set_text("")
        self.eApelidos.set_text("")
        self.eEnderezo.set_text("")
        self.eTelefono.set_text("")
        self.cmbSexo.set_active(0)

    def onSelectionChanged(self, tree_selection):
        """
        Obtén o índice da fila seleccionada do treeview, se hai unha fila seleccionada actívanse os botóns de borrar e
        xerar o informe dun cliente

        :param tree_selection:
        :return:
        """
        (self.model, self.pathlist) = tree_selection.get_selected_rows()
        self.bXerarUnCliente.set_sensitive(True)
        self.bBorrarCampos.set_sensitive(True)

    def on_bXerarUnCliente_clicked(self, widget):
        """
        Función que se executa ao pulsar o botón "Xerar informe cliente/s seleccionado/s". Crea un arquivo pdf coas fichas
        dos clientes (filas) seleccionados.

        :param widget:
        :return:
        """
        story = []
        titulo = ''

        self.check_connection()

        styles = getSampleStyleSheet()

        for path in self.pathlist:
            tree_iter = self.model.get_iter(path)
            self.value = self.model.get_value(tree_iter, 0)
            self.nome = self.model.get_value(tree_iter, 1)
            self.apelidos = self.model.get_value(tree_iter, 2)
            self.enderezo = self.model.get_value(tree_iter, 3)
            self.telefono = self.model.get_value(tree_iter, 4)
            self.sexo = self.model.get_value(tree_iter, 5)
            self.datan = self.model.get_value(tree_iter, 6)

            ptext = Paragraph('Ficha cliente con DNI %s: ' % self.value, styles['title'])

            story.append(ptext)

            t = Table([
                ['DNI:', self.value],
                ['Nome:', self.nome],
                ['Apelidos:', self.apelidos],
                ['Enderezo:', self.enderezo],
                ['Teléfono:', self.telefono],
                ['Sexo:', self.sexo],
                ['Data de nacemento:', self.datan]
            ], colWidths=200, rowHeights=30)

            t.setStyle([
                ('TEXTCOLOR', (0, 0), (0, -1), colors.purple),
                ('TEXTCOLOR', (1, 0), (1, -1), colors.darkblue),
                ('BACKGROUND', (0, 0), (-1, -1), colors.cyan),
                ('BOX', (0, 0), (-1, -1), 2, colors.magenta),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ])
            story.append(t)
            story.append(Spacer(0, 15))
            titulo += self.value

        doc = SimpleDocTemplate("fcliente_%s.pdf" % titulo, pagesize=A4)
        doc.build(story)

    def on_bXerarTodolosClientes_clicked(self, widget):
        """
        Función que se executa ao pulsar o botón "Xerar informe de tódolos clientes". Xera un arquivo pdf coas fichas de
        tódolos clientes da base de datos.

        :param widget:
        :return:
        """
        doc = SimpleDocTemplate("listaClientes.pdf", pagesize=A4)
        story = []

        self.check_connection()

        styles = getSampleStyleSheet()

        i = 1

        for row in self.cursor.execute("select * from clientes"):
            ptext = Paragraph('Ficha cliente nº %s: ' % i, styles['title'])

            story.append(ptext)

            i += 1

            t = Table([
                ['DNI:', row[0]],
                ['Nome:', row[1]],
                ['Apelidos:', row[2]],
                ['Enderezo:', row[3]],
                ['Teléfono:', row[4]],
                ['Sexo:', row[5]],
                ['Data de nacemento:', row[6]]
            ], colWidths=200, rowHeights=30)

            t.setStyle([
                ('TEXTCOLOR', (0, 0), (0, -1), colors.purple),
                ('TEXTCOLOR', (1, 0), (1, -1), colors.darkblue),
                ('BACKGROUND', (0, 0), (-1, -1), colors.cyan),
                ('BOX', (0, 0), (-1, -1), 2, colors.magenta),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ])
            story.append(t)
            story.append(Spacer(0, 15))
        doc.build(story)

    def on_bBorrarSeleccion_clicked(self, widget):
        """
        Función que se executa ao pulsar o botón "Desfacer selección". Desfai a selección de filas e desactiva os botóns
        de xerar un cliente e de borrar campos

        :param widget:
        :return:
        """
        self.tree_selection.unselect_all()
        self.bXerarUnCliente.set_sensitive(False)
        self.bBorrarCampos.set_sensitive(False)

    def on_notebook_switch_page(self, notebook, tab, index):
        """
        Función que se executa ao cambiar de pestaña no notebook. Actualiza o listore e mostra os datos actualizados no
        caso de facer unha nova inserción

        :param notebook:
        :param tab:
        :param index:
        :return:
        """
        self.modelo.clear()
        self.check_connection()
        self.select_clientes()

    def on_bBuscar_clicked(self, widget):
        """
         Función que se executa ao facer clic no botón "Buscar". Compara o valor seleccionado no ComboBox co diccionario
         e busca o campo escrito na fila seleccionada do ComboBox na base de datos

        :param widget:
        :return:
        """
        self.modelo.clear()
        self.check_connection()
        dict = {
            "DNI": "dni",
            "Nome": "nome",
            "Apelidos": "apelidos",
            "Enderezo": "enderezo",
            "Telefono": "telefono",
            "Sexo": "sexo",
            "Data de nacemento": "datan"
        }
        for key in dict:
            if key == self.cmbFiltrar.get_active_text():
                query = "select * from clientes where %s='%s'" % (dict[key], self.busqueda.get_text())
                for row in self.cursor.execute(query):
                    self.modelo.append([row[0], row[1], row[2], row[3], row[4], row[5], row[6]])

    def on_bBorrarCampos_clicked(self, widget):
        """
        Función que se executa ao facer clic no botón "Borrar campos". Busca a fila seleccionada, recolle o seu id e a
        elimina da base de datos. Unha vez realizada a operación, desconectase da base de datos

        :param widget:
        :return:
        """
        self.check_connection()

        for path in reversed(self.pathlist):
            tree_iter = self.model.get_iter(path)
            self.value = self.model.get_value(tree_iter, 0)
            self.cursor.execute("delete from clientes where dni='%s'" % self.value)
            self.model.remove(tree_iter)

        self.conn.commit()
        self.conn.close()
        self.connFlag = False

    def on_bVerTodo_clicked(self, widget):
        """
        Función que se executa ao facer clic no botón "Ver tódolos clientes". Actualiza o treeview e mostra todolos
        clientes da base de datos

        :param widget:
        :return:
        """
        self.modelo.clear()
        self.check_connection()
        self.select_clientes()

    def on_bVoltar_clicked(self, widget):
        """
        Función que se executa ao facer clic no botón "Voltar ao menú". Pecha a conexión coa base de datos e pecha a ventá
        actual

        :param widget:
        :return:
        """
        if self.connFlag:
            self.conn.close()
        self.destroy()
