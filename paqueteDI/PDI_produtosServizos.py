"""
PDI_produtosServizos.py
====================================
Arquivo que contén a interfaz do menú de produtos/servizos
"""

import gi
import sqlite3
import datetime

from reportlab.graphics.shapes import Drawing
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (SimpleDocTemplate, Spacer, Table, Paragraph)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors, styles

gi.require_version('Gtk', '3.0')
from gi.repository import Gtk


class VentaServizos(Gtk.Window):
    """
    Clase que conten as funcións do menú para xestionar os produtos e servizos
    """

    def __init__(self):
        """
        Función que se executa cada vez que se inicia a clase e inicializa os elementos contidos nela (elementos físicos,
        base de datos...)
        """
        Gtk.Window.__init__(self, title="Produtos/Servizos")

        self.datos1 = {}
        self.datos2 = {}
        self.story = []

        # dimensions da venta
        self.set_default_size(700, 500)

        # grosor do borde
        self.set_border_width(3)

        self.connFlag = True

        # conexion coa base de datos, se non existe a crea
        self.conn = sqlite3.connect('lista_clientes.db')
        self.cursor = self.conn.cursor()

        # creacion taboa produtos e insercions
        self.cursor.execute('''CREATE TABLE if not exists produtos (id integer PRIMARY KEY, nom_prod text, 
        prezo real, stock integer, color text, nvendas integer)''')

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                SELECT 1, 'mesa', '15.9', '136', '#FFB5E8', 50
                WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 1)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                        SELECT 2, 'sofa', '152.2', '4566', '#FF9CEE', 152
                        WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 2)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 3, 'lampara', '30.0', '100', '#FFCCF9', 33
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 3)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                             SELECT 4, 'estanteria', '480.4', '51', '#FCC2FF', 300
                             WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 4)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 5, 'comoda', '654.21', '15', '#F6A6FF', 41
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 5)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                             SELECT 6, 'cama', '360.0', '200', '#B28DFF', 5
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 6)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 7, 'sillon', '41.21', '24', '#C5A3FF', 214
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 7)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 8, 'hamaca', '62.0', '200', '#D5AAFF', 10
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 8)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 9, 'lavadora', '69.1', '10', '#ECD4FF', 5
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 9)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 10, 'bañeira', '480.0', '6', '#FBE4FF', 150
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 10)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 11, 'botiquin', '20.36', '40', '#DCD3FF', 87
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 11)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 12, 'grifo', '23.0', '154', '#A79AFF', 305
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 12)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 13, 'lavabo', '650.0', '52', '#B5B9FF', 65
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 13)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 14, 'fregadeiro', '98.48', '69', '#97A2FF', 25
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 14)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 15, 'vitroceramica', '458.65', '56', '#AFCBFF', 18
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 15)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 16, 'armario', '657.0', '41', '#AFF8DB', 6
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 16)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 17, 'forno', '549.9', '206', '#C4FAF8', 15
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 17)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 18, 'espello', '88.99', '48', '#85E3FF', 105
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 18)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 19, 'ducha', '20.99', '90', '#ACE7FF', 44
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 19)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 20, 'mampara', '350.3', '36', '#6EB5FF', 14
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 20)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 21, 'taladro', '56.0', '72', '#BFFCC6', 36
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 21)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 22, 'martelo', '18.34', '92', '#DBFFD6', 57
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 22)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                                    SELECT 23, 'cinta metrica', '3.58', '34', '#F3FFE3', 71
                                    WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 23)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 24, 'pegamento', '3.6', '10', '#E7FFAC', 53
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 24)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 25, 'serra', '15.15', '2', '#FFFFD1', 76
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 25)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 26, 'desaparafusador', '2.3', '61', '#FFC9DE', 465
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 26)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 27, 'cravos', '0.05', '5204', '#FFABAB', 3604
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 27)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 28, 'parafusos', '0.05', '9511', '#FFBEBC', 4980
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 28)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 29, 'mascarilla', '14.9', '304', '#FFCBC1', 165
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 29)""")

        self.cursor.execute("""INSERT INTO produtos(id, nom_prod, prezo, stock, color, nvendas) 
                            SELECT 30, 'disolvente', '18.36', '1257', '#FFF5BA', 429
                            WHERE NOT EXISTS(SELECT 1 FROM produtos WHERE id = 30)""")

        # creacion da taboa servizos e insercions
        self.cursor.execute('''CREATE TABLE if not exists servizos (id integer PRIMARY KEY, nom_serv text, 
                departamento texto, nempregados integer)''')

        self.cursor.execute("""INSERT INTO servizos(id, nom_serv, departamento, nempregados) 
                 SELECT 1, 'mantemento', 'produccion', 10 WHERE NOT EXISTS(SELECT 1 FROM servizos WHERE id = 1)""")

        self.cursor.execute("""INSERT INTO servizos(id, nom_serv, departamento, nempregados) 
                 SELECT 2, 'administracion', 'finanzas', 5 WHERE NOT EXISTS(SELECT 1 FROM servizos WHERE id = 2)""")

        self.cursor.execute("""INSERT INTO servizos(id, nom_serv, departamento, nempregados) 
        SELECT 3, 'investigacion de mercados', 'publicidade', 2 WHERE NOT EXISTS(SELECT 1 FROM servizos WHERE id = 3)""")

        self.cursor.execute("""INSERT INTO servizos(id, nom_serv, departamento, nempregados) 
                         SELECT 4, 'contabilidade', 'finanzas', 14 WHERE NOT EXISTS(SELECT 1 FROM servizos WHERE id = 4)""")

        self.cursor.execute("""INSERT INTO servizos(id, nom_serv, departamento, nempregados) 
                SELECT 5, 'programacion', 'informatica', 15 WHERE NOT EXISTS(SELECT 1 FROM servizos WHERE id = 5)""")

        self.cursor.execute("""INSERT INTO servizos(id, nom_serv, departamento, nempregados) 
                SELECT 6, 'vendas', 'publicidade', 8 WHERE NOT EXISTS(SELECT 1 FROM servizos WHERE id = 6)""")

        self.cursor.execute("""INSERT INTO servizos(id, nom_serv, departamento, nempregados) 
            SELECT 7, 'soporte tecnico', 'informatica', 10 WHERE NOT EXISTS(SELECT 1 FROM servizos WHERE id = 7)""")

        self.cursor.execute("""INSERT INTO servizos(id, nom_serv, departamento, nempregados) 
            SELECT 8, 'desarrollo', 'informatica', 5 WHERE NOT EXISTS(SELECT 1 FROM servizos WHERE id = 8)""")

        self.cursor.execute("""INSERT INTO servizos(id, nom_serv, departamento, nempregados) 
            SELECT 9, 'distribucion', 'publicidade', 8 WHERE NOT EXISTS(SELECT 1 FROM servizos WHERE id = 9)""")

        self.cursor.execute("""INSERT INTO servizos(id, nom_serv, departamento, nempregados) 
        SELECT 10, 'direccion de empresas', 'finanzas', 3 WHERE NOT EXISTS(SELECT 1 FROM servizos WHERE id = 10)""")

        self.cursor.execute("""INSERT INTO servizos(id, nom_serv, departamento, nempregados) 
             SELECT 11, 'recursos humanos', 'finanzas', 5 WHERE NOT EXISTS(SELECT 1 FROM servizos WHERE id = 11)""")

        self.conn.commit()

        self.box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)

        self.v1 = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        self.v1.pack_start(Gtk.Label("Para modificar un produto, faga dobre clic sobre a(s) columna(s) da(s) "
                                     "fila(s) que desexe editar e pulse o botón de Gardar cambios."), False, False, 0)
        self.v1.pack_start(Gtk.Label("Tamén pode ordear as columnas pulsando sobre a súa cabeceira na táboa"), False, False, 0)

        self.labelExplicacion1 = Gtk.Label()
        self.labelExplicacion1.set_markup("<b>Lista de produtos:</b>")
        self.v1.pack_start(self.labelExplicacion1, False, False, 0)

        self.modelo = Gtk.ListStore(int, str, float, int, str, int)

        self.select_produtos()

        self.treeview = Gtk.TreeView(model=self.modelo)
        self.treeview.connect("row-activated", self.on_row_activated)

        renderer_text = Gtk.CellRendererText()
        column_text = Gtk.TreeViewColumn("ID", renderer_text, text=0)
        column_text.set_sort_column_id(0)
        self.treeview.append_column(column_text)

        renderer_text = Gtk.CellRendererText()
        renderer_text.set_property('editable', True)
        renderer_text.connect("edited", self.text_edited, self.modelo, 1)
        column_text = Gtk.TreeViewColumn("Nome", renderer_text, text=1)
        column_text.set_sort_column_id(1)
        self.treeview.append_column(column_text)

        renderer_text = Gtk.CellRendererText()
        renderer_text.set_property('editable', True)
        renderer_text.connect("edited", self.text_edited, self.modelo, 2)
        column_text = Gtk.TreeViewColumn("Prezo", renderer_text, text=2)
        column_text.set_sort_column_id(2)
        # para evitar que os float se vexan asi: 1.720000000 cando o valor real é 1.72
        column_text.set_cell_data_func(renderer_text,
                                       lambda col, cell, model, iter, unused:
                                   cell.set_property("text", "%g" % self.modelo.get(iter, 2)[0]))

        self.treeview.append_column(column_text)

        renderer_text = Gtk.CellRendererText()
        renderer_text.set_property('editable', True)
        renderer_text.connect("edited", self.text_edited, self.modelo, 3)
        column_text = Gtk.TreeViewColumn("Stock", renderer_text, text=3)
        column_text.set_sort_column_id(3)
        self.treeview.append_column(column_text)

        renderer_text = Gtk.CellRendererText()
        column_text = Gtk.TreeViewColumn("Color", renderer_text, background=4)
        column_text.set_sort_column_id(4)
        self.treeview.append_column(column_text)

        renderer_text = Gtk.CellRendererText()
        renderer_text.set_property('editable', True)
        renderer_text.connect("edited", self.text_edited, self.modelo, 5)
        column_text = Gtk.TreeViewColumn("Nº de vendas", renderer_text, text=5)
        column_text.set_sort_column_id(5)
        self.treeview.append_column(column_text)

        self.tree_selection = self.treeview.get_selection()
        self.tree_selection.set_mode(Gtk.SelectionMode.MULTIPLE)

        scrollTree = Gtk.ScrolledWindow()
        scrollTree.add(self.treeview)
        self.v1.pack_start(scrollTree, True, True, 0)

        self.labelExplicacion2 = Gtk.Label()
        self.labelExplicacion2.set_markup("<b>Lista de servizos:</b>")

        self.v1.pack_start(self.labelExplicacion2, False, False, 0)

        self.modelo2 = Gtk.ListStore(int, str, str, int)

        self.select_servizos()
        self.treeview2 = Gtk.TreeView(model=self.modelo2)

        renderer_text2 = Gtk.CellRendererText()
        column_text2 = Gtk.TreeViewColumn("ID", renderer_text2, text=0)
        column_text2.set_sort_column_id(0)
        self.treeview2.append_column(column_text2)

        renderer_text2 = Gtk.CellRendererText()
        renderer_text2.set_property('editable', True)
        renderer_text2.connect("edited", self.text_edited, self.modelo2, 1)
        column_text2 = Gtk.TreeViewColumn("Nome", renderer_text2, text=1)
        column_text2.set_sort_column_id(1)
        self.treeview2.append_column(column_text2)

        renderer_text2 = Gtk.CellRendererText()
        renderer_text2.set_property('editable', True)
        renderer_text2.connect("edited", self.text_edited, self.modelo2, 2)
        column_text2 = Gtk.TreeViewColumn("Departamento", renderer_text2, text=2)
        column_text2.set_sort_column_id(2)
        self.treeview2.append_column(column_text2)

        renderer_text2 = Gtk.CellRendererText()
        renderer_text2.set_property('editable', True)
        renderer_text2.connect("edited", self.text_edited, self.modelo2, 3)
        column_text2 = Gtk.TreeViewColumn("Nº de empregados", renderer_text2, text=3)
        column_text2.set_sort_column_id(3)
        self.treeview2.append_column(column_text2)

        self.tree_selection2 = self.treeview2.get_selection()
        self.tree_selection2.set_mode(Gtk.SelectionMode.MULTIPLE)

        scrollTree2 = Gtk.ScrolledWindow()
        scrollTree2.add(self.treeview2)
        self.v1.pack_start(scrollTree2, True, True, 0)

        self.box.pack_start(self.v1, True, True, 0)

        self.v2 = Gtk.Box(spacing=6)
        self.bGardar = Gtk.Button(label="Gardar cambios")
        self.bGardar.connect("clicked", self.on_bGardar_clicked)

        self.bXerarInforme = Gtk.Button(label="Xerar informe final")
        self.bXerarInforme.connect("clicked", self.on_bXerarInforme_clicked)

        self.bVoltar = Gtk.Button(label="Voltar ao menú")
        self.bVoltar.connect("clicked", self.on_bVoltar_clicked)

        self.v2.pack_start(self.bGardar, True, True, 0)
        self.v2.pack_start(self.bXerarInforme, True, True, 0)
        self.v2.pack_start(self.bVoltar, True, True, 0)

        self.box.pack_start(self.v2, False, False, 0)

        self.add(self.box)

    def select_produtos(self):
        """
        Recolle todolos produtos contidos na taboa produtos da base de datos e os inserta no treeview

        :return:
        """
        for row in self.cursor.execute("select * from produtos"):
            self.modelo.append([row[0], row[1], row[2], row[3], row[4], row[5]])

    def select_servizos(self):
        """
        Recolle todolos produtos contidos na taboa servizos da base de datos e os inserta no treeview

        :return:
        """
        for row in self.cursor.execute("select * from servizos"):
            self.modelo2.append([row[0], row[1], row[2], row[3]])

    def check_connection(self):
        """
        Comproba a conexión da base de datos. Se non está conectada, realiza a conexión

        :return:
        """
        if not self.connFlag:
            self.conn = sqlite3.connect('lista_clientes.db')
            self.cursor = self.conn.cursor()
            self.connFlag = True

    def on_bGardar_clicked(self, widget):
        """
        Función que se executa ao facer clic no botón "gardar cambios". Recolle os campos editados no treeview e segundo
        pertenzan a un listore ou a outro os inserta nun determinado diccionario, e compara as suas claves con outros
        diccionarios equivalentes para insertalos nunha base de datos ou noutra.

        :param widget:
        :return:
        """
        model1 = {
            1: 'nom_prod',
            2: 'prezo',
            3: 'stock',
            4: 'color',
            5: 'nvendas'
        }

        model2 = {
            1: 'nom_serv',
            2: 'departamento',
            3: 'nempregados'
        }

        self.check_connection()

        if self.datos1:
            for key in model1:
                if key in self.datos1:
                    if key == 1 or key == 4:
                        query = "update produtos set %s='%s' where id=%s" % (
                            model1[key], self.datos1[key][0], self.datos1[key][1])
                        self.cursor.execute(query)

                    else:
                        query = "update produtos set %s=%s where id=%s" % (
                            model1[key], self.datos1[key][0], self.datos1[key][1])
                        self.cursor.execute(query)
            self.modelo.clear()
            self.select_produtos()

        if self.datos2:
            for key in model2:
                if key in self.datos2:
                    if key == 3:
                        query = "update servizos set %s=%s where id=%s" % (
                            model2[key], self.datos2[key][0], self.datos2[key][1])
                        self.cursor.execute(query)
                    else:
                        query = "update servizos set %s='%s' where id=%s" % (
                            model2[key], self.datos2[key][0], self.datos2[key][1])
                        self.cursor.execute(query)
            self.modelo2.clear()
            self.select_servizos()

        self.conn.commit()
        self.conn.close()
        self.connFlag = False

    def on_bXerarInforme_clicked(self, widget):
        """
        Función que se executa ao facer clic no botón "xerar informe final". Crea un arquivo en formato PDF que inserta
        as diferentes gráficas xeradas a partir da base de datos

        :param widget:
        :return:
        """
        self.now = datetime.datetime.now()

        self.check_connection()
        self.doc = SimpleDocTemplate("informe_%s-%s-%s.pdf" % (self.now.day, self.now.month, self.now.year),
                                     pagesize=A4)
        self.xerar_taboa_prods()
        self.xerar_pieChart_serv()
        self.xerar_barChart_serv()
        self.doc.build(self.story)

    def xerar_taboa_prods(self):
        """
        Xera unha táboa cos 10 produtos máis vendidos e con stock pendente de renovar cos datos obtidos da base de datos

        :return:
        """
        styles = getSampleStyleSheet()

        ptext = Paragraph('Informe de empresa con data %s-%s-%s' % (self.now.day, self.now.month, self.now.year),
                          styles['title'])

        self.story.append(ptext)

        ptext = Paragraph('Lista dos 10 produtos máis vendidos:', styles['Heading1'])

        self.story.append(ptext)

        tabp1 = Table([['ID', 'Nome', 'Nº de vendas']], colWidths=80, rowHeights=30)

        self.story.append(tabp1)

        for row in self.cursor.execute("select id, nom_prod, nvendas from produtos order by nvendas desc limit 10"):
            t = Table([
                [row[0], row[1], row[2]],
            ], colWidths=80, rowHeights=30)

            tabp1.setStyle([
                ('BOX', (0, 0), (-1, -1), 2, colors.black),
                ('BACKGROUND', (0, 0), (-1, 0), colors.pink),
            ])

            t.setStyle([
                ('TEXTCOLOR', (0, 0), (0, -1), colors.purple),
                ('TEXTCOLOR', (1, 0), (1, -1), colors.darkblue),
                ('BACKGROUND', (0, 0), (-1, -1), colors.white),
                ('BOX', (0, 0), (-1, -1), 2, colors.black)
            ])
            self.story.append(t)
        self.story.append(Spacer(0, 15))

        ptext = Paragraph('Produtos con stock pendente de renovar:', styles['Heading1'])

        self.story.append(ptext)

        tabp2 = Table([['ID', 'Nome', 'Stock']], colWidths=80, rowHeights=30)

        self.story.append(tabp2)

        for row in self.cursor.execute("select id, nom_prod, stock from produtos where stock < 16"):
            t = Table([
                [row[0], row[1], row[2]],
            ], colWidths=80, rowHeights=30)

            tabp2.setStyle([
                ('BOX', (0, 0), (-1, -1), 2, colors.black),
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgreen),
            ])

            t.setStyle([
                ('TEXTCOLOR', (0, 0), (0, -1), colors.purple),
                ('TEXTCOLOR', (1, 0), (1, -1), colors.darkblue),
                ('BACKGROUND', (0, 0), (-1, -1), colors.white),
                ('BOX', (0, 0), (-1, -1), 2, colors.black)
            ])
            self.story.append(t)
        self.story.append(Spacer(0, 55))

    def xerar_pieChart_serv(self):
        """
        Xera unha gráfica circular co número de empregados por servizo, obtidos da base de datos

        :return:
        """
        from reportlab.graphics.charts.piecharts import Pie
        from reportlab.graphics.charts.legends import Legend

        styles = getSampleStyleSheet()

        ptext = Paragraph('Número de empregados por tipo de servizo:', styles['Heading1'])

        self.story.append(ptext)

        d = Drawing(300, 200)
        pc = Pie()
        pc.x = 65
        pc.y = 12
        pc.width = 170
        pc.height = 170

        pc.labels = ['Mantemento', 'Administración', 'Investigación de mercados', 'Contabilidade', 'Programación',
                     'Vendas', 'Soporte técnico', 'Desarrollo', 'Distribución', 'Dirección de empresas',
                     'Recursos humanos']

        pc.data = [row for row in self.cursor.execute("""SELECT (SELECT nempregados FROM servizos WHERE (nom_serv = 'mantemento')) AS man,
        (SELECT nempregados FROM servizos WHERE (nom_serv = 'administracion')) AS adm,
        (SELECT nempregados FROM servizos WHERE (nom_serv = 'investigacion de mercados')) AS inv,
        (SELECT nempregados FROM servizos WHERE (nom_serv = 'contabilidade')) AS cont,
        (SELECT nempregados FROM servizos WHERE (nom_serv = 'programacion')) AS prog,
        (SELECT nempregados FROM servizos WHERE (nom_serv = 'vendas')) AS vendas,
        (SELECT nempregados FROM servizos WHERE (nom_serv = 'soporte tecnico')) AS sop,
        (SELECT nempregados FROM servizos WHERE (nom_serv = 'desarrollo')) AS des,
        (SELECT nempregados FROM servizos WHERE (nom_serv = 'distribucion')) AS dist,
        (SELECT nempregados FROM servizos WHERE (nom_serv = 'direccion de empresas')) AS dir,
        (SELECT nempregados FROM servizos WHERE (nom_serv = 'recursos humanos')) FROM servizos AS rh
        """).fetchone()]

        pc.slices.strokeWidth = 0.5

        pc.slices[2].popout = 10
        pc.slices[2].strokeWidth = 2
        pc.slices[2].strokeDashArray = [2, 2]
        pc.slices[2].labelRadius = 1.75
        pc.slices[2].fontColor = colors.purple

        pc.sideLabels = 1  # Con 0 no se muestran líneas hacia las etiquetas
        # Insertamos la leyenda

        legend = Legend()
        legend.x = 370
        legend.y = 0
        legend.dx = 8
        legend.dy = 8
        legend.fontName = 'Helvetica'
        legend.fontSize = 7
        legend.boxAnchor = 'n'
        legend.columnMaximum = 10
        legend.strokeWidth = 1
        legend.strokeColor = colors.black
        legend.deltax = 75
        legend.deltay = 10
        legend.autoXPadding = 5
        legend.yGap = 0
        legend.dxTextSpace = 5
        legend.alignment = 'right'
        legend.dividerLines = 1 | 2 | 4
        legend.dividerOffsY = 4.5
        legend.subCols.rpad = 30

        # Insertemos nuestros propios colores
        colores = [colors.darkblue, colors.cyan, colors.purple, colors.magenta, colors.pink, colors.lightgreen,
                   colors.coral, colors.crimson,
                   colors.darkgreen, colors.darkorange, colors.lavender]
        for i, color in enumerate(colores):
            pc.slices[i].fillColor = color

        legend.colorNamePairs = [(
            pc.slices[i].fillColor,
            (pc.labels[i][0:20], '%0.2f' % pc.data[i])
        ) for i in range(len(pc.data))]

        d.add(pc)
        d.add(legend)
        self.story.append(d)
        self.story.append(Spacer(0, 110))
        # self.doc.build(self.story) # descomentar se so queremos xerar esta grafica

    def xerar_barChart_serv(self):
        """
        Xera un gráfico de barras co número de empregados por departamento nos últimos 8 meses a partir da base de datos

        :return:
        """
        from reportlab.graphics.charts.barcharts import VerticalBarChart

        styles = getSampleStyleSheet()

        ptext = Paragraph('Número de empregados por departamento nos últimos 8 meses:', styles['Heading1'])

        self.story.append(ptext)

        d = Drawing(400, 200)

        p = self.cursor.execute("""SELECT (SELECT SUM(nempregados) FROM servizos WHERE (departamento = 'produccion')) AS prod, 
        (SELECT SUM(nempregados) FROM servizos WHERE (departamento = 'finanzas')) AS finanzas,
        (SELECT SUM(nempregados) FROM servizos WHERE (departamento = 'informatica')) AS informatica,
         (SELECT SUM(nempregados) FROM servizos WHERE (departamento = 'publicidade')) FROM servizos AS publicidade
        """).fetchone()

        self.conn.text_factory = str

        data = [
            (25, 34, 14, 18, 36, 22, 19, p[0]),
            (26, 31, 22, 24, 16, 10, 18, p[1]),
            (34, 40, 18, 9, 25, 7, 2, p[2]),
            (29, 20, 19, 16, 34, 40, 15, p[3])
        ]

        bc = VerticalBarChart()
        bc.x = 50
        bc.y = 55
        bc.height = 125
        bc.width = 300
        bc.data = data
        bc.strokeColor = colors.black
        bc.valueAxis.valueMin = 0
        bc.valueAxis.valueMax = 50
        bc.valueAxis.valueStep = 10  # paso de distancia entre punto y punto
        bc.categoryAxis.labels.boxAnchor = 'ne'
        bc.categoryAxis.labels.dx = 8
        bc.categoryAxis.labels.dy = -2
        bc.categoryAxis.labels.angle = 30
        bc.categoryAxis.categoryNames = ['Xan-19', 'Feb-19', 'Mar-19',
                                         'Abr-19', 'Mai-19', 'Xuñ-19',
                                         'Xul-19', 'Ago-19']
        bc.groupSpacing = 10
        bc.barSpacing = 2
        # bc.categoryAxis.style = 'stacked'  # Una variación del gráfico
        d.add(bc)
        self.story.append(d)

    def on_bVoltar_clicked(self, widget):
        """
        Función que se executa ao pulsar o botón de "Voltar ao menú". Se a base de datos está conectada, a desconecta e
        pecha a ventá actual.

        :param widget:
        :return:
        """
        if self.connFlag:
            self.conn.close()
        self.destroy()

    def text_edited(self, w, row, new_text, model, column):
        """
        Execútase cada vez que se edita unha fila do treeview, para que non se eliminen os cambios cada vez que cambiamos
        de fila. Como a listore contén tipos de datos distintos, fai un casting para que non salte unha excepción de tipo
        int, string... Se o listore editado é dun tipo concreto almacenase nun diccionario ou noutro para a súa posterior
        manipulación.

        :param w:
        :param row: fila editada
        :param new_text: texto editado
        :param model: listore editado
        :param column: columna editada
        :return:
        """
        tipo = type(model[row][column])
        try:
            if isinstance(model[row][column], tipo):
                model[row][column] = tipo(new_text)
        except ValueError:
            """Tipo de valor invalido"""

        if model == self.modelo:
            self.datos1[column] = [model[row][column], model[row][0]]
        else:
            self.datos2[column] = [model[row][column], model[row][0]]

    def color_activated(self):
        """
        Método que se executa ao facer doble nunha columna non editable, como o id, ou o color. Neste caso cambia
        o cor mostrado na táboa na base de datos, abrindo un cadro de diálogo para escoller un cor e devolve o seu
        código hexadecimal

        :return:
        """
        color = self.colorchooserdialog.get_rgba()

        red = int(color.red * 255)
        green = int(color.green * 255)
        blue = int(color.blue * 255)

        return '#%02x%02x%02x' % (red, green, blue)

    def on_row_activated(self, widget, row, col):
        """
        Método que modifica o cor seleccionado polo usuario na base de datos

        :param widget:
        :param row:
        :param col:
        :return:
        """
        model = widget.get_model()

        self.colorchooserdialog = Gtk.ColorChooserDialog(parent=self)

        if self.colorchooserdialog.run() == Gtk.ResponseType.OK:
            model[row][4] = str(self.color_activated())
            self.datos1[4] = [model[row][4], model[row][0]]
        self.colorchooserdialog.destroy()
