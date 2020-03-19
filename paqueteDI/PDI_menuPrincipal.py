"""
PDI_menuPrincipal.py
====================================
Arquivo que contén a interfaz do menú principal
"""

import gi

gi.require_version('Gtk', '3.0')
from gi.repository import Gtk


class Venta(Gtk.Window):
    """
    Clase que conten as funcións do menú principal e os seus elementos
    """
    def __init__(self):
        """
        Función que contén os elementos do menú principal: botóns, etiquetas...
        """
        Gtk.Window.__init__(self, title="Menú principal")

        self.set_default_size(500, 200)

        self.box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=6)
        self.add(self.box)

        self.labelBenvida = Gtk.Label()
        self.labelBenvida.set_markup("<b>Benvido/a:</b>")
        self.box.pack_start(self.labelBenvida, True, True, 0)

        self.bXestion = Gtk.Button(label="Xestión de clientes")
        self.bXestion.connect("clicked", self.on_bXestion_clicked)
        self.box.pack_start(self.bXestion, True, True, 0)

        self.bProdutos = Gtk.Button(label="Produtos/Servizos")
        self.bProdutos.connect("clicked", self.on_bProdutos_clicked)
        self.box.pack_start(self.bProdutos, True, True, 0)

        self.bSair = Gtk.Button(label="Saír")
        self.bSair.connect("clicked", self.on_bSair_clicked)
        self.box.pack_start(self.bSair, True, True, 0)

    def on_bXestion_clicked(self, widget):
        """
        Función que se executa ao facer clic no botón "Xestión de clientes". Importa outro arquivo e crea unha
        ventá a partir deste e conecta a esta ventá nova o eventos de peche para que volva a abrir o menú principal

        :param widget: manexa o control de eventos ao facer clic no boton
        :return:
        """
        from paqueteDI.PDI_xestionClientes import VentaClientes
        global win
        vclientes = VentaClientes()
        vclientes.connect('delete-event', self.on_destroy)
        vclientes.connect('destroy', self.on_destroy)
        vclientes.show_all()
        win.hide()

    def on_bProdutos_clicked(self, widget):
        """
        Función que se executa ao facer clic no botón "Produtos/Servizos". Importa o arquivo da ventá de servizos
        e oculta o menú principal. Se se pecha a ventá de servizos, volve a mostrar a principal

        :param widget: manexa o control de eventos ao facer clic no boton
        :return:
        """
        from paqueteDI.PDI_produtosServizos import VentaServizos
        global win
        vservizos = VentaServizos()
        vservizos.connect('delete-event', self.on_destroy)
        vservizos.connect('destroy', self.on_destroy)
        vservizos.show_all()
        win.hide()

    def on_bSair_clicked(self, widget):
        """
        Función que se executa ao facer clic no botón "Saír". Pecha a ventá do menú principal e finaliza o programa

        :param widget: manexa o control de eventos ao facer clic no boton
        :return:
        """
        Gtk.main_quit()

    def on_destroy(self, widget=None, *data):
        """
        Función que se executa cada vez que se pechen as ventás secundarias conectadas a este evento. Mostra o menú
        principal

        :param widget:
        :param data:
        :return:
        """
        global win
        win.show()


def main(args=None):
    """
    Función que executará a ventá principal para iniciar o programa. Se emprega para iniciar o programa desde a consola.

    :param args: parametro que ten que estar baleiro para poder usalo como metodo a executar no setup.py
    :return:
    """
    global win
    win.connect("destroy", Gtk.main_quit)
    win.show_all()
    Gtk.main()


win = Venta()

if __name__ == '__main__':
    main()
