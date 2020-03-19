from setuptools import setup

descripcion_longa = open('readme.rst').read()

setup(
    name="paqueteDI",
    version="0.12",
    author="Alba",
    author_email="aperezcesar@danielcastelao.org",
    url="https://www.danielcastelao.org",
    license="GLP",
    platforms="Unix",
    clasifiers=["Development Status :: 3 - Alpha",
                "Environment :: Console",
                "Topic :: Software Development :: Libraries",
                "License :: OSI Approved :: GNU General Public License",
                "Programming Language :: Python :: 3.6",
                "Operating System :: Linux Ubuntu"

                ],
    description="Aplicación de xestión de empresa",
    long_description=descripcion_longa,
    keywords="empaquetado instalador paquetes",
    packages=['paqueteDI', 'paqueteDI.docs'],  # packages = find_packages(exclude=['*.test', '*.test.*'])
    package_data={
        'paqueteDI': ['notas.txt']
    },
    # data_files=[('datos', ['dat/datos.txt'])],
    entry_points={'console_scripts': ['miProyecto = paqueteDI.PDI_menuPrincipal:main', ], }
)
