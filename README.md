# Hub de Lectura

Este proyecto fue desarrollado utilizando Visual Basic 6.0 y SQL Server 2019. Es una aplicación de escritorio que permite a los usuarios gestionar un catálogo de libros, marcarlos como leídos, guardarlos como favoritos, entre otras funcionalidades.

## Descripción

El Hub de Lectura es una aplicación diseñada para organizar y administrar un catálogo personal de libros. Permite a los usuarios realizar diferentes acciones sobre sus libros, como marcarlos como leídos, agregar a favoritos, libros que no les gustaron, y mucho más.

## Características

- **Catálogo de Libros**: Visualiza todos los libros en tu catálogo y agrega nuevas lecturas.
- **Marcar como Leído**: Registra los libros que ya has leído.
- **Libros Favoritos**: Guarda tus libros favoritos para fácil acceso.
- **Libros Que No Te Gustaron**: Registra los libros que no te gustaron y sepáralos del resto.
- **Libros Que Quieres Leer**: Organiza una lista de libros que te gustaría leer en el futuro.
- **Generos Favoritos**: Muestra los géneros favoritos basado en los libros que has leído.
- **Modo Oscuro/Claro**: Cambia entre modos oscuro y claro para una mejor experiencia visual.

## Tecnologías Utilizadas

- **Visual Basic 6.0**
- **SQL Server 2019**
- **ADO (ActiveX Data Objects)**

## Instalación

1. Clona el repositorio:
    ```sh
    git clone https://github.com/tu-usuario/hub-de-lectura.git
    ```
2. Abre el proyecto en Visual Basic 6.0.
3. Configura la cadena de conexión en el módulo `modConexion.bas` para que apunte a tu instancia de SQL Server.
4. Ejecuta los scripts SQL proporcionados para crear la base de datos y las tablas necesarias.
5. Compila y ejecuta el proyecto desde Visual Basic.

## Uso

1. Al iniciar la aplicación, se te presentará el menú principal con diferentes opciones para acceder a las distintas funcionalidades.
2. Navega por las opciones para ver el catálogo de libros, marcar libros como leídos, recomendados, entre otros.

![Pantalla Principal](./imagenes/PantallaPrincipal.png)

3. Selecciona cualquier libro realizar acciones como marcar como leído, agregar a favoritos, etc.

![Detalles del Libro](./imagenes/detalles.png)

4. Utiliza el botón de modo oscuro/claro para cambiar el esquema de colores de la aplicación.

![Modo Oscuro](./imagenes/sol.png)

## Proceso que seguí para hacerlo

Primero, se desarrolló la estructura básica de la base de datos utilizando SQL Server. Se crearon las tablas necesarias para almacenar los libros, géneros, y los estados de los libros como leídos o favoritos.

Luego, se comenzó con el diseño de la interfaz de usuario en Visual Basic 6.0, implementando formularios para cada sección de la aplicación. Se usó ADO para conectar la aplicación con la base de datos y realizar operaciones CRUD (Crear, Leer, Actualizar, Borrar).

Finalmente, se añadieron funcionalidades adicionales como el modo oscuro/claro, y se realizaron pruebas para asegurar el correcto funcionamiento de la aplicación.

## Diagrama de Entidad-Relación de la Base de Datos
![Diagrama de ER](./imagenes/diagrama.png)

## Problemas Conocidos

- **Responsividad**: La interfaz de usuario no es completamente responsiva.

## Retrospectiva

| ¿Qué salió bien? | ¿Qué puedo hacer diferente? | ¿Qué no salió bien? |
|------------------|-----------------------------|---------------------|
| Se logró implementar la funcionalidad principal del Hub de Lectura, incluyendo la gestión de libros leídos, recomendados y géneros favoritos. | Mejorar el diseño de la interfaz gráfica para hacerlo más atractivo y moderno. | La implementación del perfil de usuario con preferencias quedó limitada al cambio entre modo claro y oscuro, sin más detalles. |
| La integración con SQL Server para la gestión de la base de datos fue exitosa y eficiente. | Dedicar más tiempo a la planificación del diseño de la interfaz, considerando aspectos como la organización visual y la experiencia del usuario. | La sección de "Perfil de Usuario" no se desarrolló de manera completa y puede mejorarse en el futuro. |
| Se añadió un modo oscuro y claro que mejora la accesibilidad visual. | Considerar mejoras en la estructura del código para facilitar futuras expansiones y mantenimiento del proyecto. | La gestión de la responsividad al reorganizar ventanas no se abordó de manera completa y puede ser mejorada en futuros desarrollos. |