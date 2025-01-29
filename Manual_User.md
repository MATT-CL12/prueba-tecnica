# Manual de Usuario para el Sistema de Ingreso de Datos


Este sistema permite automatizar el ingreso de datos en una plataforma web a partir de un archivo Excel. El código está diseñado para procesar la información contenida en un archivo de entrada, estructurado con varias hojas que representan diferentes categorías de datos (como OTs, Tareas, Materiales, etc.). Este manual describe cómo preparar los datos y ejecutar el sistema correctamente.

---
![Captura de pantalla](https://github.com/Zarcasmo/Web_Scraping_MX/raw/Develop_2/img/1.png)


Para iniciar debemos instalar el programa llamado ANACONDA NAVIGATOR, 
este deberá ser instalado con permisos de administrado. favor llamar a soporte 
Tecnico para la instalación.

---

![Captura de pantalla](https://github.com/Zarcasmo/Web_Scraping_MX/raw/Develop_2/img/2.png)

Luego ser instalado el programa debemos ingresar al apartado Environments y 
luego create.

---

![Captura de pantalla](https://github.com/Zarcasmo/Web_Scraping_MX/raw/Develop_2/img/3.png)

En el Name debemos colocar el nombre Maximo para fácil ubicación, luego 
asegurarse de marcar la casilla de python y por último asegurarse de escoger 
la version correcta (3.9.21) si no se selecciona la correcta el sistema no 
funcionara. por último, le damos a create. esto puede tardar un poco favor 
tener paciencia hasta que se instale correctamente, puede verificar que se está 
instalando en la parte inferior derecha de pa pantalla de ANACONDA. 


---

![Captura de pantalla](https://github.com/Zarcasmo/Web_Scraping_MX/raw/Develop_2/img/4.png)

Luego de terminar la instalación de python debemos volver al home eh instalar 
el programa spyder, este se instalará adecuadamente solo con darle clic en 
install. cuando este esté instalado le damos a launch para iniciarlo.


---

![Captura de pantalla](https://github.com/Zarcasmo/Web_Scraping_MX/raw/Develop_2/img/5.png)

Una vez iniciado el spyder debemos seleccionar el botón de archivos, luego el 
de abrir, luego buscaremos el archivo donde tenemos el bot guardado.


---

![Captura de pantalla](https://github.com/Zarcasmo/Web_Scraping_MX/raw/Develop_2/img/6.png)

Una vez hallamos encontrado el archivo del bot seleccionaremos el documento 
llamado Main_Cargar_Servicios.

---

![Captura de pantalla](https://github.com/Zarcasmo/Web_Scraping_MX/raw/Develop_2/img/7.png)

Luego de seleccionar el main debemos ir al aparatado de la derecha y 
seleccionar archivos, este nos desplegará todos los documentos que tenemos 
en la carpeta del bot, en caso de no visualizar la carpeta también la podemos 
buscar desde el botón de archivos en la parte superior derecha (es el que 
parece un símbolo de archivos justo a la izquierda de la flecha blanca que 
apunta hacia arriba) una vez aquí buscaremos la carpeta llamada Input, en esta 
se encuentra un archivo llamado User_Config.json lo seleccionamos y nos 
aparecerá una información en la parte izquierda de la pantalla, debemos buscar 
el apartado donde dice USER y el apartado onde dice PASS, estos tiene a su 
derecha un texto entre comillas, en este caso debe de decir usuario y 
contraseña. en estos apartados debemos borrar usuario y colocar su usuario 
de máximo, de igual forma con la contraseña. cabe aclarar que se debe hacer 
dentro de las comillas sin borrarlas. luego de digitar los datos debemos 
ingresar (control + s) para guardar el cambio o en su defecto seleccionar el 
botón de guardado que se encuentra justo debajo del apartado de buscar (si 
tiene dudas de si es el botón correcto con solo dejar el mouse encima del 
botón sin hacer nada más le dirá la función de tal botón)


---

![Captura de pantalla](https://github.com/Zarcasmo/Web_Scraping_MX/raw/Develop_2/img/8.png)

Luego debemos ingresar al apartado Herramientas y luego el de preferencias. 
tal y como se muestra en la imagen.


---

![Captura de pantalla](https://github.com/Zarcasmo/Web_Scraping_MX/raw/Develop_2/img/9.png)

En este apartado debemos de prestar mucha atención y si es necesario 
realizarlo varias veces para asegurarse de que los cambios se implementen 
correctamente.  debemos ingresar al apartado interprete de python, luego 
seleccionamos la opción de (Usar el siguiente interprete) luego de esto 
buscaremos el intérprete anteriormente creado como máximo o como usted 
haya colocado el nombre. en siguiente paso le damos aplicar y luego aceptar, 
por favor volver a ingresar para verificar que el intérprete que salga 
seleccionado sea máximo si no el bot no funcionara.


---

![Captura de pantalla](https://github.com/Zarcasmo/Web_Scraping_MX/raw/Develop_2/img/10.png)

En el siguiente paso debemos de ingresar a la terminal esta se encuentra en la 
parte inferior derecha, verificar que en los botones de abajo este seleccionado 
terminal y no historial, luego de verificar le damos clic dentro de la terminal 
aquí le permitirá escribir así que debemos escribir el siguiente comando: ```  pip 
install pandas==1.3.5 ```   y le das Enter tal y como se ve en la imagen van a 
aparecer un datos de información, no prestar atención solo revisar cuando ya 
te permita escribir de nuevo para digitar el siguiente comando: ```  pip install 
selenium==4.27.1 ```  esperar a que este comando instale también, esto tardara 
un poco.


---

![Captura de pantalla](https://github.com/Zarcasmo/Web_Scraping_MX/raw/Develop_2/img/11.png)

Verificar que el archivo que almacena los datos este en la carpeta input y con 
el nombre como aparece en la foto Datos_entrada


---

![Captura de pantalla](https://github.com/Zarcasmo/Web_Scraping_MX/raw/Develop_2/img/12.png)

Dentro del excel podemos encontrar una tabla donde demos diligenciar los 
datos para las OT, se puede llenar todos los datos que dese en vertical. tener 
en cuenta 2 cosas:
1. solo llenar los datos de las casillas azules ya que las filas que están en 
blanco se llenaran automáticamente, estas no se deben de tocar.
2. los datos que se encuentran señalados con flechas son las acciones que 
desea hacer en la OT, tener en cuenta que estas casillas solo se pueden 
llenar en ```  0  ``` o ```  1  ``` el bot tomara un 1 como activación ósea que debe 
realizar esa acción y cuando está en cero es que no debe realizar esta 
acción. estas casillas se deben llenar según no requerido. 


---

![Captura de pantalla](https://github.com/Zarcasmo/Web_Scraping_MX/raw/Develop_2/img/13.png)

Por último cada que dese activar el bot luego de digitar y guardar los datos en 
el excel. se debe abrir el spyder fijarse que ya este seleccionado el archivo 
main y darle play en el botón verde que esta seleccionado en la imagen.


---
# Importante

Una vez este el bot en funcionamiento evitar ingresar manualmente a Maximo y 
mucho menos modificar las ventanas que se abren automáticamente.

Si desea verificar el proceso que tuvo el bot en diligenciar los datos o buscar 
algún error al digitar la información esta se podrá encontrar en la carpeta 
llamada Output (este se encuentra dentro de la carpeta del bot) esta está 
dividida en carpetas llamadas report y debug el debug almacena un excel con 
todos los datos digitados y el report almacena un archivo que puedes ver todo 
el proceso del bot.
