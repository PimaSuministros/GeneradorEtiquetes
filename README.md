# Generador d'etiquetes
Ferramenta per a generar un document .odt per a imprimir sobre una fulla A4 d'etiquetes de 70x30mm. 

## Fulla d'etiquetes
Q-CONNECT KF10642 70x30mm

## Configurar sistema
- Instal·lar virtualenv si no ho està
 	- `pip3 install virtualenv`
- Crear un nou entorn virtual
	- `virtualenv python-env`
- Entrar a l'entorn virtual 
	- `python-env\Scripts\activate` (Windows)
	- `source python-env/bin/activate` (GNU/Linux)
- Instal·lar dependències
	- `pip install python-docx` (Window i GNU/Linux)
	- `apt-get install python3-tk` (GNU/Linux)
