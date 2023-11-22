latex-to-excel

condiciones de uso:
- es importante que en los latex utilizados se siga siempre el mismo formato:
	\begin{key} <letra>
		...
	Eje Temático: <texto>\\
	Contenido: <texto>\\
	Habilidad: <texto>\\
	Dificultad: <texto>\\
		...
  de no ser asi es probable que hayan errores o no se ejecute ya que no detectara correctamente el texto a buscar.

notas:
- las columnas en blanco las mantuve para que permanezca el formato,
  si quieres puedo añadir una forma de pasar el contenido del latex a esas columnas.
- en caso de que algun dato de las columnas "clave", "eje", "contenido", "habilidad" y "dificultad" este en blanco 
  revisa el latex primero.
- segun el formato, en caso de que solo <texto> este vacío, ej:

	"Eje Temático: \\"

  la casilla correspondiente en el excel estará en blanco.