
Para usar las herramientas homologa y homologa2 se siguen los siguientes pasos:

1. Ingresas al capp del estudiante
2. Guardas (click derecho, guardar como), en la carpeta donde está el script, la página web completa del capp usando el nombre capp.html. 
3. Editas el archivo de python para incluir: nombre, id y línea del estudiante.
4. Corres el archivo python homologa.py u homologa2.py y el te genera un excel con el plan de estudios homologado. Tanto homologa como homologa2 dibujan 
   el plan de estudios;sin embargo, homologa2 muestra el capp en consola coloreando los cursos que no están homologados, no homologados con créditos a favor, 
   homologados con menos créditos y homologados con más créditos.

homologa3 te muestra la misma información que homologa2 más la información 
adicional.

Para usar la herramienta homologa3 sigue los siguientes pasos:

1. Ingresas al capp del estudiante
2. Guardas (click derecho, guardar como), en la carpeta donde está el script, la página web completa del capp usando el nombre capp.html. 
3. Guarda la página Información adicional con el nombre infoAdd.html
4. Editas el archivo de homologa3 para incluir: nombre, id y línea del estudiante.

homologaAll permite dibujar los planes de estudio homologados usando como entrada una lista de estudiantes almancenada en excel. No es necesario descargar manualmente el capp de cada estudiante porque homologaAll lo hace automáticamente. Para usar la herramienta homologaAll necesitas:

1. Descargar un driver para GoogleChrome: https://sites.google.com/chromium.org/driver/downloads
2. Alimenta el archivo listado con la información de tus estudiantes
3. Ingresa tus credenciales en el archivo sigaa.json
4. Ejecuta $ python homologaAll.py
5. Observa la terminal para verificar si la línea de énfasis seleccionada por el estudiante concuerda con la línea de énfasis en el capp. En caso contrario verás
   que un mensaje con el ID del estudiante y la indicación que su curso no está en el capp.

   

