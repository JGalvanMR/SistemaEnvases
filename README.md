# 📦 Nombre del Módulo: SistemaEnvases

## 🧭 Propósito

SistemaEnvases es una aplicación de escritorio Windows Forms (.NET Framework 4.0) para Comercializadora GAB, S.A. de C.V., que gestiona el inventario diario de envases (cajones, canastillas de espárrago, etc.) utilizados en operaciones agrícolas. Controla entradas y salidas de envases por proveedor, rancho, tabla y producto, realiza cortes diarios de inventario, genera reportes (kardex, inventario por proveedor, inventario físico) y exporta dicha información a Excel.

## ⚙️ Responsabilidades

- Registrar **salidas de envases** (Form1, pestaña SALIDAS) asociadas a proveedor, rancho, tabla, producto y envase, con folio autogenerado.
- Registrar **entradas de envases** (Form1, pestaña ENTRADAS) con la misma estructura de datos.
- Permitir **cancelación** de folios de entrada/salida, revirtiendo los movimientos asociados.
- Calcular y mantener un **inventario diario** (`TB_MSTR_INV_ENVASES_dos`, `TB_MSTR_INV_ENVASES_sin_corte`) con inventario inicial, entradas, salidas e inventario final.
- Ejecutar un **corte diario automático** (BtnActxDia / bgCorteDiario) que calcula inventarios iniciales del día siguiente a partir de entradas/salidas del día anterior, para tres ámbitos: envases generales, cajones por proveedor (claves "02" y "80") y canastillas de espárrago por proveedor (clave "81").
- Permitir **modificación manual del inventario inicial** (pestaña INVENTARIO INICIAL, INVENTARIO CAJONES, INVENTARIO ESPARRAGO) mediante BackgroundWorkers que recalculan en cascada los días posteriores hasta la fecha actual.
- Generar **reportes de inventario** por proveedor, por envase, por cajones y por espárrago (CAJONES, ESPARRAGO, Form2/DetalleEnvase) con exportación a Excel vía interoperabilidad COM.
- Generar **reporte Kardex** (entradas y salidas combinadas) por rango de fechas y exportarlo a Excel.
- Gestionar **inventario físico de canastillas** (pestaña CAPTURA INVENTARIO): captura manual de cantidades físicas, validación de cuadre contra categorías (en campo, con producto, vacías, con basura, por reparar, envase roto) y envío de correo con resumen.
- **Auto-registrar proveedores nuevos** de espárrago (`Tb_ENV_PROV_CAJ_ESPARRAGO`, `TB_MSTR_INV_CAJAS_PROV_ESPARRAGO`) cuando un proveedor mueve el envase clave "81" por primera vez (excepto proveedores exentos).
- Verificar **actualizaciones de versión del ejecutable** comparando fechas de archivo local vs. servidor, y disparar un actualizador externo (`DownFile.exe`) al cerrar la aplicación si corresponde.
- Validar **sesión de usuario** activa al iniciar (`Program.cs`) y registrar/cerrar sesiones en `tb_cat_historial_dia`.

## 🔄 Flujo de Funcionamiento

1. **Inicio de la aplicación** (`Program.Main`): valida IP (`Utilerias.Class1.validar_ip`), valida que exista una sesión activa del usuario en `tb_cat_historial_dia` para el equipo actual; si no existe, muestra aviso y lanza `SIPGAB.exe`. Si la sesión es válida, abre `Form1`.

2. **Carga de Form1** (`Form1_Load`):
   - Obtiene la fecha del servidor (`SYSDATETIME()`).
   - Verifica si hay una versión más reciente del ejecutable (`ValidaActualizacion`); si la hay, cierra el formulario para forzar actualización.
   - Inicializa estructuras `DataTable` para inventarios, reportes de entradas/salidas, kardex e inventario físico.
   - Carga el inventario inicial del día (`TB_MSTR_INV_ENVASES_dos`) en `dataGridInfoInvenario` y `DGInvFis`.
   - Carga inventario inicial de cajones por proveedor (`TB_MSTR_INV_CAJAS_PROV`) y de espárrago por proveedor (`TB_MSTR_INV_CAJAS_PROV_ESPARRAGO`).
   - Si el usuario no está en la lista `usuariostotales`, oculta pestañas administrativas (INVENTARIO INICIAL, INVENTARIO CAJONES, INVENTARIO ESPARRAGO) y deshabilita edición de fechas.

3. **Registro de salida** (`BtnSplitPed_Click`):
   - Valida proveedor, chofer, operador y al menos un envase en la grilla.
   - Abre transacción SQL, inserta encabezado en `TB_SALIDAS_ENVASES` (parametrizado, captura folio con `SCOPE_IDENTITY()`).
   - Por cada envase: inserta detalle en `TB_DETSALIDAS_ENVASES`, actualiza `TB_MSTR_ENVASES`, `TB_MSTR_INV_ENVASES`, `TB_MSTR_INV_ENVASES_dos` y `TB_MSTR_INV_ENVASES_sin_corte`.
   - Si el envase es clave "81" (espárrago), llama `validarproveedoresparragoT` dentro de la misma transacción.
   - Registra movimiento en `tb_registro_movimientos` (tipo SALIDA o ATRASADO según fecha).
   - Si la fecha es anterior a hoy, ejecuta `realizar_corte` para recalcular inventarios desde esa fecha hasta hoy.
   - Confirma transacción, imprime ticket y limpia el formulario.

4. **Registro de entrada** (`guardar_ent_Click`): flujo análogo al de salida pero sin uso de transacción explícita (consultas individuales con concatenación de cadenas SQL), actualizando `ENV_ENTR_CANT` en lugar de `ENV_SAL_CANT`.

5. **Corte diario** (`bg_DoWorkcorte`, disparado por `button18_Click`):
   - Verifica si ya existe un registro de corte para el día (`tb_registro_movimientos` con detalle específico).
   - Si no existe, para cada envase con `env_inventario = '1'`: calcula inventario inicial del día anterior + entradas − salidas, inserta nuevo registro en `TB_MSTR_INV_ENVASES_dos` y `TB_MSTR_INV_ENVASES_sin_corte` para el día actual.
   - Repite proceso análogo para cajones por proveedor (envases "02" y "80") sobre `TB_MSTR_INV_CAJAS_PROV`.
   - Repite proceso análogo para espárrago por proveedor (envase "81") sobre `TB_MSTR_INV_CAJAS_PROV_ESPARRAGO`, **invirtiendo la fórmula** (inventario = inicial + salidas − entradas).
   - Cada sección registra su propio marcador en `tb_registro_movimientos` para evitar reprocesos.

6. **Recalculo manual de inventario inicial** (`bg_DoWork`, `bgcajones_DoWork`, `bgESPARRAGO_DoWork`):
   - Toma la fecha seleccionada en el DateTimePicker correspondiente.
   - Si el usuario modificó manualmente "CANTIDAD MODIFICADA", actualiza `ENV_INV_INI_CANT` para esa fecha.
   - Itera día por día desde la fecha seleccionada hasta hoy, recalculando entradas, salidas e inventario inicial del día siguiente, reportando progreso a una barra de progreso.

7. **Reportes** (botones 7, 10, 11, 14, 15, 16, 17 en pestaña REPORTES):
   - Consultan tablas de inventario/movimientos según filtros de fecha, proveedor o envase.
   - Construyen `DataTable` de resultados, los muestran en formularios secundarios (CAJONES, ESPARRAGO, DetalleEnvase, Form3, Form4, detalle_envase).
   - Exportan a Excel mediante `Microsoft.Office.Interop.Excel`, aplicando formato condicional (colores por fila: cancelado, total, etc.).

8. **Inventario físico** (pestaña CAPTURA INVENTARIO):
   - Carga `TB_MSTR_INV_ENVASES_dos` del día seleccionado (`DtInvFis_ValueChanged`) en `DGInvFis`.
   - El usuario captura cantidades por categoría (en campo, con producto, vacías, con basura, por reparar, roto) y observaciones.
   - Al editar celdas (`DGInvFis_CellEndEdit`), se recalcula la suma total.
   - Al guardar (`BtnSavFis_Click`): valida que la suma de categorías coincida con el inventario físico declarado por envase (`Acumula`); si es correcto, actualiza `TB_MSTR_INV_ENVASES_dos` y envía correo HTML con el detalle (`SendMail`).

9. **Cierre de la aplicación** (`Form1_FormClosed`): si se detectó actualización pendiente, lanza `c:\sisgabweb\DownFile.exe SistemaEnvases.exe`.

## 📐 Reglas de Negocio

### 🔒 Restricciones
- Solo los usuarios incluidos en el arreglo `usuariostotales` (`"JAVIER"`, `"N"`, `"ADMINISTRA"`, `"RODOLFO"`) pueden ver y operar las pestañas INVENTARIO INICIAL, INVENTARIO CAJONES e INVENTARIO ESPARRAGO, así como modificar las fechas de salida/entrada.
- Los proveedores con clave `"01"`, `"03"`, `"RO"` y `"212"` **no** se registran automáticamente en el control de espárrago (`Tb_ENV_PROV_CAJ_ESPARRAGO`), aunque muevan el envase clave "81".
- El corte diario para envases generales, cajones y espárrago **no se reprocesa** si ya existe un registro correspondiente en `tb_registro_movimientos` para la fecha actual (detalle: "CORTE INVENTARIO ENVASES", "CORTE INVENTARIO ENVASES CAJONES PROVEEDOR", "CORTE INVENTARIO ENVASES CAJONES ESPARRAGO PROVEEDOR").
- La pantalla de cambios de inventario inicial de cajones y espárrago (`dateTimecajones_ValueChanged`, `dateTimePickerEsparrago_ValueChanged`) solo permite operar sobre el **mes actual de la máquina** (`mes_actual_Maquina`); fuera de ese mes, se muestra advertencia y se bloquea la operación.
- El guardado del inventario físico de canastillas (`BtnSavFis_Click`) **se deshabilita permanentemente** una vez guardado correctamente para esa fecha (controla mediante `InvF > 0` al cargar y `BtnSavFis.Enabled = false` tras guardar).
- Si se detecta una versión más nueva del ejecutable en el servidor (comparando fechas de archivo), la aplicación se cierra forzosamente antes de permitir operación.

### ✅ Validaciones
- Para registrar una salida o entrada: proveedor, chofer y operador no pueden estar vacíos, y debe existir al menos un envase agregado en la grilla de productos.
- La cantidad de un envase debe ser numérica y mayor a cero (`cant.Text`, `cantidad_entr.Text`).
- No se permite agregar dos veces la misma combinación de envase + producto en la grilla de salida/entrada (`ProductosGuardar.Select(...)`).
- En el cierre de inventario físico (`Acumula`), la suma de las categorías (con producto + vacías + con basura + por reparar + envase roto) debe **coincidir exactamente** con el valor de "INVENTARIO FISICO" capturado para cada envase; si no coincide, se bloquea el guardado y se muestra el detalle de la discrepancia.
- Si la suma total del inventario físico (`SumaInvFis`) es igual a 0, se considera inválido y se bloquea el guardado (`Acumula` retorna `false`).
- Si no se encuentra información para un proveedor/fecha en consultas de inventario, se muestra un mensaje "NO ENCONTRE INFORMACION..." y se aborta la operación.

### 🔁 Agrupaciones
- Los reportes de inventario por proveedor (CAJONES, ESPARRAGO, button10/11/17_Click) agrupan los registros insertando una **fila de encabezado por proveedor** (con `ID = ""`, `NOMBRE = nombre del proveedor`) seguida de las filas de detalle por envase.
- El reporte general muestra una **fila final "TOTALES"** que suma inventario inicial, entradas, salidas e inventario final de todos los registros.
- El Kardex (`button14_Click`, `button15_Click`) agrupa movimientos por proveedor/rancho/tabla, insertando subtotales ("TOTAL PROV: ...") cuando cambia el proveedor.
- En reportes Excel, las filas con `ESTATUS = 'C'` (cancelado) se resaltan en amarillo, las filas "TOTAL"/"TOTALES" en verde, y en el Kardex las entradas se resaltan en verde claro y las salidas en rojo.

### ⚙️ Reglas Operativas
- Cada salida/entrada genera un **folio secuencial** obtenido mediante `SELECT TOP(1) FOLIO ... ORDER BY FOLIO DESC` (o `SCOPE_IDENTITY()` en el flujo transaccional nuevo de `BtnSplitPed_Click`).
- El cálculo estándar de inventario final es: `Inventario Final = Inventario Inicial + Entradas − Salidas`.
- Para **canastillas de espárrago (clave "81")**, la fórmula de inventario inicial del día siguiente se **invierte**: `Inventario Inicial (día+1) = Inventario Inicial + Salidas − Entradas` (comentario explícito: "se invierte posición debido a que lo solicitó Javier Castrejón, las Salidas Ahora serán positivas").
- El corte diario para cajones por proveedor solo procesa las claves de envase **"02" y "80"**; el corte de espárrago solo procesa la clave **"81"**.
- Si una salida o entrada se registra con fecha anterior a la fecha actual del sistema (`fecha_hoy_hoy`), se marca como movimiento **"ATRASADO"** en `tb_registro_movimientos` y se ejecuta `realizar_corte` con motivo "DESTIEMPO" para recalcular inventarios día por día hasta la fecha actual.
- Las cancelaciones de folios (`button5_Click` para salidas, `cancelar_ent_Click` para entradas) marcan el registro con `SAL_STATUS`/`ENT_STATUS = 'C'`, revierten cantidades en `TB_MSTR_ENVASES` y `TB_MSTR_INV_ENVASES`, y ejecutan `realizar_corte` con motivo "CANCELACION".
- El envío de correo de inventario físico usa SMTP en `mail1.mrlucky.com.mx`, puerto 587, con SSL habilitado y remitente `sistemas@mrlucky.com.mx` / `ricardo.cortes@mrlucky.com.mx`.

## 🔗 Dependencias

- **`Utilerias.dll`** (externa, `C:\SisGabWeb\Utilerias.dll`): provee `ConnectionString`, `Usu_login`, `Inicio_sesion`, `Nombre_equipo`, `Idioma`, `Login`, `validar_ip()`, `SendMail()`, `registrar_movimiento3()`.
- **`Microsoft.Office.Interop.Excel`** (`C:\SisGabWeb\Microsoft.Office.Interop.Excel.dll`): generación de reportes en Excel (CAJONES, ESPARRAGO, Form1, Form2).
- **`Microsoft.Office.Interop.Outlook`** y **`Interop.Microsoft.Office.Core`**: referenciados en el proyecto, sin uso evidente identificado en el código analizado.
- **`System.Data.SqlClient`**: acceso a SQL Server mediante `SqlConnection`, `SqlCommand`, `SqlDataAdapter`, `SqlDataReader`, `SqlTransaction`.
- **`System.Net.Mail`**: envío de correos (`SmtpClient`, `MailMessage`) para reporte de inventario físico.
- **Tablas SQL Server principales**: `tb_cat_proveedor`, `tb_cat_ranchos`, `tb_cat_tablas`, `tb_cat_producto`, `tb_cat_envases`, `TB_SALIDAS_ENVASES`, `TB_DETSALIDAS_ENVASES`, `TB_ENTRADAS_ENVASES`, `TB_DETENTRADAS_ENVASES`, `TB_MSTR_ENVASES`, `TB_MSTR_INV_ENVASES`, `TB_MSTR_INV_ENVASES_dos`, `TB_MSTR_INV_ENVASES_sin_corte`, `Tb_ENV_PROV_CAJ`, `TB_MSTR_INV_CAJAS_PROV`, `Tb_ENV_PROV_CAJ_ESPARRAGO`, `TB_MSTR_INV_CAJAS_PROV_ESPARRAGO`, `tb_registro_movimientos`, `tb_cat_historial_dia`, `tb_mstr_email`.
- **Formularios internos relacionados**: `CAJONES`, `ESPARRAGO`, `DetalleEnvase` (Form2), `Form3`, `Form4`, `detalle_envase`, `GENERAR`, `Mensaje`, `Reporte_Movimientos`.
- **Actualizador externo**: `c:\sisgabweb\DownFile.exe`, archivo de control `\\gabira1\sisgabweb\Valida.txt`, ejecutable de referencia `\\gabira1\sisgabweb\SistemaEnvases.exe` y `c:\sisgabweb\SistemaEnvases.exe`.
- **`DataHelperProveedor`**: clase vacía sin implementación.

## ⚠️ Riesgos Técnicos

- **Inyección SQL generalizada**: la gran mayoría de las consultas (entradas, salidas, reportes, cortes, validaciones) construyen cadenas SQL por concatenación directa de valores de controles de UI (`clbprov.Text`, `Datetimepickerrepor.Text`, etc.) sin parametrización, exponiendo el sistema a inyección SQL. Solo `BtnSplitPed_Click` y `validarproveedoresparragoT` usan parámetros.
- **Manejo inconsistente de transacciones**: el flujo de salidas (`BtnSplitPed_Click`) usa `SqlTransaction`, pero el flujo de entradas (`guardar_ent_Click`) y `realizar_corte` ejecutan múltiples `ExecuteNonQuery` sin transacción, lo que puede dejar datos inconsistentes ante fallos parciales.
- **Apertura/cierre repetitivo de conexiones SQL** dentro de bucles (`thisConnecion.Open()`/`Close()` en cada iteración de `foreach`), lo que genera sobrecarga de rendimiento y riesgo de fugas de conexión si ocurre una excepción entre `Open` y `Close`.
- **Dependencia de interoperabilidad COM con Excel** (`Microsoft.Office.Interop.Excel`): requiere Excel instalado en cada máquina cliente; los objetos COM no siempre se liberan correctamente (`Marshal.ReleaseComObject` se usa de forma inconsistente, en algunos casos antes de terminar de usar el objeto, lo cual es un error funcional).
- **Lógica de actualización basada en rutas UNC fijas** (`\\gabira1\sisgabweb\...`, `c:\sisgabweb\...`): frágil ante cambios de infraestructura de red o nombres de servidor.
- **Hardcoding de credenciales y direcciones de correo** (`UsuMail`, `PwdMail`, `UsuMail2`, `PwdMail2`, credenciales SMTP) directamente en el código fuente.
- **Cadena de conexión gestionada externamente** (`Utilerias.Class1.ConnectionString`) pero replicada en comentarios con credenciales en texto plano (`"user id=sa; password=Gabira1;..."`), indicando exposición histórica de credenciales.
- **Uso de `dynamic` y `Type.GetTypeFromProgID("Excel.Application")`** mezclado con instancias tipadas de Excel Interop en el mismo método (`button14_Click`, `button15_Click`), generando código duplicado y rutas de ejecución confusas; en `button15_Click` se libera el objeto COM (`Marshal.ReleaseComObject`) **antes** de usar `aplicacion.Columns.AutoFit()`, lo que puede causar excepciones.
- **Bloques de código comentado extensos** (especialmente en `Form1_Load` y `INVENTARIO_CAJONES_PROVEEDOR`) indican lógica legacy parcialmente migrada o desactivada sin limpieza, dificultando el mantenimiento.
- **Recalculo iterativo día por día** (`bg_DoWork`, `bgcajones_DoWork`, `bgESPARRAGO_DoWork`, `realizar_corte`) con `Thread.Sleep(50)` por iteración: el tiempo de procesamiento crece linealmente con la cantidad de días/envases/proveedores, lo que puede volverse muy lento si el usuario selecciona una fecha muy antigua.
- **Múltiples controles llamados `button3`, `Form2`, etc., con responsabilidades distintas** entre formularios (por ejemplo, `button3_Click` en Form1 cierra sesión, mientras que en CAJONES/ESPARRAGO exporta a Excel), lo que incrementa el riesgo de confusión durante mantenimiento.
- **Acceso directo a `Cells[index]` por posición numérica** en grillas y filas de `DataTable` (en lugar de por nombre de columna) en múltiples reportes Excel, frágil ante cambios de esquema de columnas.

## 🧪 Casos Edge

- Si `inv_x_proveedorini.Rows.Count == 0` en los reportes de corte (button7_Click), se informa "CORTE DIARIO NO REALIZADO, NO SE PUEDE MOSTRAR INFORMACION" y se aborta, pero no se ofrece acción correctiva automática.
- Si el rancho o tabla seleccionados son `"Sin Ranchos Disponibles"` / `"Sin Tablas Disponibles"`, se almacenan como cadena vacía en la base de datos.
- En `dateTimeinventarioinicial_ValueChanged` (Form2/DetalleEnvase), si no se encuentra información para la fecha/proveedor, se muestra "NO ENCONTRE INFORMACION PARA ESTE PROVEEDOR" — mensaje genérico que no distingue entre "fecha sin corte" y "proveedor inexistente".
- En `BtnActxDia_Click`, si `cmd.ExecuteNonQuery()` (UPDATE) afecta **más de 0 filas** (`nRegs > 0`), el código inserta un **nuevo** registro en `TB_MSTR_INV_CAJAS_PROV_ESPARRAGO`, lo cual es contraintuitivo: normalmente un UPDATE exitoso no debería requerir un INSERT adicional; esto podría generar registros duplicados.
- El cálculo de `Total_Inventario_Final` en varios reportes usa `Inventario Inicial + Salidas − Entradas` (orden invertido respecto a la fórmula de fila individual `Inventario Inicial + Entradas − Salidas`), generando una posible discrepancia entre el total y la suma de las filas individuales.
- Si `DGInvFis.Rows[i].Cells["INVENTARIO FISICO"].Value` no es convertible a decimal (vacío o no numérico), `Acumula()` y `Suma()` lanzarían una excepción no controlada (`Convert.ToDecimal`).
- Si el archivo `\\gabira1\sisgabweb\Valida.txt`, `SistemaEnvases.exe` (servidor) o `SistemaEnvases.exe` (local) no existen, `ValidaActualizacion()` simplemente no detecta actualización (retorna `false`), sin notificar al usuario que la validación no pudo completarse.
- En `button14_Click`/`button15_Click`/`button13_Click`, si `radioButton1`, `radioButton2` ni la condición "else" (Kardex) tienen un valor `Checked` consistente entre ellos (mutuamente excluyentes garantizados por `GroupBox`), el flujo "else" siempre captura tanto entradas como salidas como Kardex.

## 🧱 Suposiciones Detectadas

- Se asume que `Utilerias.Class1.ConnectionString`, `Usu_login`, `Inicio_sesion` y `Nombre_equipo` están correctamente inicializados antes de usarse en `Form1` y `Program`.
- Se asume que el servidor SQL siempre responde a `SELECT SYSDATETIME()` y que esta fecha es la fuente de verdad para "hoy" (`fecha_hoy_hoy`).
- Se asume que las fechas almacenadas en `TB_MSTR_INV_ENVASES_dos` y tablas relacionadas siguen el formato `dd/MM/yyyy` de forma consistente para comparaciones de cadenas y `Convert.ToDateTime`.
- Se asume que Excel está instalado y registrado (ProgID `"Excel.Application"`) en cada equipo cliente donde se ejecuten los reportes.
- Se asume que solo una instancia de la aplicación corre por máquina/usuario a la vez, dado el control de sesión basado en `nombre_maquina` y `fin_sesion IS NULL`.
- Se asume que los proveedores, ranchos, tablas, productos y envases no cambian sus claves una vez creados, ya que se usan como claves de unión (`JOIN`) y de actualización en múltiples tablas.
- Se asume que el campo `env_inventario = '1'` en `tb_cat_envases` identifica de forma estable los envases sujetos a control de inventario diario.

## 📈 Recomendaciones Técnicas

- Migrar todas las consultas SQL construidas por concatenación a **consultas parametrizadas** (`SqlParameter`), priorizando los flujos de entradas, cortes y reportes que actualmente no usan parámetros, para eliminar el riesgo de inyección SQL.
- Envolver en **transacciones SQL** (`SqlTransaction`) el flujo completo de `guardar_ent_Click` y `realizar_corte`, replicando el patrón ya aplicado en `BtnSplitPed_Click`, para garantizar atomicidad ante fallos.
- Extraer la lógica de cálculo de inventario (fórmulas estándar e invertida para espárrago) a un **método/servicio compartido** y documentar explícitamente por qué la fórmula de espárrago está invertida, para evitar que futuros mantenedores la "corrijan" por error.
- Reemplazar credenciales y direcciones de correo hardcodeadas por **configuración externa** (archivo de configuración o servicio seguro de secretos).
- Revisar y eliminar el código duplicado/comentado de generación de Excel (`dynamic` vs. tipado, múltiples instancias `aplicacion`), unificando en un solo enfoque y corrigiendo el orden de `Marshal.ReleaseComObject` para que se invoque después de finalizar todas las operaciones sobre el objeto COM.
- Sustituir el acceso a conexiones SQL dentro de bucles por **una sola conexión abierta durante todo el proceso** (o uso de `using` para garantizar liberación), reduciendo overhead y riesgo de fugas.
- Agregar manejo de excepciones (`try/catch`) alrededor de conversiones (`Convert.ToDecimal`, `Convert.ToInt32`) en `Acumula()`, `Suma()` y cálculos de inventario, para evitar caídas de la aplicación ante datos nulos o no numéricos.
- Revisar la lógica de `BtnActxDia_Click` para confirmar si el INSERT adicional tras un UPDATE exitoso es intencional o un error que provoca duplicados en `TB_MSTR_INV_CAJAS_PROV_ESPARRAGO`.
- Documentar y, de ser posible, eliminar el mecanismo de actualización basado en comparación de fechas de archivo en rutas UNC fijas, sustituyéndolo por un mecanismo de versionado más robusto (por ejemplo, número de versión en base de datos o servicio de actualización dedicado).
- Limpiar bloques de código comentado de gran tamaño (lógica legacy de corte diario duplicada en `Form1_Load`, `INVENTARIO_CAJONES_PROVEEDOR`) para reducir ruido y confusión en el mantenimiento.

## 🧾 Resumen Ejecutivo

SistemaEnvases es la herramienta que la empresa utiliza día a día para llevar el control de los envases (cajones y canastillas de espárrago) que entran y salen de la planta, organizados por proveedor, rancho y producto. Cada movimiento (entrada o salida) genera un folio, se imprime un comprobante y se actualiza automáticamente el inventario disponible. Cada noche o al inicio del día, el sistema realiza un "corte" que calcula cuánto inventario debería quedar para el día siguiente, tanto a nivel general como por proveedor. Además, permite capturar el conteo físico real de canastillas, compararlo contra lo que el sistema calculó, y enviar un correo con el resumen y las diferencias encontradas. También genera reportes en Excel para análisis administrativo y, al cerrarse, verifica si hay una versión más nueva del programa disponible en el servidor para actualizarse automáticamente.

Desde el punto de vista de riesgo de negocio, el sistema depende fuertemente de que los datos de fecha y las cantidades capturadas sean correctos, ya que los cálculos de inventario de los días siguientes se construyen en cadena a partir de los anteriores; un error o dato faltante en un día puede propagarse a los días posteriores hasta que se ejecute un nuevo corte o recálculo manual. Asimismo, el sistema depende de tener Microsoft Excel instalado para generar reportes, y de la disponibilidad de ciertas rutas de red compartidas para el proceso de actualización automática.
