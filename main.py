from os import mkdir, path
import sqlite3
import xlsxwriter
import PySimpleGUI as sg

camino = 'C:/NOMEDU'

if not path.exists(camino):
    mkdir(camino)
    conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
    cursor_obj = conexion_obj.cursor()
    cursor_obj.execute("""DROP TABLE IF EXISTS PERSONAL""")
    tabla1 = """ CREATE TABLE PERSONAL(
    ID INTEGER PRIMARY KEY,
    Puesto CHAR(120) NOT NULL,
    Nombre CHAR(40) NOT NULL,
    Ap_Paterno CHAR(40) NOT NULL,
    Ap_Materno CHAR(40) NOT NULL,
    RFC VARCHAR(13) NOT NULL UNIQUE,
    Tel_Contacto INT(10) NOT NULL,
    Tel_Emergencia INT(10),
    Sueldo_Hora INT NOT NULL,
    Hora_Entrada TEXT NOT NULL,
    Hora_Salida TEXT NOT NULL,
    Horas_Obligatorias INT(3) DEFAULT 0,
    Horas_Trabajadas INT(3) DEFAULT 0,
    RetardosQuincena INT NOT NULL DEFAULT 0
    ); """
    cursor_obj.execute(tabla1)
    tabla2 = """CREATE TABLE ENTSAL(
    ID INTEGER NOT NULL,
    Llegada TEXT NOT NULL,
    Salida TEXT NOT NULL,
    Retardo INTEGER NOT NULL DEFAULT 0,
    Horas_Trabajadas INTEGER NOT NULL DEFAULT 0,
    Notas TEXT
    );"""
    cursor_obj.execute(tabla2)
    tabla3 = """CREATE TABLE HORARIOS(
    Grupo TEXT NOT NULL,
    Dia TEXT NOT NULL,
    Materia TEXT NOT NULL,
    Profesor TEXT NOT NULL,
    Salon TEXT NOT NULL,
    Horario_Entrada TEXT NOT NULL,
    Horario_Salida TEXT NOT NULL
    );"""
    cursor_obj.execute(tabla3)
    tabla4 = """CREATE TABLE INGEG(
    Forma TEXT NOT NULL,
    Concepto TEXT NOT NULL,
    Monto INTEGER NOT NULL,
    Fecha TEXT NOT NULL
    );"""
    cursor_obj.execute(tabla4)
    tabla5 = """CREATE TABLE RECFECHA(
    Fecha TEXT NOT NULL
    );"""
    cursor_obj.execute(tabla5)

    conexion_obj.commit()
    conexion_obj.close()

layout1 = [[sg.Button('Registrar o Revisar Personal')],
           [sg.Button('Registrar Entrada/Salida')],
           [sg.Button('Registrar o Consultar Horarios')],
           [sg.Button('Solicitar Tutoría')],
           [sg.Button('Registro de Ingresos/Egresos')],
           [sg.Button('Salir')]
           ]

layout2 = [[sg.Text('Registro o Revisión de Personal ')],
           [sg.Button('Crear un nuevo Registro')],
           [sg.Button('Revisar un Registro')],
           [sg.Button('Regresar', key='reg2')],
           [sg.Button('Salir', key='salir2')]
           ]

layout3 = [[sg.Text('Creación de Registro')],

           # Cambiar por lista de puestos como prof maestra limpieza director otros
           [sg.Text('Puesto')],
           [sg.Input(key='puesto')],
           [sg.Text('Nombre')],
           [sg.Input(key='nombre')],
           [sg.Text('Apellido Paterno')],
           [sg.Input(key='appaterno')],
           [sg.Text('Apellido Materno')],
           [sg.Input(key='apmaterno')],
           [sg.Text('RFC')],
           [sg.Input(key='rfc')],
           [sg.Text('Teléfono de Contacto')],
           [sg.Input(key='tel_contacto')],
           [sg.Text('Teléfono de Emergencia')],
           [sg.Input(key='tel_emergencia')],
           [sg.Text('Sueldo por Hora')],
           [sg.Input(key='sueldo_hora')],
           [sg.Text('Hora de Entrada (Formato HH:MM:SS)')],
           [sg.Input(key='hora_entrada')],
           [sg.Text('Hora de Salida (Formato HH:MM:SS)')],
           [sg.Input(key='hora_salida')],
           [sg.Text('Horas contratadas por Quincena')],
           [sg.Input(key='hora_contratada')],
           [sg.Button('Registrar')],
           [sg.Text('Número de Registro')],
           [sg.Input(key='numRegistro', disabled=True, disabled_readonly_background_color='dark gray')],
           [sg.Button('Regresar', key='reg3')],
           [sg.Button('Salir', key='salir3')]
           ]

layout4 = [[sg.Text('Revisión de Registro')],
           [sg.Text('Número de Registro a Buscar')],
           [sg.Input(key='numregbus')],
           [sg.Button('Buscar'), sg.Text('', key='errorBusqueda')],
           [sg.Button('Regresar', key='reg4')],
           [sg.Button('Salir', key='salir4')]
           ]

layout5 = [[sg.Text('Registro de Entrada/Salida')],
           [sg.Button('Registrar Entrada/Salida', key='botentsal')],
           [sg.Button('Registrar Incapacidad/Vacaciones')],
           [sg.Button('Regresar', key='reg5')],
           [sg.Button('Salir', key='salir5')]
           ]

layout6 = [[sg.Text('Registro de Entrada/Salida')],
           [sg.Text('Número de Registro')],
           [sg.Input(key='numregentsal')],
           [sg.Button('Registrar Entrada'), sg.Text('', key='errorRegEnt')],
           [sg.Button('Registrar Salida'), sg.Text('', key='errorRegSal')],
           [sg.Button('Regresar', key='reg6')],
           [sg.Button('Salir', key='salir6')]
           ]

layout7 = [[sg.Text('Registro de Incapacidad/Vacaciones')],
           [sg.Text('Número de Registro')],
           [sg.Input(key='numregincvac')],
           [sg.Radio('Incapacidad', 'incvac', key='radioinc'), sg.Radio('Vacaciones', 'incvac', key='radiovac')],
           [sg.Text('Razón (Opcional)')],
           [sg.Input(key='razon')],
           [sg.Text('Fecha de Inicio (Formato YYYY-MM-DD)')],
           [sg.Input(key='fechaInicio')],
           [sg.Text('Fecha de Termino (Formato YYYY-MM-DD)')],
           [sg.Input(key='fechaTermino')],
           [sg.Button('Registrar', key='registrarincvac'), sg.Text('', key='errorRegistrarincvac')],
           [sg.Button('Regresar', key='reg7')],
           [sg.Button('Salir', key='salir7')]]

layout8 = [[sg.Text('Revisión de Registro')],
           [sg.Text('Puesto')],
           [sg.Input(key='puestoB')],
           [sg.Text('Nombre')],
           [sg.Input(key='nombreB')],
           [sg.Text('Apellido Paterno')],
           [sg.Input(key='appaternoB')],
           [sg.Text('Apellido Materno')],
           [sg.Input(key='apmaternoB')],
           [sg.Text('RFC')],
           [sg.Input(key='rfcB')],
           [sg.Text('Teléfono de Contacto')],
           [sg.Input(key='tel_contactoB')],
           [sg.Text('Teléfono de Emergencia')],
           [sg.Input(key='tel_emergenciaB')],
           [sg.Text('Sueldo por Hora')],
           [sg.Input(key='sueldo_horaB')],
           [sg.Text('Hora de Entrada (Formato HH:MM:SS)')],
           [sg.Input(key='hora_entradaB')],
           [sg.Text('Hora de Salida (Formato HH:MM:SS)')],
           [sg.Input(key='hora_salidaB')],
           [sg.Text('Horas contratadas por mes')],
           [sg.Input(key='hora_contratadaB')],
           [sg.Button('Borrar')],
           [sg.Button('Modificar'), sg.Text('', key='menMod')],
           [sg.Button('Regresar', key='reg8')],
           [sg.Button('Salir', key='salir8')]
           ]

layout9 = [[sg.Text('Registro o Consulta de Horarios')],
           [sg.Button('Registrar Horario', key='regHor')],
           [sg.Button('Consultar Horario', key='conHor')],
           [sg.Button('Regresar', key='reg9')],
           [sg.Button('Salir', key='salir9')]
           ]

layout12 = [[sg.Text('Registro de Horarios')],
            [sg.Text('Ingresar Grupo')],
            [sg.Input(key='grupoHor')],
            [sg.Text('Seleccione los días de la materia')],
            [sg.Checkbox('Lunes', key='lunes'), sg.Checkbox('Martes', key='martes'),
             sg.Checkbox('Miércoles', key='miercoles'),
             sg.Checkbox('Jueves', key='jueves'), sg.Checkbox('Viernes', key='viernes')],
            [sg.Text('Ingresar Materia')],
            [sg.Input(key='materia')],
            [sg.Text('Ingresar Profesor')],
            [sg.Input(key='profe')],
            [sg.Text('Ingrese el Salón')],
            [sg.Input(key='salon')],
            [sg.Text('Ingrese Horario de Entrada (Fomato HH:MM:SS)')],
            [sg.Input(key='horEntrada')],
            [sg.Text('Ingrese Horario de Salida (Fomato HH:MM:SS)')],
            [sg.Input(key='horSalida')],
            [sg.Button('Registrar', key='regHorComp'), sg.Text('', key='menRegHor')],
            [sg.Button('Regresar', key='reg12')],
            [sg.Button('Salir', key='salir12')]]

headings = [' Grupo ', '  Día  ', ' Materia ', ' Profesor ', ' Salón ', ' Entrada ', ' Salida ']
datos = []

col1 = [[sg.Text('Consulta de Horarios')],
        [sg.Text('Grupo')],
        [sg.Input(key='grupoBus')],
        [sg.Button('Buscar', key='buscargrupo')],
        [sg.Table(values=datos, headings=headings, key='tablaHor', enable_events=True)],
        [sg.Button('Modificar', key='modClaseGrupo', disabled=True)],
        [sg.Button('Borrar Grupo', key='borGrupo')],
        [sg.Button('Regresar', key='reg13')],
        [sg.Button('Salir', key='salir13')]
        ]

col2 = [[sg.Text('Grupo')],
        [sg.Input(key='grupoHorB')],
        [sg.Text('Días de la materia')],
        [sg.Checkbox('Lunes', key='lunesB'), sg.Checkbox('Martes', key='martesB'),
         sg.Checkbox('Miércoles', key='miercolesB'),
         sg.Checkbox('Jueves', key='juevesB'), sg.Checkbox('Viernes', key='viernesB')],
        [sg.Text('Materia')],
        [sg.Input(key='materiaB')],
        [sg.Text('Profesor')],
        [sg.Input(key='profeB')],
        [sg.Text('Salón')],
        [sg.Input(key='salonB')],
        [sg.Text('Horario de Entrada (Fomato HH:MM:SS)')],
        [sg.Input(key='horEntradaB')],
        [sg.Text('Horario de Salida (Fomato HH:MM:SS)')],
        [sg.Input(key='horSalidaB')],
        [sg.Button('Modificar Clase', key='modHor'), sg.Text('', key='menModHor')],
        [sg.Button('Borrar Clase', key='borClase')]
        ]

layout13 = [[sg.Column(col1), sg.Column(col2, visible=False, key='colbus')]]

col3 = [[sg.Text('Grupo del alumno')],
        [sg.Input(key='grupoHorT'), sg.Button('Buscar', key='busTutGrupo')],
        [sg.Text('Nombre del alumno')],
        [sg.Input(key='alumno', disabled=True)],
        [sg.Text('Días de la Tutoria')],
        [sg.Checkbox('Lunes', key='lunesT', disabled=True), sg.Checkbox('Martes', key='martesT', disabled=True),
         sg.Checkbox('Miércoles', key='miercolesT', disabled=True),
         sg.Checkbox('Jueves', key='juevesT', disabled=True), sg.Checkbox('Viernes', key='viernesT', disabled=True)],
        [sg.Text('Materia de la Tutoría')],
        [sg.Input(key='materiaT', disabled=True)],
        [sg.Text('Profesor de la Tutoría')],
        [sg.Input(key='profeT', disabled=True)],
        [sg.Text('Salón')],
        [sg.Input(key='salonT', disabled=True)],
        [sg.Text('Horario de Entrada (Fomato HH:MM:SS)')],
        [sg.Input(key='horEntradaT', disabled=True)],
        [sg.Text('Horario de Salida (Fomato HH:MM:SS)')],
        [sg.Input(key='horSalidaT', disabled=True)],
        [sg.Button('Registrar Tutoría', key='regTut'), sg.Text('', key='menTut')], ]

datTut = []

col4 = [[sg.Table(values=datTut, headings=headings, key='tablaTut')], ]

layout14 = [[sg.Column(col3), sg.Column(col4, key='colTabTut', visible=False)],
            [sg.Button('Regresar', key='reg14')],
            [sg.Button('Salir', key='salir14')]
            ]

layout10 = [[sg.Text('Registro de Ingresos/Egresos')],
            [sg.Button('Registrar Ingreso/Egreso', key='IngEg')],
            [sg.Button('Calcular Nómina de Personal', key='calcNom'), sg.Text('', key='menCalcNom')],
            [sg.Text('Ingresar día de inicio del Reporte (Formato YYYY-MM-DD)')],
            [sg.Input(key='diaInicio')],
            [sg.Text('Ingresar día de fin del Reporte (Formato YYYY-MM-DD)')],
            [sg.Input(key='diaFin')],
            [sg.Button('Generar Reporte', key='genRep'), sg.Text('', key='menGenRep')],
            [sg.Button('Regresar', key='reg10')],
            [sg.Button('Salir', key='salir10')]
            ]

layout11 = [[sg.Text('Ingresos/Egresos')],
            [sg.Radio('Ingreso', 'ingeng', key='ing'), sg.Radio('Egreso', 'ingeng', key='eg')],
            [sg.Text('Ingresar Concepto')],
            [sg.Input(key='concepto')],
            [sg.Text('Ingrese Monto')],
            [sg.Input(key='monto')],
            [sg.Button('Registrar', key='regIngEg'), sg.Text('', key='menRegIngEg')],
            [sg.Button('Regresar', key='reg11')],
            [sg.Button('Salir', key='salir11')]
            ]

# ----------- Create actual layout using Columns and a row of Buttons
layout = [
    [sg.Text('Sistema de Nomina Escolar "Nomedu"')],
    [sg.Column(layout1, key='-COL1-'),
     sg.Column(layout2, visible=False, key='-COL2-'),
     sg.Column(layout3, visible=False, key='-COL3-'),
     sg.Column(layout4, visible=False, key='-COL4-'),
     sg.Column(layout5, visible=False, key='-COL5-'),
     sg.Column(layout6, visible=False, key='-COL6-'),
     sg.Column(layout7, visible=False, key='-COL7-'),
     sg.Column(layout8, visible=False, key='-COL8-'),
     sg.Column(layout9, visible=False, key='-COL9-'),
     sg.Column(layout10, visible=False, key='-COL10-'),
     sg.Column(layout11, visible=False, key='-COL11-'),
     sg.Column(layout12, visible=False, key='-COL12-'),
     sg.Column(layout13, visible=False, key='-COL13-'),
     sg.Column(layout14, visible=False, key='-COL14-')

     ],
]

window = sg.Window('NOMEDU', layout, location=(300, 0))

layout = 1
incvac = ''
forma = ''
listVal = []
while True:
    event, values = window.read()
    if event in (None, 'Salir'):
        break

    if event == 'Registrar o Revisar Personal':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '2'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'Crear un nuevo Registro':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '3'
        window[f'-COL{layout}-'].update(visible=True)
        window['numRegistro'].update('')

    if event == 'Registrar':
        puesto = values["puesto"]
        window["puesto"].update("")
        nombre = values["nombre"]
        window["nombre"].update("")
        appaterno = values["appaterno"]
        window["appaterno"].update("")
        apmaterno = values["apmaterno"]
        window["apmaterno"].update("")
        rfc = values["rfc"]
        window["rfc"].update("")
        tel_contacto = int(values["tel_contacto"])
        window["tel_contacto"].update("")
        tel_emergencia = int(values["tel_emergencia"])
        window["tel_emergencia"].update("")
        sueldo_hora = float(values["sueldo_hora"])
        window["sueldo_hora"].update("")
        hora_entrada = values["hora_entrada"]
        window["hora_entrada"].update("")
        hora_salida = values["hora_salida"]
        window["hora_salida"].update("")
        horas_contrato = values['hora_contratada']
        window['hora_contratada'].update('')
        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()
        cursor_obj.execute("""SELECT MAX(ID) FROM PERSONAL""")
        rid = cursor_obj.fetchone()
        if rid[0] is None:
            rid = 1
        else:
            rid = rid[0]
            rid += 1
        comando = """INSERT INTO PERSONAL VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)"""
        valores = (rid, puesto, nombre, appaterno, apmaterno, rfc, tel_contacto, tel_emergencia, sueldo_hora,
                   hora_entrada, hora_salida, horas_contrato, 0, 0)
        cursor_obj.execute(comando, valores)
        window["numRegistro"].update(rid)
        conexion_obj.commit()
        conexion_obj.close()

    if event == 'Revisar un Registro':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '4'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'Buscar':
        if not values['numregbus'] == '':
            conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
            cursor_obj = conexion_obj.cursor()
            cursor_obj.execute("""SELECT * FROM PERSONAL WHERE ID = ?""", (int(values['numregbus']),))
            fila = cursor_obj.fetchone()
            if fila is not None:
                window['errorBusqueda'].update('')
                window["puestoB"].update(fila[1])
                window["nombreB"].update(fila[2])
                window["appaternoB"].update(fila[3])
                window["apmaternoB"].update(fila[4])
                window["rfcB"].update(fila[5])
                window["tel_contactoB"].update(fila[6])
                window["tel_emergenciaB"].update(fila[7])
                window["sueldo_horaB"].update(fila[8])
                window["hora_entradaB"].update(fila[9])
                window["hora_salidaB"].update(fila[10])
                window['hora_contratadaB'].update(fila[11])

                window["puestoB"].update(disabled=False)
                window["nombreB"].update(disabled=False)
                window["appaternoB"].update(disabled=False)
                window["apmaternoB"].update(disabled=False)
                window["rfcB"].update(disabled=False)
                window["tel_contactoB"].update(disabled=False)
                window["tel_emergenciaB"].update(disabled=False)
                window["sueldo_horaB"].update(disabled=False)
                window["hora_entradaB"].update(disabled=False)
                window["hora_salidaB"].update(disabled=False)
                window["hora_contratadaB"].update(disabled=False)
                window['Borrar'].update(disabled=False)
                window['Modificar'].update(disabled=False)
                window['menMod'].update('')

                window[f'-COL{layout}-'].update(visible=False)
                layout = '8'
                window[f'-COL{layout}-'].update(visible=True)
            else:
                window['numregbus'].update('')
                window['errorBusqueda'].update('Registro no encontrado')
            conexion_obj.close()

    if event == 'Borrar':
        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()
        cursor_obj.execute("""DELETE FROM PERSONAL WHERE ID = ?""", (int(values['numregbus']),))
        conexion_obj.commit()
        conexion_obj.close()
        window["puestoB"].update('')
        window["nombreB"].update('')
        window["appaternoB"].update('')
        window["apmaternoB"].update('')
        window["rfcB"].update('')
        window["tel_contactoB"].update('')
        window["tel_emergenciaB"].update('')
        window["sueldo_horaB"].update('')
        window["hora_entradaB"].update('')
        window["hora_salidaB"].update('')
        window['hora_contratadaB'].update('')
        window['numregbus'].update('')
        window[f'-COL{layout}-'].update(visible=False)
        layout = '4'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'Modificar':
        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()
        puestoB = values["puestoB"]
        window["puestoB"].update("")
        window["puestoB"].update(disabled=True)
        nombreB = values["nombreB"]
        window["nombreB"].update("")
        window["nombreB"].update(disabled=True)
        appaternoB = values["appaternoB"]
        window["appaternoB"].update("")
        window["appaternoB"].update(disabled=True)
        apmaternoB = values["apmaternoB"]
        window["apmaternoB"].update("")
        window["apmaternoB"].update(disabled=True)
        rfcB = values["rfcB"]
        window["rfcB"].update("")
        window["rfcB"].update(disabled=True)
        tel_contactoB = int(values["tel_contactoB"])
        window["tel_contactoB"].update("")
        window["tel_contactoB"].update(disabled=True)
        tel_emergenciaB = int(values["tel_emergenciaB"])
        window["tel_emergenciaB"].update("")
        window["tel_emergenciaB"].update(disabled=True)
        sueldo_horaB = float(values["sueldo_horaB"])
        window["sueldo_horaB"].update("")
        window["sueldo_horaB"].update(disabled=True)
        hora_entradaB = values["hora_entradaB"]
        window["hora_entradaB"].update("")
        window["hora_entradaB"].update(disabled=True)
        hora_salidaB = values["hora_salidaB"]
        window["hora_salidaB"].update("")
        window["hora_salidaB"].update(disabled=True)
        horas_contratoB = values['hora_contratadaB']
        window["hora_contratadaB"].update("")
        window["hora_contratadaB"].update(disabled=True)
        window['Borrar'].update(disabled=True)
        window['Modificar'].update(disabled=True)
        comando = """UPDATE PERSONAL SET Puesto = ?, Nombre = ?, Ap_Paterno = ?, Ap_Materno = ?, RFC = ?,
         Tel_Contacto = ?, Tel_Emergencia = ?, Sueldo_Hora = ?, Hora_Entrada = ?, Hora_Salida = ?, 
         Horas_Obligatorias = ? WHERE ID = ?"""
        valores = (puestoB, nombreB, appaternoB, apmaternoB, rfcB, tel_contactoB, tel_emergenciaB, sueldo_horaB,
                   hora_entradaB, hora_salidaB, horas_contratoB, int(values['numregbus']))
        cursor_obj.execute(comando, valores)
        conexion_obj.commit()
        conexion_obj.close()
        window['menMod'].update('Registro Actualizado')

    if event == 'reg2':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '1'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'reg3':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '2'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'reg4':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '2'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'Registrar Entrada/Salida':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '5'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'botentsal':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '6'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'Registrar Entrada':
        window['errorRegEnt'].update('')
        window['errorRegSal'].update('')
        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()
        if not values['numregentsal'] == '':
            cursor_obj.execute("""SELECT * FROM PERSONAL WHERE ID = ?""", (int(values['numregentsal']),))
            check1 = cursor_obj.fetchone()
            if check1 is not None:
                cursor_obj.execute("""SELECT * FROM ENTSAL WHERE ID = ? AND date(Llegada) = date('now', 'localtime')""",
                                   (int(values['numregentsal']),))
                check = cursor_obj.fetchone()
                if check is None:
                    comando = """INSERT INTO ENTSAL VALUES(?, datetime('now', 'localtime'), '', 0, 0, '');"""
                    cursor_obj.execute(comando, (int(values['numregentsal']),))
                    cursor_obj.execute("""SELECT sum(strftime('%s', Hora_Entrada) - 
                            strftime('%s', (SELECT time(Llegada) FROM ENTSAL 
                            WHERE ID = ? AND date(Llegada) = date('now', 'localtime')))) 
                            FROM PERSONAL WHERE ID = ?""", (int(values['numregentsal']), int(values['numregentsal'])))
                    hora = cursor_obj.fetchone()
                    if hora[0] < -1800:
                        cursor_obj.execute("""UPDATE ENTSAL SET Retardo = 1 
                        WHERE ID = ? AND date(Llegada) = date('now', 'localtime')""", (int(values['numregentsal']),))
                        cursor_obj.execute("""UPDATE PERSONAL SET 
                        RetardosQuincena = (RetardosQuincena + 1) WHERE ID = ?""", (int(values['numregentsal']),))
                    window['errorRegEnt'].update('Entrada registrada correctamente')
                else:
                    window['errorRegEnt'].update('Ya se registro la entrada')
            else:
                window['errorRegEnt'].update('Registro no encontrado')
        else:
            window['errorRegEnt'].update('Ingrese un Número de Registro')
        window['numregentsal'].update('')
        conexion_obj.commit()
        conexion_obj.close()

    if event == 'Registrar Salida':
        window['errorRegEnt'].update('')
        window['errorRegSal'].update('')
        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()
        if not values['numregentsal'] == '':
            cursor_obj.execute("""SELECT * FROM PERSONAL WHERE ID = ?""", (int(values['numregentsal']),))
            check1 = cursor_obj.fetchone()
            if check1 is not None:
                cursor_obj.execute("""SELECT * FROM ENTSAL 
                WHERE ID = ? AND date(Llegada) = date('now', 'localtime') AND Salida = '';""",
                                   (int(values['numregentsal']),))
                check = cursor_obj.fetchone()

                if check is not None:
                    comando = """UPDATE ENTSAL SET Salida = datetime('now', 'localtime') 
                    WHERE ID = ? AND date(Llegada) = date('now', 'localtime');"""
                    cursor_obj.execute(comando, (int(values['numregentsal']),))
                    cursor_obj.execute("""SELECT sum(strftime('%s', Salida) - strftime('%s', Llegada)) 
                                    FROM ENTSAL WHERE ID = ? AND date(Llegada) = date('now', 'localtime')""",
                                       (int(values['numregentsal']),))
                    sec = cursor_obj.fetchone()
                    sec = sec[0] / 60 / 60
                    if (sec % 1) >= 0.83:
                        sec = sec // 1
                        sec += 1
                    else:
                        sec = sec // 1

                    cursor_obj.execute("""UPDATE ENTSAL SET Horas_Trabajadas = ? 
                    WHERE ID = ? AND date(Llegada) = date('now', 'localtime') 
                    AND date(Salida) = date('now', 'localtime')""", (sec, int(values['numregentsal'])))
                    cursor_obj.execute("""UPDATE PERSONAL SET Horas_Trabajadas = (Horas_Trabajadas + ?) WHERE ID = ?""",
                                       (sec, int(values['numregentsal'])))
                    window['errorRegSal'].update('Salida registrada correctamente')
                else:
                    window['errorRegSal'].update('Ya se registro la salida')
            else:
                window['errorRegSal'].update('Registro no encontrado')
        else:
            window['errorRegSal'].update('Ingrese un Número de Registro')
        window['numregentsal'].update('')
        conexion_obj.commit()
        conexion_obj.close()

    if event == 'Registrar Incapacidad/Vacaciones':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '7'
        window[f'-COL{layout}-'].update(visible=True)

    if values['radioinc']:
        incvac = 'Incapacidad '

    if values['radiovac']:
        incvac = 'Vacaciones '

    if event == 'registrarincvac':
        window['errorRegistrarincvac'].update('')
        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()
        if not values['numregincvac'] == '':
            cursor_obj.execute("""SELECT * FROM PERSONAL WHERE ID = ?""", (int(values['numregincvac']),))
            check1 = cursor_obj.fetchone()
            if check1 is not None:
                incvac += values['razon']
                window['razon'].update('')
                fechaInicio = values['fechaInicio']
                window['fechaInicio'].update('')
                fechaTermino = values['fechaTermino']
                window['fechaTermino'].update('')
                stop = True
                dias = 0

                valores = (fechaInicio, fechaTermino)
                valoresCom = (int(values['numregincvac']), fechaInicio, int(values['numregincvac']), fechaInicio,
                              int(values['numregincvac']), incvac)

                while stop:
                    cursor_obj.execute(f"""SELECT IIF((SELECT date(?, '+{dias} days')) = date(?),
                     'Si', 'No')""", valores)
                    check2 = cursor_obj.fetchone()
                    check2 = str(check2[0])
                    if check2 == 'No':
                        cursor_obj.execute(f"""SELECT strftime('%w', date(?, '+{dias} days'))""", (fechaInicio,))
                        diaSemana = cursor_obj.fetchone()
                        if diaSemana[0] == '0' or diaSemana[0] == '6':
                            dias += 1
                            continue
                        else:
                            cursor_obj.execute(f"""INSERT INTO ENTSAL VALUES (?, 
                            datetime(? || ' ' ||(SELECT Hora_Entrada FROM PERSONAL WHERE ID = ?), '+{dias} days'),
                            datetime(? || ' ' ||(SELECT Hora_Salida FROM PERSONAL WHERE ID = ?), '+{dias} days'), 
                            0, 0, ?);""", valoresCom)

                            cursor_obj.execute(f"""SELECT sum(strftime('%s', Salida) - strftime('%s', Llegada)) 
                                                                FROM ENTSAL WHERE ID = ? 
                                                                AND date(Llegada) = date(?, '+{dias} days') 
                                                                AND date(Salida) = date(?, '+{dias} days')""",
                                               (int(values['numregincvac']), fechaInicio, fechaInicio))
                            sec = cursor_obj.fetchone()
                            sec = sec[0] / 60 / 60
                            if (sec % 1) >= 0.83:
                                sec = sec // 1
                                sec += 1
                            else:
                                sec = sec // 1
                            cursor_obj.execute(f"""UPDATE ENTSAL SET Horas_Trabajadas = ? 
                                                                    WHERE ID = ? AND date(Llegada) = date(?, '+{dias} days') 
                                                                    AND date(Salida) = date(?, '+{dias} days')""",
                                               (sec, int(values['numregincvac']), fechaInicio, fechaInicio))
                            cursor_obj.execute(
                                """UPDATE PERSONAL SET Horas_Trabajadas = (Horas_Trabajadas + ?) WHERE ID = ?""",
                                (sec, int(values['numregincvac'])))

                            dias += 1

                    else:
                        cursor_obj.execute(f"""SELECT strftime('%w', date(?, '+{dias} days'))""", (fechaInicio,))
                        diaSemana = cursor_obj.fetchone()
                        if diaSemana[0] == '0' or diaSemana[0] == '6':
                            window['errorRegistrarincvac'].update('Registro exitoso')
                            stop = False
                        else:
                            cursor_obj.execute(f"""INSERT INTO ENTSAL VALUES (?, 
                                                    datetime(? || ' ' ||(SELECT Hora_Entrada FROM PERSONAL WHERE ID = ?),
                                                     '+{dias} days'),
                                                    datetime(? || ' ' ||(SELECT Hora_Salida FROM PERSONAL WHERE ID = ?),
                                                     '+{dias} days'), 
                                                    0, 0, ?);""", valoresCom)

                            cursor_obj.execute(f"""SELECT sum(strftime('%s', Salida) - strftime('%s', Llegada)) 
                                                                FROM ENTSAL WHERE ID = ? 
                                                                AND date(Llegada) = date(?, '+{dias} days') 
                                                                AND date(Salida) = date(?, '+{dias} days')""",
                                               (int(values['numregincvac']), fechaInicio, fechaInicio))
                            sec = cursor_obj.fetchone()
                            sec = sec[0] / 60 / 60
                            if (sec % 1) >= 0.83:
                                sec = sec // 1
                                sec += 1
                            else:
                                sec = sec // 1
                            cursor_obj.execute(f"""UPDATE ENTSAL SET Horas_Trabajadas = ? 
                                                                    WHERE ID = ? AND date(Llegada) = date(?, '+{dias} days') 
                                                                    AND date(Salida) = date(?, '+{dias} days')""",
                                               (sec, int(values['numregincvac']), fechaInicio, fechaInicio))
                            cursor_obj.execute(
                                """UPDATE PERSONAL SET Horas_Trabajadas = (Horas_Trabajadas + ?) WHERE ID = ?""",
                                (sec, int(values['numregincvac'])))
                            window['errorRegistrarincvac'].update('Registro exitoso')
                            stop = False

            else:
                window['errorRegistrarincvac'].update('Registro no encontrado')
        else:
            window['errorRegistrarincvac'].update('Ingrese un Número de Registro')
        window['numregincvac'].update('')
        conexion_obj.commit()
        conexion_obj.close()

    if event == 'Registrar o Consultar Horarios':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '9'
        window[f'-COL{layout}-'].update(visible=True)

    if values['ing']:
        forma = 'Ingreso'

    if values['eg']:
        forma = 'Egreso'

    if event == 'regIngEg':
        concepto = values['concepto']
        window['concepto'].update('')
        monto = float(values['monto'])
        window['monto'].update('')

        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()

        cursor_obj.execute("""INSERT INTO INGEG VALUES (?,?,?, date('now','localtime'))""", (forma, concepto, monto))

        conexion_obj.commit()
        conexion_obj.close()
        window['menRegIngEg'].update('Registrado correctamente')

    if event == 'calcNom':
        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()

        cursor_obj.execute("""SELECT COUNT(*) FROM PERSONAL;""")
        cuenta = cursor_obj.fetchone()
        cursor_obj.execute("""SELECT * FROM PERSONAL;""")
        parar = True
        forma = 'Egreso'
        for i in range(0, cuenta[0]):
            pers = cursor_obj.fetchone()
            nombre = pers[1] + ' ' + pers[3] + ' ' + pers[4] + ' ' + pers[2] + ' Nómina'
            retardos = pers[13]
            horas = pers[12]
            while parar:
                if retardos >= 3:
                    horas -= 4
                    retardos -= 3
                else:
                    parar = False
            resto = horas - pers[11]
            if resto < 0:
                resto = 0
            sueldo = pers[8] * (horas - resto)
            sueldo += (pers[8] * 2) * resto

            if not sueldo == 0:
                conexion_obj.execute("""INSERT INTO INGEG VALUES (?,?,?, date('now', 'localtime'));""",
                                     (forma, nombre, sueldo))

            conexion_obj.execute("""UPDATE PERSONAL SET Horas_Trabajadas = 0, RetardosQuincena = 0 WHERE ID = ?;""",
                                 (pers[0],))
        cursor_obj.execute("""SELECT * FROM RECFECHA""")
        rec = cursor_obj.fetchone()
        if rec == None:
            cursor_obj.execute("""INSERT INTO RECFECHA VALUES (date('now', 'localtime'));""")
        else:
            cursor_obj.execute("""UPDATE RECFECHA SET Fecha = date('now', 'localtime');""")

        window['menCalcNom'].update('Calculo realizado correctamente')
        conexion_obj.commit()
        conexion_obj.close()

    if event == 'genRep':
        row = 0
        col = 0

        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()

        fecha1 = values['diaInicio']
        fecha2 = values['diaFin']

        cursor_obj.execute("""SELECT COUNT(*) FROM INGEG WHERE Fecha BETWEEN date(?) AND date(?)""", (fecha1, fecha2))
        lineaC = cursor_obj.fetchone()
        cursor_obj.execute("""SELECT * FROM INGEG WHERE Fecha BETWEEN date(?) AND date(?)""", (fecha1, fecha2))

        dir = path.expanduser('~') + '\Desktop\Reporte_de_' + fecha1 + '_al_' + fecha2 + '.xlsx'

        workbook = xlsxwriter.Workbook(dir)
        worksheet = workbook.add_worksheet()

        worksheet.write(row, col, 'Forma')
        worksheet.write(row, col + 1, 'Concepto')
        worksheet.write(row, col + 2, 'Monto')
        worksheet.write(row, col + 3, 'Fecha de Registro')
        row += 1

        for i in range(0, lineaC[0]):
            linea = cursor_obj.fetchone()
            worksheet.write(row, col, linea[0])
            worksheet.write(row, col + 1, linea[1])
            worksheet.write(row, col + 2, linea[2])
            worksheet.write(row, col + 3, linea[3])
            row += 1

        workbook.close()
        conexion_obj.close()
        window['menGenRep'].update('Reporte Generado y guardado en el escritorio')
        window['diaInicio'].update('')
        window['diaFin'].update('')

    if event == 'reg5':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '1'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'reg6' or event == 'reg7':
        window['errorRegEnt'].update('')
        window['errorRegSal'].update('')
        window['errorRegistrarincvac'].update('')
        window[f'-COL{layout}-'].update(visible=False)
        layout = '5'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'reg8':
        window['numregbus'].update('')
        window[f'-COL{layout}-'].update(visible=False)
        layout = '4'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'reg9':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '1'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'Registro de Ingresos/Egresos':
        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()
        cursor_obj.execute("""SELECT * FROM RECFECHA""")
        rec = cursor_obj.fetchone()
        if rec == None:
            window['menCalcNom'].update('Nunca se ha calculado la Nómina')
        else:
            window['menCalcNom'].update('Nómina calculada el: ' + rec[0])

        window[f'-COL{layout}-'].update(visible=False)
        layout = '10'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'regHor':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '12'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'conHor':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '13'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'IngEg':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '11'
        window[f'-COL{layout}-'].update(visible=True)
        window['menRegIngEg'].update('')

    if event == 'regHorComp':
        grupo = values['grupoHor']
        window['grupoHor'].update('')
        materia = values['materia']
        window['materia'].update('')
        profe = values['profe']
        window['profe'].update('')
        salon = values['salon']
        window['salon'].update('')
        horEnt = values['horEntrada']
        window['horEntrada'].update('')
        horSal = values['horSalida']
        window['horSalida'].update('')

        valHor = (grupo, materia, profe, salon, horEnt, horSal)

        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()

        if values['lunes']:
            cursor_obj.execute("""INSERT INTO HORARIOS VALUES(?, 'Lunes', ?, ?, ?, ?, ?)""", valHor)

        if values['martes']:
            cursor_obj.execute("""INSERT INTO HORARIOS VALUES(?, 'Martes', ?, ?, ?, ?, ?)""", valHor)

        if values['miercoles']:
            cursor_obj.execute("""INSERT INTO HORARIOS VALUES(?, 'Miércoles', ?, ?, ?, ?, ?)""", valHor)

        if values['jueves']:
            cursor_obj.execute("""INSERT INTO HORARIOS VALUES(?, 'Jueves', ?, ?, ?, ?, ?)""", valHor)

        if values['viernes']:
            cursor_obj.execute("""INSERT INTO HORARIOS VALUES(?, 'Viernes', ?, ?, ?, ?, ?)""", valHor)

        window['lunes'].update(False)
        window['martes'].update(False)
        window['miercoles'].update(False)
        window['jueves'].update(False)
        window['viernes'].update(False)

        conexion_obj.commit()
        conexion_obj.close()

    if event == 'buscargrupo':

        grupo = values['grupoBus']
        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()
        cursor_obj.execute("""SELECT * FROM HORARIOS WHERE Grupo = ? 
        ORDER BY CASE
        WHEN Dia = 'Lunes' THEN 1
        WHEN Dia = 'Martes' THEN 2
        WHEN Dia = 'Miércoles' THEN 3
        WHEN Dia = 'Jueves' THEN 4
        WHEN Dia = 'Viernes' THEN 5
        END ASC, time(Horario_Entrada) ASC""", (grupo,))
        val = cursor_obj.fetchone()
        while val is not None:
            lista = []
            for i in range(0, 7):
                lista.append(val[i])
            datos.append(lista)
            val = cursor_obj.fetchone()

        window['tablaHor'].update(values=datos)
        window.refresh()
        conexion_obj.commit()
        conexion_obj.close()

    if event == 'borGrupo':
        grupo = values['grupoBus']
        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()
        cursor_obj.execute("""DELETE FROM HORARIOS WHERE Grupo = ?""", (grupo,))
        conexion_obj.commit()
        conexion_obj.close()
        window['grupoBus'].update('')
        datos = []
        window['tablaHor'].update(values=datos)
        window.refresh()
        window[f'-COL{layout}-'].update(visible=False)
        layout = '9'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'borClase':
        grupo = values['grupoHorB']
        window['grupoHorB'].update('')

        if values['lunesB']:
            dia = 'Lunes'

        if values['martesB']:
            dia = 'Martes'

        if values['miercolesB']:
            dia = 'Miércoles'

        if values['juevesB']:
            dia = 'Jueves'

        if values['viernesB']:
            dia = 'Viernes'

        materia = values['materiaB']
        window['materiaB'].update('')
        profe = values['profeB']
        window['profeB'].update('')
        salon = values['salonB']
        window['salonB'].update('')
        horEnt = values['horEntradaB']
        window['horEntradaB'].update('')
        horSal = values['horSalidaB']
        window['horSalidaB'].update('')
        window['lunesB'].update(False)
        window['martesB'].update(False)
        window['miercolesB'].update(False)
        window['juevesB'].update(False)
        window['viernesB'].update(False)

        valBorrar = (grupo, dia, materia, profe, salon, horEnt, horSal)

        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()

        cursor_obj.execute("""DELETE FROM HORARIOS WHERE Grupo = ? AND Dia = ? AND Materia = ? AND Profesor = ?
        AND Salon = ? AND Horario_Entrada = ? AND Horario_Salida = ?""", valBorrar)

        datos = []
        window['tablaHor'].update(values=datos)
        window.refresh()

        cursor_obj.execute("""SELECT * FROM HORARIOS WHERE Grupo = ? 
                ORDER BY CASE
                WHEN Dia = 'Lunes' THEN 1
                WHEN Dia = 'Martes' THEN 2
                WHEN Dia = 'Miércoles' THEN 3
                WHEN Dia = 'Jueves' THEN 4
                WHEN Dia = 'Viernes' THEN 5
                END ASC, time(Horario_Entrada) ASC""", (grupo,))
        val = cursor_obj.fetchone()
        while val is not None:
            lista = []
            for i in range(0, 7):
                lista.append(val[i])
            datos.append(lista)
            val = cursor_obj.fetchone()

        window['tablaHor'].update(values=datos)
        window.refresh()

        conexion_obj.commit()
        conexion_obj.close()

    if event == 'modHor':

        grupoCopia = values['grupoHorB']
        window['grupoHorB'].update('')

        if values['lunesB']:
            diaCopia = 'Lunes'

        if values['martesB']:
            diaCopia = 'Martes'

        if values['miercolesB']:
            diaCopia = 'Miércoles'

        if values['juevesB']:
            diaCopia = 'Jueves'

        if values['viernesB']:
            diaCopia = 'Viernes'

        materiaCopia = values['materiaB']
        window['materiaB'].update('')
        profeCopia = values['profeB']
        window['profeB'].update('')
        salonCopia = values['salonB']
        window['salonB'].update('')
        horEntCopia = values['horEntradaB']
        window['horEntradaB'].update('')
        horSalCopia = values['horSalidaB']
        window['horSalidaB'].update('')
        window['lunesB'].update(False)
        window['martesB'].update(False)
        window['miercolesB'].update(False)
        window['juevesB'].update(False)
        window['viernesB'].update(False)

        grupo = listVal[0]
        dia = listVal[1]
        materia = listVal[2]
        profe = listVal[3]
        salon = listVal[4]
        horEnt = listVal[5]
        horSal = listVal[6]

        valMod = (grupoCopia, diaCopia, materiaCopia, profeCopia, salonCopia, horEntCopia, horSalCopia, grupo,
                  dia, materia, profe, salon, horEnt, horSal)

        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()

        cursor_obj.execute("""UPDATE HORARIOS SET Grupo = ?, Dia = ?, Materia = ?, Profesor = ?, Salon = ?, 
        Horario_Entrada = ?, Horario_Salida = ? WHERE Grupo = ? AND Dia = ? AND Materia = ? AND Profesor = ?
                AND Salon = ? AND Horario_Entrada = ? AND Horario_Salida = ?""", valMod)

        datos = []
        window['tablaHor'].update(values=datos)
        window.refresh()

        cursor_obj.execute("""SELECT * FROM HORARIOS WHERE Grupo = ? 
                        ORDER BY CASE
                        WHEN Dia = 'Lunes' THEN 1
                        WHEN Dia = 'Martes' THEN 2
                        WHEN Dia = 'Miércoles' THEN 3
                        WHEN Dia = 'Jueves' THEN 4
                        WHEN Dia = 'Viernes' THEN 5
                        END ASC, time(Horario_Entrada) ASC""", (grupoCopia,))
        val = cursor_obj.fetchone()
        while val is not None:
            lista = []
            for i in range(0, 7):
                lista.append(val[i])
            datos.append(lista)
            val = cursor_obj.fetchone()

        window['tablaHor'].update(values=datos)
        window.refresh()

        conexion_obj.commit()
        conexion_obj.close()

    if event == 'tablaHor':
        window['modClaseGrupo'].update(disabled=False)

    if event == 'modClaseGrupo':
        colselec = values['tablaHor'][0]

        window['lunesB'].update(False)
        window['martesB'].update(False)
        window['miercolesB'].update(False)
        window['juevesB'].update(False)
        window['viernesB'].update(False)
        grupo = values['grupoBus']

        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()

        cursor_obj.execute("""SELECT * FROM HORARIOS WHERE Grupo = ? 
                ORDER BY CASE
                WHEN Dia = 'Lunes' THEN 1
                WHEN Dia = 'Martes' THEN 2
                WHEN Dia = 'Miércoles' THEN 3
                WHEN Dia = 'Jueves' THEN 4
                WHEN Dia = 'Viernes' THEN 5
                END ASC, time(Horario_Entrada) ASC""", (grupo,))

        for i in range(0, colselec + 1):
            data = cursor_obj.fetchone()

        window['grupoHorB'].update(data[0])
        if data[1] == 'Lunes':
            window['lunesB'].update(True)
        if data[1] == 'Martes':
            window['martesB'].update(True)
        if data[1] == 'Miércoles':
            window['miercolesB'].update(True)
        if data[1] == 'Jueves':
            window['juevesB'].update(True)
        if data[1] == 'Viernes':
            window['viernesB'].update(True)
        window['materiaB'].update(data[2])
        window['profeB'].update(data[3])
        window['salonB'].update(data[4])
        window['horEntradaB'].update(data[5])
        window['horSalidaB'].update(data[6])

        listVal = [data[0], data[1], data[2], data[3], data[4], data[5], data[6]]

        conexion_obj.commit()
        conexion_obj.close()
        window['colbus'].update(visible=True)

    if event == 'Solicitar Tutoría':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '14'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'busTutGrupo':

        grupo = values['grupoHorT']

        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()

        cursor_obj.execute("""SELECT * FROM HORARIOS WHERE Grupo = ? 
                        ORDER BY CASE
                        WHEN Dia = 'Lunes' THEN 1
                        WHEN Dia = 'Martes' THEN 2
                        WHEN Dia = 'Miércoles' THEN 3
                        WHEN Dia = 'Jueves' THEN 4
                        WHEN Dia = 'Viernes' THEN 5
                        END ASC, time(Horario_Entrada) ASC""", (grupo,))

        val = cursor_obj.fetchone()
        while val is not None:
            lista = []
            for i in range(0, 7):
                lista.append(val[i])
            datTut.append(lista)
            val = cursor_obj.fetchone()

        window['tablaTut'].update(values=datTut)
        window.refresh()

        conexion_obj.commit()
        conexion_obj.close()

        window['grupoHorT'].update(disabled=False)
        window['alumno'].update(disabled=False)
        window['materiaT'].update(disabled=False)
        window['profeT'].update(disabled=False)
        window['salonT'].update(disabled=False)
        window['horEntradaT'].update(disabled=False)
        window['horSalidaT'].update(disabled=False)
        window['lunesT'].update(disabled=False)
        window['martesT'].update(disabled=False)
        window['miercolesT'].update(disabled=False)
        window['juevesT'].update(disabled=False)
        window['viernesT'].update(disabled=False)

        window['colTabTut'].update(visible=True)

    if event == 'regTut':

        grupo = values['grupoHorT']
        window['grupoHorT'].update('')
        alumno = values['alumno']
        window['alumno'].update('')
        materia = values['materiaT']
        window['materiaT'].update('')
        profe = values['profeT']
        window['profeT'].update('')
        salon = values['salonT']
        window['salonT'].update('')
        horEnt = values['horEntradaT']
        window['horEntradaT'].update('')
        horSal = values['horSalidaT']
        window['horSalidaT'].update('')

        tut = 'Tutoría ' + materia + ': ' + alumno

        valHor = (grupo, tut, profe, salon, horEnt, horSal)

        conexion_obj = sqlite3.connect(r"C:/NOMEDU/nomedubd.db")
        cursor_obj = conexion_obj.cursor()

        if values['lunesT']:
            cursor_obj.execute("""INSERT INTO HORARIOS VALUES(?, 'Lunes', ?, ?, ?, ?, ?)""", valHor)

        if values['martesT']:
            cursor_obj.execute("""INSERT INTO HORARIOS VALUES(?, 'Martes', ?, ?, ?, ?, ?)""", valHor)

        if values['miercolesT']:
            cursor_obj.execute("""INSERT INTO HORARIOS VALUES(?, 'Miércoles', ?, ?, ?, ?, ?)""", valHor)

        if values['juevesT']:
            cursor_obj.execute("""INSERT INTO HORARIOS VALUES(?, 'Jueves', ?, ?, ?, ?, ?)""", valHor)

        if values['viernesT']:
            cursor_obj.execute("""INSERT INTO HORARIOS VALUES(?, 'Viernes', ?, ?, ?, ?, ?)""", valHor)

        conexion_obj.commit()
        conexion_obj.close()

        window['alumno'].update(disabled=True)
        window['materiaT'].update(disabled=True)
        window['profeT'].update(disabled=True)
        window['salonT'].update(disabled=True)
        window['horEntradaT'].update(disabled=True)
        window['horSalidaT'].update(disabled=True)

        window['lunesT'].update(disabled=True)
        window['martesT'].update(disabled=True)
        window['miercolesT'].update(disabled=True)
        window['juevesT'].update(disabled=True)
        window['viernesT'].update(disabled=True)

        window['lunesT'].update(False)
        window['martesT'].update(False)
        window['miercolesT'].update(False)
        window['juevesT'].update(False)
        window['viernesT'].update(False)
        window['colTabTut'].update(False)

        window['menTut'].update('Registro Exitoso')

    if event == 'reg10':
        window['menGenRep'].update('')
        window[f'-COL{layout}-'].update(visible=False)
        layout = '1'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'reg11':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '10'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'reg12':
        window[f'-COL{layout}-'].update(visible=False)
        layout = '9'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'reg13':
        window['grupoBus'].update('')
        datos = []
        window['tablaHor'].update(values=datos)
        window['colbus'].update(visible=False)
        window.refresh()
        window[f'-COL{layout}-'].update(visible=False)
        layout = '9'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'reg14':
        window['grupoHorT'].update('')
        window['materiaT'].update('')
        window['profeT'].update('')
        window['salonT'].update('')
        window['horEntradaT'].update('')
        window['horSalidaT'].update('')
        window['lunesT'].update(False)
        window['martesT'].update(False)
        window['miercolesT'].update(False)
        window['juevesT'].update(False)
        window['viernesT'].update(False)
        window['colTabTut'].update(visible=False)
        window['menTut'].update('')
        window[f'-COL{layout}-'].update(visible=False)
        layout = '1'
        window[f'-COL{layout}-'].update(visible=True)

    if event == 'salir2' or event == 'salir3' or event == 'salir4' or event == 'salir5' \
            or event == 'salir6' or event == 'salir7' or event == 'salir8' or event == 'salir9' \
            or event == 'salir10' or event == 'salir11' or event == 'salir12' or event == 'salir13' \
            or event == 'salir14':
        break

window.close()
