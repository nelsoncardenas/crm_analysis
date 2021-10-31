from openpyxl import load_workbook, Workbook


class TeamEvaluator:
    """A module to perform an architect team evaluation."""

    def __init__(self, conf) -> None:
        """Initializes the instance.

        Args:
            conf ([type]): configuration dictionary.
        """
        self.path_input = conf["PATHS"]["INPUT"]
        self.path_equipo = conf["PATHS"]["EQUIPO"]
        self.path_output = conf["PATHS"]["OUTPUT"]

        self.texto_ejecutada = conf["TEXTO_EJECUTADA"]
        self.texto_programada = conf["TEXTO_PROGRAMADA"]
        self.estado_map = conf["ESTADO_MAP"]

        self.col_equipo = conf["COL_EQUIPO"]
        self.col_ejecucion = conf["COL_EJECUCION"]
        self.col_nombre = conf["COL_NOMBRE"]
        self.col_estado = conf["COL_ESTADO"]

        self.equipo_preventa_list = list()

        self.libro_input = load_workbook(self.path_input, data_only=True)
        self.libro_equipo = load_workbook(self.path_equipo, data_only=True)
        self.libro_output = Workbook()

        self.hoja_input = self.libro_input["Hoja1"]
        self.hoja_equipo = self.libro_equipo.active
        self.hoja_output = self.libro_output.active

        self.cont_errores = dict()
        self.cont_ejecutadas = dict()

        self.equipo_preventa_list, self.equipo_preventa_set = self._load_team()

        self.preventas_calificados = set()

    def evaluate(self):
        print("---------- Empezando el análisis ----------\n")
        print(
            "Advertencia, recuerda que el nombre de la persona en el archivo debe coincidir totalmente con el nombre en el CRM"
        )
        self._error_analysis()
        self._evaluate_scheduled_and_posponed()
        self._evaluate_execution()
        self._save_output_sheet()

    def _load_team(self):
        print("Loading team architects")
        equipo_preventa_list = []
        fila = 1
        while True:
            name = self.hoja_equipo.cell(row=fila, column=self.col_equipo).value
            if name is None:
                break
            equipo_preventa_list.append(name)
            fila += 1
        equipo_preventa_list.sort()
        equipo_preventa_set = set(equipo_preventa_list)
        print(f"El equipo de preventa tiene un tamaño de {len(equipo_preventa_set)}.")
        return equipo_preventa_list, equipo_preventa_set

    def _error_analysis(self):
        print("Generating data about errors.")
        fila = 2
        cell_to_validate = self.hoja_input.cell(row=fila, column=12).value
        while cell_to_validate != None:
            ejecucion = self.hoja_input.cell(row=fila, column=self.col_ejecucion).value
            ejecucion = " ".join(ejecucion.split()[:4])
            nombre = self.hoja_input.cell(row=fila, column=self.col_nombre).value
            if (
                ejecucion == self.texto_programada
                and nombre in self.equipo_preventa_set
            ):
                estado = self.estado_map[
                    self.hoja_input.cell(row=fila, column=self.col_estado).value
                ]
                nombre_estado = nombre + " + " + estado
                self.cont_errores[nombre_estado] = (
                    self.cont_errores.get(nombre_estado, 0) + 1
                )
            if ejecucion == self.texto_ejecutada and nombre in self.equipo_preventa_set:
                self.cont_ejecutadas[nombre] = self.cont_ejecutadas.get(nombre, 0) + 1
            fila += 1
            cell_to_validate = self.hoja_input.cell(row=fila, column=12).value
        print(fila - 1, "líneas analizadas.")

    def _evaluate_scheduled_and_posponed(self):
        print("Evaluating schedule and posponed proposals.")
        self.hoja_output.cell(row=1, column=1).value = "Nombre"
        self.hoja_output.cell(row=1, column=2).value = "Puntaje"
        self.hoja_output.cell(row=1, column=3).value = "Texto Errores"

        errores_keys = list(self.cont_errores.keys())
        errores_keys.sort()
        fila = 0
        while fila < len(errores_keys):
            architect_errors = 0
            nombre, _, num_errores_actual, texto_errores = self._get_error_info(
                errors_counter=self.cont_errores,
                name_error=errores_keys[fila],
                texto_errores="",
            )
            architect_errors += num_errores_actual
            if (
                fila < len(errores_keys) - 1
                and nombre == errores_keys[fila + 1].split(" + ")[0]
            ):  # si en la siguiente posición es el mismo preventa
                nombre, _, num_errores_actual, texto_errores = self._get_error_info(
                    self.cont_errores, errores_keys[fila + 1], texto_errores
                )
                architect_errors = architect_errors + num_errores_actual
                fila += 1
            puntaje = self._get_score(architect_errors)
            self.preventas_calificados.update([nombre])
            self.hoja_output.cell(
                row=self.equipo_preventa_list.index(nombre) + 2, column=2
            ).value = puntaje
            self.hoja_output.cell(
                row=self.equipo_preventa_list.index(nombre) + 2, column=3
            ).value = texto_errores
            fila += 1

    def _get_error_info(
        self, errors_counter: dict, name_error: str, texto_errores: str = ""
    ):
        nombre, error = name_error.split(" + ")
        num_errores_actual = errors_counter[name_error]
        texto_errores = f"{texto_errores} Tiene {num_errores_actual} preventas {error} incorrectamente."
        return nombre, error, num_errores_actual, texto_errores

    def _get_score(self, architect_errors: int):
        if architect_errors == 0:
            return 1
        elif 1 <= architect_errors <= 2:
            return 0.5
        else:
            return 0

    def _evaluate_execution(self):
        for nombre in self.equipo_preventa_list:
            self.hoja_output.cell(
                row=self.equipo_preventa_list.index(nombre) + 2, column=1
            ).value = nombre
            if nombre not in self.preventas_calificados:
                self.hoja_output.cell(
                    row=self.equipo_preventa_list.index(nombre) + 2, column=2
                ).value = 1
            if nombre not in self.cont_ejecutadas:
                self.hoja_output.cell(
                    row=self.equipo_preventa_list.index(nombre) + 2, column=2
                ).value = 0
                # Si está vacía la celda
                celda_errores = str(
                    self.hoja_output.cell(
                        row=self.equipo_preventa_list.index(nombre) + 2, column=3
                    ).value
                )
                if celda_errores == "None":
                    self.hoja_output.cell(
                        row=self.equipo_preventa_list.index(nombre) + 2, column=3
                    ).value = "No ha ejecutado preventas en las últimas dos semanas"
                else:
                    self.hoja_output.cell(
                        row=self.equipo_preventa_list.index(nombre) + 2, column=3
                    ).value = (
                        celda_errores
                        + " No han ejecutado preventas en las últimas dos semanas."
                    )

    def _save_output_sheet(self):
        self.libro_output.save(self.path_output)
        print("Programa ejecutado correctamente.")
