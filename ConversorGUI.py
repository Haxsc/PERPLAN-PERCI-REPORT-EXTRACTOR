import flet as ft
import os
import time
from datetime import datetime, timedelta
import pandas as pd
from openpyxl import load_workbook
import re

def calculate_start_row_array(selected_time, start_row_base=16, interval_minutes=15, increment=103, num_values=7):
    base_time = datetime.strptime("00:00", "%H:%M")
    selected_time = datetime.strptime(selected_time, "%H:%M")
    difference_minutes = (selected_time - base_time).seconds // 60
    
    relative_row = difference_minutes // interval_minutes
    start_row = start_row_base + relative_row

    start_row_array = [start_row + i * increment for i in range(num_values)]
    return start_row_array

def filter_by_date_and_time(df, date_column, target_date, start_time, end_time):
    df[date_column] = pd.to_datetime(df[date_column], errors="coerce", format="%Y-%m-%d %H:%M:%S")
    
    target_date = pd.to_datetime(target_date , format='%d-%m-%Y').date()
    start_datetime = pd.to_datetime(f"{target_date} {start_time}")
    end_datetime = pd.to_datetime(f"{target_date} {end_time}")

    return df[
        (df[date_column] >= start_datetime) &
        (df[date_column] <= end_datetime)
    ]

nrows = 0 
def convert_to_excel(csv_file, output_folder, filter_date, start_hour, end_hour, date_column="horaDas"):
    global nrows
    df = pd.read_csv(csv_file)
    if filter_date:
        filtered_df = filter_by_date_and_time(df, date_column, filter_date, start_hour, end_hour)
        nrows = len(filtered_df)

    base_name = os.path.splitext(os.path.basename(csv_file))[0]
    output_file = os.path.join(output_folder, f"{base_name}_{filter_date}.xlsx")
    filtered_df.to_excel(output_file, index=False)

    return output_file

def process_configuration(config, output_folder,log):
    global nrows
    writer = pd.ExcelWriter(path=config["excel_target"], engine="openpyxl", mode="a", if_sheet_exists="overlay")
    temporary_files = []
    
    for group_key in ["files_to_process_group_a", "files_to_process_group_b"]:  
        for csv_file, filter_date in config[group_key]:
            if csv_file.endswith(".csv"):
                if filter_date == "empty":
                    temporary_files.append((group_key,"empty"))
                else: 
                    try:
                        xlsx_file = convert_to_excel(csv_file, output_folder, filter_date, config["start_hour"], config["end_hour"])
                        temporary_files.append((group_key, xlsx_file))
                    except Exception as e:
                        log(f"Error converting {csv_file}: {str(e)}")
                
    log("Transferindo dados para Excel...")
    
    data_frames_a = []
    data_frames_b = []
    

    
    for group_key, xlsx_file in temporary_files:
        if group_key == "files_to_process_group_a":
            if xlsx_file == "empty":
                data_frames_a.append("")
                continue
            for csv,_ in config["files_to_process_group_a"] :
                name_mov_a = os.path.basename(csv).split('.')[0]
                if name_mov_a in xlsx_file:
                    df = pd.read_excel(xlsx_file,usecols="D:AD", sheet_name="Sheet1", skiprows=1 , nrows=nrows,engine="openpyxl",header=None)
                    data_frames_a.append(df)
                    break
        elif group_key == "files_to_process_group_b":
            if xlsx_file == "empty":
                data_frames_b.append("")
                continue
            for csv,_ in config["files_to_process_group_b"] :
                name_mov_b = os.path.basename(csv).split('.')[0]
                if name_mov_b in xlsx_file:
                    df = pd.read_excel(xlsx_file,usecols="D:AD", sheet_name="Sheet1", skiprows=1 , nrows=nrows ,engine="openpyxl",header=None)
                    data_frames_b.append(df)
                    break


    for df, start_row, day_control in zip(data_frames_a, config["start_rows"], config["days_controls"]):
        if day_control["boolean"]:
            if config["name"] == "Período Diurno":
                df["Period"] = "Diurno"
            else:
                df["Period"] = "Noturno"

            df.to_excel(
                writer,
                sheet_name="Contagens A (EXCLUIR)",
                startrow=start_row,
                startcol=4,  # Dados começam na coluna 4
                header=False,
                index=False
            )
            
            pd.DataFrame(df["Period"]).to_excel(
                writer,
                sheet_name="Contagens A (EXCLUIR)",
                startrow=start_row,
                startcol=36,  # Coluna AK
                header=False,
                index=False
            )

    for df, start_row, day_control in zip(data_frames_b, config["start_rows"], config["days_controls"] ):
        if day_control["boolean"]:  # Apenas processa se o checkbox estiver marcado
            if config["name"] == "Período Diurno":
                df["Period"] = "Diurno"
            else:
                df["Period"] = "Noturno"
            
            pd.DataFrame(df["Period"]).to_excel(
                writer,
                sheet_name="Contagens B (EXCLUIR)",
                startrow=start_row,
                startcol=36,  # Coluna AK
                header=False,
                index=False
            )
            df.to_excel(
                writer,
                sheet_name="Contagens B (EXCLUIR)",
                startrow=start_row,
                startcol=4,
                header=False,
                index=False
            )

    data_value = config["days_controls"][0]["data"]
    date_object = datetime.strptime(data_value, "%d-%m-%Y")
    formatted_date = date_object.strftime("%d/%m/%Y")
    df = pd.DataFrame([[formatted_date]], columns=["Data"])
    df.to_excel(writer, sheet_name="Títulos", startrow=22, startcol=1, header=False, index=False)
    
    file_name = os.path.splitext(os.path.basename(config["excel_target"]))[0]
    df_file_name = pd.DataFrame([[file_name]], columns=["Arquivo"])
    df_file_name.to_excel(writer, sheet_name="Títulos", startrow=19, startcol=1, header=False, index=False)

    writer.close()

    # Ensure files are deleted only once
    unique_temp_files = set(file for _, file in temporary_files)
    for temp_file in unique_temp_files:
        if temp_file == "empty":
            continue
        try:
            os.remove(temp_file)
            print(f"Deleted temporary file: {temp_file}")
        except FileNotFoundError:
            print(f"Temporary file already deleted: {temp_file}")
        except Exception as e:
            print(f"Error deleting temporary file {temp_file}: {str(e)}")

def move_files_to_old_folder(configurations, old_folder):
    processed_files = set()  # Para rastrear arquivos únicos

    for config in configurations:
        for group_key in ["files_to_process_group_a", "files_to_process_group_b"]:
            if group_key in config:  # Garante que a chave exista no dicionário
                for csv_file, _ in config[group_key]:
                    if csv_file not in processed_files:  # Verifica se o arquivo já foi processado
                        processed_files.add(csv_file)  # Adiciona o arquivo ao conjunto
                        destination = os.path.join(old_folder, os.path.basename(csv_file))
                        # try:
                        #     if os.path.exists(destination):
                        #          os.remove(destination)
                        #     shutil.move(csv_file, old_folder)
                        #     print(f"Moved file to 'old': {csv_file}")
                        # except Exception as e:
                        #     print(f"Error moving {csv_file}: {str(e)}")
                        
def main(page: ft.Page):
    page.title = "Reports Converter"
    page.scroll = "adaptive"
    page.theme_mode = ft.ThemeMode.DARK
    page.window.width = 800        # Largura da janela
    page.window.height = 1000       # Altura da janela
    
    # Default values for configuration
    current_directory = os.path.dirname(os.path.abspath(__file__))
    output_folder = current_directory
    old_folder =os.path.join(current_directory, "old")
    excel_target = None
    days_controls = []

    button_target_excel = ft.ElevatedButton(
            "Selecione o Arquivo Excel",
            icon=ft.Icons.UPLOAD_FILE,
            on_click=lambda _: excel_file_picker.pick_files(),
            width=300
    )
    
    def pick_excel_file(e: ft.FilePickerResultEvent):
        nonlocal excel_target
        path = None
        if e.files:
            path = e.files[0].path  # Atualiza com o caminho do primeiro arquivo

        if path is None:
            button_target_excel.text = "Arquivo Não Selecionado"
            button_target_excel.icon = ft.Icons.ERROR_ROUNDED
        else:
            button_target_excel.text = e.files[0].name
            button_target_excel.icon = ft.Icons.CHECK_ROUNDED

            excel_target = path
        button_target_excel.update()
    
    excel_file_picker = ft.FilePicker(on_result=pick_excel_file )
    page.overlay.append(excel_file_picker)  # Adiciona o FilePicker ao overlay

    days_columns = ft.Container(
        content=ft.Column([]),
        padding=10,
        alignment=ft.alignment.center
    )
    

    def update_days_columns(e):
        # Obtém o número de dias do dropdown
        num_days = int(days_process.value)

        # Limpa a lista de controles e a interface
        days_controls.clear()
        new_controls = []

        # Gera os campos para cada dia
        for day in range(1, num_days + 1):
            checkbox = ft.Checkbox(label=f"Dia {day}", value=True)
            text_field = ft.TextField(
                label=f"Data do Dia {day}",
                value=(datetime.now() + timedelta(days=day - 1)).strftime("%d-%m-%Y"),
                width=200,
            )

            # Atualiza days_controls quando o checkbox é alterado
            def on_checkbox_change(event, day_index=day - 1):
                # Atualiza o valor do checkbox no days_controls
                days_controls[day_index]["boolean"] = event.control.value

                if not event.control.value:
                    # Se desmarcado, limpa o TextField correspondente
                    new_controls[day_index].controls[1].value = ""
                    new_controls[day_index].controls[1].update()
                else:
                    # Se marcado novamente, recalcula a data e reescreve no TextField
                    recalculate_dates(day_index)

            # Atualiza days_controls quando o TextField é alterado
            def on_textfield_change(event, day_index=day - 1):
                # Tenta validar a entrada como uma data completa no formato YYYY-MM-DD
                if is_valid_date_format(event.control.value):
                    # Atualiza o valor da data no days_controls
                    days_controls[day_index]["data"] = event.control.value
                    # Recalcula as datas dos dias subsequentes
                    recalculate_dates(day_index)

            # Recalcula as datas subsequentes
            def recalculate_dates(start_index):
                try:
                    # Obtém a data inicial do índice alterado
                    start_date = datetime.strptime(days_controls[start_index]["data"], "%d-%m-%Y")
                    for i in range(start_index, len(days_controls)):
                        # Apenas atualiza se o checkbox estiver marcado
                        if days_controls[i]["boolean"]:
                            new_date = start_date + timedelta(days=i - start_index)
                            days_controls[i]["data"] = new_date.strftime("%d-%m-%Y")
                            new_controls[i].controls[1].value = days_controls[i]["data"]  # Atualiza o TextField
                            new_controls[i].controls[1].update()
                except ValueError:
                    # Ignora se o campo não contém uma data válida
                    pass

            # Valida o formato da data
            def is_valid_date_format(date_str):
                # Valida o padrão DD-MM-YYYY usando regex
                pattern = r"^\d{2}-\d{2}-\d{4}$"
                if re.match(pattern, date_str):
                    try:
                        datetime.strptime(date_str, "%d-%m-%Y")
                        return True
                    except ValueError:
                        return False
                return False

            # Adiciona os callbacks aos controles
            checkbox.on_change = lambda event, idx=day - 1: on_checkbox_change(event, idx)
            text_field.on_change = lambda event, idx=day - 1: on_textfield_change(event, idx)

            # Armazena os controles em uma lista
            days_controls.append({"boolean": checkbox.value, "data": text_field.value})

            # Adiciona os controles à interface
            new_controls.append(
                ft.Row(
                    [checkbox, text_field],
                    spacing=10,
                    alignment=ft.MainAxisAlignment.START
                )
            )

        # Atualiza o conteúdo do container
        days_columns.content.controls = new_controls
        days_columns.update()
      
    days_process = ft.Dropdown(
        label="Quantidade de dias",
        options=[ft.dropdown.Option(str(i)) for i in range(1, 8)],
        border_color="white",
        width=190,
        on_change=update_days_columns  # Chama a função ao alterar o valor
    )

    def pick_mov_a_day_files(e: ft.FilePickerResultEvent):
        path = None
        if e.files:
            path = ", ".join(map(lambda f: f.path, e.files))

        if path is None:
            button_day_a.text = "Arquivo Não Selecionado"
            button_day_a.icon = ft.Icons.ERROR_ROUNDED
        else:
            button_day_a.text = e.files[0].name
            button_day_a.icon = ft.Icons.CHECK_ROUNDED

            report_daytime_a.clear()
            report_daytime_a.insert(0,path)

        button_day_a.update()

    def pick_mov_b_day_files(e: ft.FilePickerResultEvent):
        path = None
        if e.files:
            path = ", ".join(map(lambda f: f.path, e.files))

        if path is None:
            button_day_b.text = "Arquivo Não Selecionado"
            button_day_b.icon = ft.Icons.ERROR_ROUNDED
        else:
            button_day_b.text = e.files[0].name
            button_day_b.icon = ft.Icons.CHECK_ROUNDED

            report_daytime_b.clear()
            report_daytime_b.insert(1,path)
        button_day_b.update()

    def pick_mov_a_evening_files(e: ft.FilePickerResultEvent):
        path = None
        if e.files:
            path = ", ".join(map(lambda f: f.path, e.files))

        if path is None:
            button_evening_a.text = "Arquivo Não Selecionado"
            button_evening_a.icon = ft.Icons.ERROR_ROUNDED
        else:
            button_evening_a.text = e.files[0].name
            button_evening_a.icon = ft.Icons.CHECK_ROUNDED

            report_evening_a.clear()
            report_evening_a.insert(0,path)

        button_evening_a.update()

    def pick_mov_b_evening_files(e: ft.FilePickerResultEvent):
        path = None
        if e.files:
            path = ", ".join(map(lambda f: f.path, e.files))

        if path is None:
            button_evening_b.text = "Arquivo Não Selecionado"
            button_evening_b.icon = ft.Icons.ERROR_ROUNDED

        else:
            button_evening_b.text = e.files[0].name
            button_evening_b.icon = ft.Icons.CHECK_ROUNDED

            report_evening_b.clear()
            report_evening_b.insert(1,path)

        button_evening_b.update()

    # Criando FilePickers para cada caso
    file_picker_mov_a_day = ft.FilePicker(on_result=pick_mov_a_day_files)
    file_picker_mov_b_day = ft.FilePicker(on_result=pick_mov_b_day_files)
    file_picker_mov_a_evening = ft.FilePicker(on_result=pick_mov_a_evening_files)
    file_picker_mov_b_evening = ft.FilePicker(on_result=pick_mov_b_evening_files)

    # Adicionando os FilePickers ao overlay
    page.overlay.append(file_picker_mov_a_day)
    page.overlay.append(file_picker_mov_b_day)
    page.overlay.append(file_picker_mov_a_evening)
    page.overlay.append(file_picker_mov_b_evening)
    
    
    # Array combinado
    report_daytime_a = []
    report_daytime_b = []
    report_evening_a = []
    report_evening_b = []

    button_day_a = ft.ElevatedButton(
        "Selecione o Mov.A",
        icon=ft.Icons.UPLOAD_FILE,
        on_click=lambda _: file_picker_mov_a_day.pick_files(),
        width=300
    )

    button_day_b = ft.ElevatedButton(
        "Selecione o Mov.B",
        icon=ft.Icons.UPLOAD_FILE,
        on_click=lambda _: file_picker_mov_b_day.pick_files(),
        width=300
    )

    button_evening_a = ft.ElevatedButton(
        "Selecione o Mov.A",
        icon=ft.Icons.UPLOAD_FILE,
        on_click=lambda _: file_picker_mov_a_evening.pick_files(),
        width=300
    )

    button_evening_b = ft.ElevatedButton(
        "Selecione o Mov.B",
        icon=ft.Icons.UPLOAD_FILE,
        on_click=lambda _: file_picker_mov_b_evening.pick_files(),
        width=300
    )

    day_files_container = ft.Container(
        content=ft.Column(
            [
                ft.Text("Arquivos do periodo diurno:", size=16, weight="bold" ,text_align= ft.TextAlign.CENTER ),
                ft.Column(
                    [
                        button_day_a,
                    ],
                    spacing=10,
                ),
                ft.Column(
                    [
                        button_day_b,
                    ],
                    spacing=10,
                ),
            ],
            spacing=10,

        ),
        padding=10,
    )

    evening_files_container = ft.Container(
        content=ft.Column(
            [
                ft.Text("Arquivos do periodo noturno:", size=16, weight="bold",text_align= ft.TextAlign.CENTER),
                ft.Column(
                    [
                        button_evening_a,
                        button_evening_b,
                    ],
                    spacing=10,
                    alignment=ft.MainAxisAlignment.CENTER
                ),
            ],
            alignment=ft.MainAxisAlignment.CENTER,
            spacing=10,
        ),
        padding=10,
        border_radius=5,
    )
    
    end_time_madrugada = "05:45"
    start_time = ft.Text()
    end_time = ft.Text()
    start_time_night = "18:00"

    def handle_starttime(e):
        nonlocal end_time_madrugada
        start_time.value = time.strftime(
            "%H:%M", time.gmtime(int(e.data))
        )
        start_time_dt = datetime.strptime(start_time.value, "%H:%M")
        end_time_madrugada = (start_time_dt - timedelta(minutes=15)).strftime("%H:%M")
        button_start_diurno.text = start_time.value
        button_start_diurno.update()
        
    def handle_endtime(e):
        nonlocal start_time_night
        end_time.value = time.strftime(
            "%H:%M", time.gmtime(int(e.data) - 60 * 15)
        )
        start_time_dt = datetime.strptime(end_time.value , "%H:%M")
        start_time_night = (start_time_dt + timedelta(minutes=15)).strftime("%H:%M")
        button_end_diurno.text = (start_time_dt + timedelta(minutes=15)).strftime("%H:%M")
        button_end_diurno.update()
        print(end_time.value)


    def open_time_picker_diurno(e):
        page.open(
                ft.CupertinoBottomSheet(
                    ft.CupertinoTimerPicker(
                        value=60 * 60 * 6,
                        minute_interval=15,
                        mode=ft.CupertinoTimerPickerMode.HOUR_MINUTE,
                        on_change=handle_starttime,
                ),
                height=216,  # Altura do picker
                padding=ft.padding.only(top=6),
            )
        )
        if e.data == "":
            start_time.value = time.strftime("%H:%M", time.gmtime(int(60 * 60 * 6)))
            button_start_diurno.text = start_time.value
            button_start_diurno.update()    
        
    def open_time_picker_noturno(e):
        page.open(
                ft.CupertinoBottomSheet(
                    ft.CupertinoTimerPicker(
                        value=60 * 60 * 18,
                        minute_interval=15,
                        mode=ft.CupertinoTimerPickerMode.HOUR_MINUTE,
                        on_change=handle_endtime,
                ),
                height=216,  # Altura do picker
                padding=ft.padding.only(top=6),
            )
        )
        if e.data == "":
            end_time.value = time.strftime("%H:%M", time.gmtime(int(60 * 60 * 18) - 60 * 15))
            button_end_diurno.text = time.strftime("%H:%M", time.gmtime(int(60 * 60 * 18)))
            button_end_diurno.update()

    button_start_diurno = ft.ElevatedButton(
        "Início do periodo DIURNO",
        width=250,
        on_click=lambda e: open_time_picker_diurno(e)
    )
    button_end_diurno = ft.ElevatedButton(
        "Fim do periodo DIURNO",
        width=250,
        on_click=lambda e: open_time_picker_noturno(e)
    )


    log_output = ft.Text("Informacões seram geradas aqui...\n", width=600, height=400 ,text_align= ft.TextAlign.CENTER )

    def show_error_dialog(message):
        # Cria o AlertDialog
        error_dialog = ft.AlertDialog(
            title=ft.Text("Error"),
            content=ft.Text(message),
            actions=[
                ft.TextButton("OK", on_click=lambda _: page.close(error_dialog))  # Fecha o diálogo ao clicar
            ],
            modal=True
        )   

        # Abre o diálogo
        page.open(error_dialog)
        page.update()



    def log(message):
        log_output.value += f"{message}\n"
        page.update()

    def generate_date_range(num_days, days_controls):
        valid_dates = []

        # Localiza o primeiro dia marcado como True
        for i, day in enumerate(days_controls):
            if day["boolean"]:
                start = datetime.strptime(day["data"], "%d-%m-%Y")
                break
        else:
            # Se nenhum dia estiver marcado como True, retorna uma lista vazia
            return valid_dates

        # Gera a lista de datas válidas
        for i in range(num_days):
            current_date = (start + timedelta(days=i)).strftime("%d-%m-%Y")
            if i < len(days_controls) and days_controls[i]["boolean"]:
                valid_dates.append(current_date)
            else:
                valid_dates.append("empty")

        return valid_dates
    
    def reset_app():
        # Limpar variáveis globais
        nonlocal log_output, report_daytime_a, report_daytime_b, report_evening_a, report_evening_b, excel_target
        report_daytime_a.clear()
        report_daytime_b.clear()
        report_evening_a.clear()
        report_evening_b.clear()
        log_output.value = "Informacões seram geradas aqui...\n"
        excel_target = None

        # Resetar interface
        button_target_excel.text = "Selecione o Arquivo Excel"
        button_target_excel.icon = ft.Icons.UPLOAD_FILE
        button_target_excel.update()

        button_day_a.text = "Selecione o Mov.A"
        button_day_a.icon = ft.Icons.UPLOAD_FILE
        button_day_a.update()

        button_day_b.text = "Selecione o Mov.B"
        button_day_b.icon = ft.Icons.UPLOAD_FILE
        button_day_b.update()

        button_evening_a.text = "Selecione o Mov.A"
        button_evening_a.icon = ft.Icons.UPLOAD_FILE
        button_evening_a.update()

        button_evening_b.text = "Selecione o Mov.B"
        button_evening_b.icon = ft.Icons.UPLOAD_FILE
        button_evening_b.update()

        # Limpar colunas de dias
        days_columns.content.controls = []
        days_columns.update()

        # Resetar dropdown de dias
        days_process.value = None
        days_process.update()

        page.update()


    

    def run_script(e):
        try:
            if not excel_target:
                show_error_dialog("Arquivo do excel não selecionado")
                return
            if not days_process.value:
                show_error_dialog("Nenhum dia foi selecionado")
                return
            if not start_time.value or not end_time.value :
                if not start_time.value:
                    show_error_dialog("Horário de Início nao selecionado")
                    return
                if not end_time.value:
                    show_error_dialog("Horário de Fim nao selecionado")
                    return
            if not report_evening_a and not report_daytime_a and not report_evening_b and not report_daytime_b:  # Verifica se é None ou lista vazia
                show_error_dialog("Nenhum report foi selecionado")
                return 
            
            # Configuration
            MAINCONFIG = {
                "output_folder": output_folder,
                "old_folder": old_folder,
                "excel_target": excel_target,
                "days_process": int(days_process.value),
                "report_evening_a": report_evening_a,
                "report_evening_b": report_evening_b,
                "report_daytime_a": report_daytime_a,   
                "report_daytime_b": report_daytime_b,
            }
            DATES_TO_PROCESS = generate_date_range(MAINCONFIG["days_process"] , days_controls)
            
            CONFIGURATIONS = [
                            {
                                "name": "Madrugada",
                                "start_hour": "00:00",
                                "end_hour": end_time_madrugada,
                                "days_process": MAINCONFIG["days_process"],    
                                "excel_target": MAINCONFIG["excel_target"],
                                "start_rows": calculate_start_row_array("00:00"),
                                "days_controls": days_controls,
                                "move_files": False,
                                "files_to_process_group_a": [(file, date) for file in MAINCONFIG["report_evening_a"] for date in DATES_TO_PROCESS],
                                "files_to_process_group_b": [(file, date) for file in MAINCONFIG["report_evening_b"] for date in DATES_TO_PROCESS],
                            },
                            {
                                "name": "Período Diurno",
                                "start_hour": start_time.value,
                                "end_hour": end_time.value,
                                "days_process": MAINCONFIG["days_process"],
                                "excel_target": MAINCONFIG["excel_target"],
                                "start_rows": calculate_start_row_array(start_time.value),
                                "days_controls": days_controls,
                                "move_files": True,
                                "files_to_process_group_a": [(file, date) for file in MAINCONFIG["report_daytime_a"] for date in DATES_TO_PROCESS],
                                "files_to_process_group_b": [(file, date) for file in MAINCONFIG["report_daytime_b"] for date in DATES_TO_PROCESS],
                            },
                            {
                                "name": "Período Noturno",
                                "start_hour": start_time_night,
                                "end_hour": "23:45",
                                "days_process": MAINCONFIG["days_process"],
                                "excel_target": MAINCONFIG["excel_target"],
                                "start_rows": calculate_start_row_array(start_time_night),
                                "days_controls": days_controls,
                                "move_files": True,
                                "files_to_process_group_a": [(file, date) for file in MAINCONFIG["report_evening_a"] for date in DATES_TO_PROCESS],
                                "files_to_process_group_b": [(file, date) for file in MAINCONFIG["report_evening_b"] for date in DATES_TO_PROCESS],
                            },
                        ]
        

                
            os.makedirs(old_folder, exist_ok=True)
            for config in CONFIGURATIONS:
                log(f"Iniciando o processamento de {config['name']}...")
                process_configuration(config, output_folder, log)
                log(f"Processamento de {config['name']} concluído com sucesso.")

            move_files_to_old_folder(CONFIGURATIONS, old_folder)
            log("Script concluido com sucesso.")  # Green text


            time.sleep(5)
            reset_app()
            #page.window.close()  # Fecha o programa
        except Exception as ex:
            log(f"Error: {str(ex)}")  # Loga o erro na interface e no console


    # UI Layout
    page.add(
        ft.Container(
            content=ft.Column(
                [
                    ft.Container(
                        ft.Text("Reports Conversor", size=20, weight="bold", text_align="center"),
                        alignment=ft.alignment.center,
                    ),
                    ft.Container(
                        button_target_excel,
                        alignment=ft.alignment.center,
                        padding=20,
                     ),
                    ft.Container(
                        ft.Row(
                            [days_columns, days_process , ft.Column([button_start_diurno, button_end_diurno])],
                            spacing=20,
                            alignment=ft.MainAxisAlignment.CENTER,  # Centraliza os itens na linha
                        ),
                        alignment=ft.alignment.top_center
                    ),
                    ft.Row(
                        [
                            day_files_container,
                            evening_files_container,
                        ],
                        spacing=20,
                        alignment=ft.MainAxisAlignment.CENTER,  # Centraliza os itens na linha
                    ),
                    ft.Container(
                        ft.ElevatedButton("Iniciar Processamento ",icon = ft.Icons.CHECK_ROUNDED, on_click=run_script, width=400 , color=ft.Colors.GREEN),
                        ft.Text("Logs:", size=16, weight="bold"),
                        alignment=ft.alignment.center,
                    ),
                        ft.Container(
                            ft.Container(
                                log_output,
                                bgcolor=ft.Colors.GREY_900,
                                width=500,
                                border_radius=5,
                            ),
                            alignment=ft.alignment.center,
                        )
                ],
                spacing=10,
                alignment=ft.MainAxisAlignment.CENTER,  # Centraliza toda a coluna
            ),
            alignment=ft.alignment.top_center,
        ),
    )

ft.app(target=main)


