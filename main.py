import numpy as np
import openpyxl
import os
import datetime
from collections.abc import Iterable
from shutil import copy


def read_readcard(folder, name="readcard.csv", autorecover_name=False):
    filename = f'{folder}/{name}'
    if autorecover_name:
        file_list = os.listdir('datadump')
        file_list = sorted(file_list)
        file_list = [i for i in file_list if 'readcard_' in i]
        filename = f"datadump/" + file_list[-1]

        if os.path.isfile(f'{folder}/readcard_copy.csv'):
            os.remove(f'{folder}/readcard_copy.csv')
        copy(filename, f'{folder}/readcard_copy.csv')
    with open(filename, 'r', encoding='iso-8859-1') as conn:
        course_table = conn.readlines()
    header = course_table[0].strip('\n').split(';')
    result_table = []
    for line_id, line in enumerate(course_table[1:]):
        result_table.append(line.strip('\n').split(';'))
    max_len = max([len(i) for i in result_table])
    header = header[:max_len]
    for id1, line in enumerate(result_table):
        result_table[id1] += [''] * (max_len - len(line))
    head1 = [(i, "<40U") for i in header]
    tab1 = [tuple(i) for i in result_table]
    array_table = np.array(tab1, dtype=head1)
    return array_table


def get_team_raw_table(siid, readcard_table):
    id_in_readcard_table = np.where(readcard_table['SIID'] == str(siid))[0]
    if len(id_in_readcard_table) == 0:
        return False
    elif len(id_in_readcard_table) > 1:
        raise Exception("There should be only one ID!!")
    line = readcard_table[id_in_readcard_table[0]]

    start_time = line['Start time']
    finish_time = line['Finish time']

    record_number = int(line['No. of records'])

    output_array = [('START', datetime.datetime.strptime(start_time.replace(" ", ''), '%H:%M:%S'))]
    for i in range(1, record_number + 1):
        cp_id = int(line[f"Record {i} CN"])
        cp_time = datetime.datetime.strptime(line[f"Record {i} time"].replace(" ", ''), '%H:%M:%S')
        output_array.append((cp_id, cp_time))
    output_array.append(('FINISH', datetime.datetime.strptime(finish_time.replace(" ", ''), '%H:%M:%S')))
    output_array = np.array(output_array, dtype=[('cp_id', object), ('time', object)])
    return output_array


def print_log(data_table, team_raw_table, start_time, finish_time, final_cumulative_dead_time, number_of_cp):
    text_output = " id       punched   | cp      dead     max time    cumulative   real max        in /\n" \
                  "                    |         time       w d t         d t        time          out \n"
    for line_id, line in enumerate(team_raw_table):
        text_output += f"{line['cp_id']:>6}    {line['time'].strftime('%H:%M:%S')}"  # 18 signs I think

        if line_id in data_table['print_on']:
            id1 = np.where(data_table['print_on'] == line_id)[0][0]
            rline = data_table[id1]
            if rline['found_point'] is False:
                raise Exception("This should be impossible")
            text_output += f"  | {rline['cp_number']:>2}    {rline['dead_time']}    " \
                           f"{rline['maximum_time_wdt'].strftime('%H:%M:%S')}  +  " \
                           f"{rline['cumulative_dead_time']}  =  {rline['maximum_time'].strftime('%H:%M:%S')}    " \
                           f"{'outside' if rline['exceeded_max_time'] else 'inside':>8}\n"
        else:
            text_output += '  |\n'
        if str(line_id) + "+" in data_table['print_on']:
            ids = np.where(data_table['print_on'] == str(line_id) + "+")[0]
            for id1 in ids:
                rline = data_table[id1]
                text_output += f"{' ' * 18}  | {rline['cp_number']:>2}"
                text_output += f"    <<{'-' * 20} missing cp {rline['cp_number']:>2} {'-' * 19}\n"

    text_output += f"\nNumber of valid control points: {number_of_cp}"
    text_output += f"\nFinal cumulative dead time: {final_cumulative_dead_time}"
    text_output += f"\nTime without dead time: {finish_time - start_time}"
    text_output += f"\nTotal time: {finish_time - start_time - final_cumulative_dead_time}"

    return text_output


def convert_from_timedelta_to_time(timedelta_obj):
    return datetime.time(timedelta_obj.seconds // 3600, timedelta_obj.seconds % 3600 // 60,
                         timedelta_obj.seconds % 60)


def read_course_table(folder, category, track_csv_separator):
    with open(f'{folder}/{category}.csv', 'r') as conn:
        text = conn.readlines()
    course_table = []
    for line_id, line in enumerate(text):
        if line[0] == "#":
            continue
        line_split = line.strip('\n').split(track_csv_separator)
        line_split = [i for i in line_split if i != '']
        if len(line_split) == 0:
            continue
        cp_id = int(line_split[0])
        time_list = line_split[1].split(":")
        max_time_dif = datetime.timedelta(hours=float(time_list[0]), minutes=float(time_list[1]),
                                          seconds=float(time_list[2]))
        cp_number = int(line_split[2])
        additional_arguments = line_split[3:]
        course_table.append((cp_id, max_time_dif, cp_number, additional_arguments))
    course_table = np.array(course_table, dtype=[('cp_id', int),
                                                 ('max_time_dif', object),
                                                 ('cp_number', int),
                                                 ('additional_args', object)])
    return course_table


def calculate_results_for_one_team(team_siid, team_number, folder, readcard_table, track_csv_separator,
                                   comply_with_deadtime_tag=True):
    local_log = f'\tTeam number: {team_number}\n'
    dead_time_text = ""
    warning_text = ""
    error_text = ""

    if not os.path.isdir(f'{folder}/logs'):
        os.mkdir(f'{folder}/logs')

    team_raw_table = get_team_raw_table(team_siid, readcard_table)
    if team_raw_table is False:
        error_text = "No data from this card yet"
        local_log += '\n There is no data from this team!'
        return local_log, error_text, warning_text, dead_time_text, None, None, None, None, None, None, None, None

    category = 100 * (int(team_number) // 100)

    course_table = read_course_table(folder, category, track_csv_separator)
    # Numpy array with columns: control point id, maximum time to get to control point, number of control point and
    # additional arguments (cp_id, max_time_dif, cp_number, additional_args)

    data_table = np.zeros(len(course_table), dtype=[('dead_time', object),
                                                    ('cumulative_dead_time', object),
                                                    ('maximum_time', object),
                                                    ('found_point', bool),
                                                    ('exceeded_max_time', bool),
                                                    ('print_on', object),
                                                    ('cp_number', int),
                                                    ('maximum_time_wdt', object)  # without dead time
                                                    ])

    if team_raw_table[0]['cp_id'] != 'START':
        raise Exception(f'Expected that first point is "START" and not {team_raw_table[0][0]}')
    if team_raw_table[-1]['cp_id'] != 'FINISH':
        raise Exception(f'Expected that first point is "FINISH" and not {team_raw_table[-1][0]}')

    start_time = team_raw_table[0]['time']
    finish_time = team_raw_table[-1]['time']

    speed_trial_start = False
    speed_trial_finish = False

    previous_matched_raw_team_table_id = 0  # this is here to check if order of control points is okay
    for id1, cp in enumerate(course_table):
        # 3 things that need to be done in this for loop:
        #   1. calculate maximum time (with dead time) for each control point
        #   2. find dead time for each control point
        #   3. check if control point was taken in right order

        data_table['cp_number'][id1] = cp['cp_number']

        if id1 == 0:  # First iteration
            data_table['maximum_time_wdt'][id1] = start_time + cp['max_time_dif']
            data_table['cumulative_dead_time'] = datetime.timedelta(0)
        else:
            data_table['maximum_time_wdt'][id1] = start_time + cp['max_time_dif']
            data_table['cumulative_dead_time'][id1] = data_table['cumulative_dead_time'][id1 - 1] + \
                data_table['dead_time'][id1 - 1]

        data_table['maximum_time'][id1] = start_time + data_table['cumulative_dead_time'][id1] + cp['max_time_dif']
        matched_ids = np.where(team_raw_table['cp_id'] == cp['cp_id'])
        matched_ids = matched_ids[0]

        if len(matched_ids) == 0:  # There is no records of this control point on the card
            data_table['dead_time'][id1] = datetime.timedelta()
            data_table['found_point'][id1] = False
            data_table['print_on'][id1] = str(previous_matched_raw_team_table_id) + "+"  # After previous match
            continue

        if not min(matched_ids) > previous_matched_raw_team_table_id:
            # There was already a record of control point recorded after current matched cp. This means that order
            # wasn't correct
            error_text += "Order of control points is not correct! "
        previous_matched_raw_team_table_id = max(matched_ids)

        if len(matched_ids) == 1:
            # This means that control point was recorded only once so there is no problem with dead time
            data_table['dead_time'][id1] = datetime.timedelta(0)
        elif len(matched_ids) == 2:
            # This means that control point was recorded twice. The first control point is now start of dead time
            # and second cp is the end.
            dead_time_start_id = matched_ids[0]
            dead_time_finish_id = matched_ids[1]
            if abs(dead_time_start_id - dead_time_finish_id) != 1:
                warning_text += f'Control points for dead time should be one after another' \
                                f' but they are not! check!'
            if 'mrtvi_cas' not in cp['additional_args']:  # To be sure that dead time is given only on control points
                # that are meant for it we introduce a keyword 'mrtvi_cas' in order to explicitly hint on which cps we
                # expect dead time. Otherwise we get a warning in the table.
                warning_text += f'There is dead time where it was not supposed to happen' \
                                f' (cp id = {cp[0]}), cp {cp[2]}'

            if comply_with_deadtime_tag and ('mrtvi_cas' not in cp['additional_args']):
                    data_table['dead_time'][id1] = datetime.datetime(0, 0, 0)
            else:
                data_table['dead_time'][id1] = team_raw_table['time'][dead_time_finish_id] - team_raw_table['time'][
                    dead_time_start_id]
                dead_time_text += f"CP{data_table['cp_number'][id1]} (id {cp['cp_id']}): " \
                                  f"{data_table['dead_time'][id1]}, "
        else:
            warning_text += 'Did not expect more than 2 records of the card!'

        data_table['found_point'][id1] = True  # There is at least one record of this control point
        exceeded_max_time: bool = team_raw_table['time'][matched_ids[0]] > data_table['maximum_time'][id1]
        data_table['exceeded_max_time'][id1] = exceeded_max_time
        data_table['print_on'][id1] = matched_ids[0]

        if 'hitrostna_start' in cp['additional_args']:
            speed_trial_start = team_raw_table['time'][matched_ids[-1]]
        if 'hitrostna_cilj' in cp['additional_args']:
            speed_trial_finish = team_raw_table['time'][matched_ids[0]]

    data_table['cumulative_dead_time'][-1] = data_table['cumulative_dead_time'][-2] + data_table['dead_time'][-2]
    final_cumulative_dead_time = data_table['cumulative_dead_time'][-1]
    valid_cp = np.logical_and(data_table['found_point'], np.logical_not(data_table['exceeded_max_time']))
    valid_cp_num = np.sum(valid_cp)

    local_log += print_log(data_table, team_raw_table, start_time, finish_time, final_cumulative_dead_time,
                           valid_cp_num)
    local_log += '\nTime trial time: '
    if type(speed_trial_start) != bool and type(speed_trial_finish) != bool:
        time_trial_return = convert_from_timedelta_to_time(speed_trial_finish - speed_trial_start)
        local_log += f'{time_trial_return} = ' \
                     f'{speed_trial_finish.strftime("%H:%M:%S")} - {speed_trial_start.strftime("%H:%M:%S")}\n'
        # time trial
    else:
        time_trial_return = f"{'no start ' * (speed_trial_start == False)}{'no finish' * (speed_trial_finish == False)}"
        local_log += f"{'no start ' * (speed_trial_start == False)}" \
                     f"{'no finish' * (speed_trial_finish == False)}\n"

    control_points_index_order = [i for i in data_table['print_on'] if not isinstance(i, Iterable)]
    if control_points_index_order == sorted(control_points_index_order):
        correct_order_text = "Correct order of control points"
    else:
        correct_order_text = 'Wrong order!!'
        error_text += "Order of control points is not correct! "

    with open(f'{folder}/logs/{team_number}.txt', 'w') as conn:
        conn.write(local_log)

    return local_log, error_text, warning_text, dead_time_text, valid_cp, valid_cp_num, final_cumulative_dead_time, \
        time_trial_return, data_table, correct_order_text, start_time, finish_time


def recalculate_results(folder='test_system', track_csv_separator=',', automatic_readcard_name=True,
                        readcard_filename='readcard.csv', comply_with_deadtime_tag=True):
    workbook = openpyxl.load_workbook(filename=f"{folder}/results_input.xlsx")  # load excel file
    sheet = workbook.active  # open workbook
    excel_row_index = 2  # Start with second line
    result_table_string = f'{" " * 96}KT\nteam   siid    |  start       finish    mrtvi cas |  skupni cas   #KT ' \
                          f' hitrostna{" " * 10} | 1  2  3  4  5  6  7  8  9 10 11 12 13 14 15 16 warnings\n'
    readcard_table = read_readcard(folder, readcard_filename, autorecover_name=automatic_readcard_name)

    global_log = ""

    while True:  # While loop over all rows (teams) in results_input.xlsx

        team_number = sheet[f'A{excel_row_index}'].value

        if team_number == 'STOP' or team_number is None:
            break
        team_siid = int(sheet[f'B{excel_row_index}'].value)
        global_log += f'\n{"-" * 83}\n{"-" * 83}\n\n'

        ret = calculate_results_for_one_team(team_siid, team_number, folder, readcard_table, track_csv_separator,
                                             comply_with_deadtime_tag)
        local_log, error_text, warning_text, dead_time_text, valid_cp, valid_cp_num, final_cumulative_dead_time, \
            time_trial_return, data_table, correct_order_text, start_time, finish_time = ret
        eri = excel_row_index
        sheet[f'N{eri}'].value = error_text
        if error_text == '':
            sheet[f'C{eri}'].value = start_time.time()  # Start
            sheet[f'D{eri}'].value = finish_time.time()  # Finish
            sheet[f'E{eri}'].value = convert_from_timedelta_to_time(finish_time - start_time)  # without dt
            sheet[f'F{eri}'].value = dead_time_text
            sheet[f'G{eri}'].value = convert_from_timedelta_to_time(final_cumulative_dead_time)  # tot dt
            sheet[f'H{eri}'].value = convert_from_timedelta_to_time(
                finish_time - start_time - final_cumulative_dead_time)  # tot time wdt
            sheet[f'I{eri}'].value = correct_order_text
            sheet[f'J{eri}'].value = sum(data_table['found_point']) - valid_cp_num  # # of exceeded max time
            sheet[f'K{eri}'].value = valid_cp_num
            sheet[f'L{eri}'].value = time_trial_return  # Time trial
            sheet[f'M{excel_row_index}'].value = warning_text
            time_trial_string = f"{str(sheet[f'L{eri}'].value):<18}"
            str1 = f"{team_number}    {team_siid}  |  {start_time.time()}    {finish_time.time()}  " \
                   f"{sheet[f'G{eri}'].value}  |  {sheet[f'H{eri}'].value}    " \
                   f"{sheet[f'K{eri}'].value:>2}    {time_trial_string}  |"
            for i in range(len(valid_cp)):
                cell_name = f'{openpyxl.utils.cell.get_column_letter(i + 15)}{eri}'
                sheet[cell_name].value = '+' if valid_cp[i] else '-'
                str1 += f" {'+' if valid_cp[i] else '-'} "
        else:
            str1 = f"{team_number}    {team_siid}  |  {error_text}  ||"
        result_table_string += str1 + f'  {warning_text}\n'
        global_log += local_log
        excel_row_index += 1
    print(global_log, '\n\n\n')
    print(result_table_string)
    with open(f'{folder}/logger.txt', 'w', encoding='UTF-8') as conn:
        conn.write(global_log)
    workbook.save(filename=f"{folder}/results_output.xlsx")


if __name__ == "__main__":
    recalculate_results(folder='test_system')
