import numpy as np
import openpyxl
import os
import datetime
import warnings


def read_readcard(folder, name="readcard.csv"):
    with open(f'{folder}/{name}', 'r') as conn:
        course_table = conn.readlines()
    header = course_table[0].strip('\n').split('\t')
    result_table = []
    for line_id, line in enumerate(course_table[1:]):
        result_table.append(line.strip('\n').split('\t'))
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
    text_output = " id       punched    cp      dead     max time    cumulative   real max        in /\n" \
                  "                             time       w d t         d t        time          out \n"
    for line_id, line in enumerate(team_raw_table):
        text_output += f"{line['cp_id']:>6}    {line['time'].strftime('%H:%M:%S')}"  # 18 signs I think

        if line_id in data_table['print_on']:
            id1 = np.where(data_table['print_on'] == line_id)[0][0]
            rline = data_table[id1]
            if rline['found_point'] is False:
                raise Exception("This should be impossible")
            text_output += f"   {rline['cp_number']:>2}    {rline['dead_time']}    " \
                           f"{rline['maximum_time_wdt'].strftime('%H:%M:%S')}  +  " \
                           f"{rline['cumulative_dead_time']}  =  {rline['maximum_time'].strftime('%H:%M:%S')}    " \
                           f"{'outside' if rline['exceeded_maximum_time'] else 'inside':>8}\n"
        else:
            text_output += '\n'
        if str(line_id) + "+" in data_table['print_on']:
            ids = np.where(data_table['print_on'] == str(line_id) + "+")[0]
            for id1 in ids:
                rline = data_table[id1]
                text_output += f"{'-' * 3} missing cp {'-' * 3}"
                text_output += f"   {rline['cp_number']:>2}    {'-' * 56}\n"

    text_output += f"\nNumber of valid control points: {number_of_cp}"
    text_output += f"\nfinal cumulative dead time: {final_cumulative_dead_time}"
    text_output += f"\nTime withoud dead time: {finish_time - start_time}"
    text_output += f"\nTotal time: {finish_time - start_time - final_cumulative_dead_time}"

    return text_output


def convert_from_timedelta_to_time(timedelta_obj):
    return datetime.time(timedelta_obj.seconds // 3600, timedelta_obj.seconds % 3600 // 60,
                         timedelta_obj.seconds % 60)


def recalculate_results(folder='track_day1_example', track_csv_separator=','):
    workbook = openpyxl.load_workbook(filename=f"{folder}/results_input.xlsx")  # load excel file
    sheet = workbook.active  # open workbook
    excel_row_index = 2  # Start with second line

    readcard_table = read_readcard(folder)

    while True:  # While loop over all rows (teams) in results_input.xlsx
        team_number = sheet[f'A{excel_row_index}'].value

        if team_number == 'STOP' or team_number is None:
            break
        team_siid = int(sheet[f'B{excel_row_index}'].value)
        print(f'\n{"-" * 83}\n{"-" * 83}\n\n\tTeam number: {team_number}\n')

        # if not os.path.isfile(f'kljucki/{team_number}.txt'):
        #     print(f'team {team_number} is not in the folder yet.')
        #     excel_row_index += 1
        #     continue
        # with open(f'kljucki/{team_number}.txt', 'r') as conn:
        #     text = conn.read()
        # team_raw_table = eval(text)
        # team_raw_table = [(i[0], datetime.datetime.strptime(i[1], '%H:%M:%S')) for i in team_raw_table]
        # team_raw_table = np.array(team_raw_table, dtype=[('cp_id', object), ('time', object)])
        team_raw_table = get_team_raw_table(team_siid, readcard_table)
        category = 100 * (int(team_number) // 100)

        with open(f'{folder}/{category}.csv', 'r') as conn:
            text = conn.readlines()
        course_table = []
        for line_id, line in enumerate(text):
            if line[0] == "#":
                continue
            course_table.append(line.strip('\n').split(track_csv_separator))
            course_table[-1][0] = int(course_table[-1][0])
            time_list = course_table[-1][1].split(":")
            # course_table[line_id][1] = datetime.datetime.strptime(course_table[line_id][1], '%H:%M:%S').time()
            course_table[-1][1] = datetime.timedelta(hours=float(time_list[0]), minutes=float(time_list[1]),
                                                          seconds=float(time_list[2]))

        data_table = np.zeros(len(course_table), dtype=[('dead_time', object),
                                                        ('cumulative_dead_time', object),
                                                        ('maximum_time', object),
                                                        ('found_point', bool),
                                                        ('exceeded_maximum_time', bool),
                                                        ('print_on', object),
                                                        ('cp_number', int),
                                                        ('maximum_time_wdt', object)  # without dead time
                                                        ]
                              )

        if team_raw_table[0]['cp_id'] != 'START':
            raise Exception(f'Expected that first point is "START" and not {team_raw_table[0][0]}')
        start_time = team_raw_table[0]['time']

        if team_raw_table[-1]['cp_id'] != 'FINISH':
            raise Exception(f'Expected that first point is "FINISH" and not {team_raw_table[-1][0]}')
        finish_time = team_raw_table[-1]['time']

        previous_team_id = 0  # this is here to check if order of control points is okay
        for id1, cp in enumerate(course_table):
            # 3 things that need to be done in this for loop:
            #   1. calculate maximum time (with dead time) for each control point
            #   2. find dead time for each control point
            #   3. check if control point was taken in right order

            cp_id = cp[0]
            time_diff: datetime.timedelta = cp[1]

            data_table['cp_number'][id1] = cp[2]

            if id1 == 0:
                data_table['maximum_time_wdt'][id1] = start_time + time_diff
                data_table['cumulative_dead_time'] = datetime.timedelta(0)
            else:
                data_table['maximum_time_wdt'][id1] = start_time + time_diff
                data_table['cumulative_dead_time'][id1] = data_table['cumulative_dead_time'][id1 - 1] + \
                                                          data_table['dead_time'][id1 - 1]

            data_table['maximum_time'][id1] = start_time + data_table['cumulative_dead_time'][id1] + time_diff
            matched_ids = np.where(team_raw_table['cp_id'] == cp_id)
            matched_ids = matched_ids[0]

            if len(matched_ids) == 0:
                data_table['dead_time'][id1] = datetime.timedelta()
                data_table['found_point'][id1] = False
                data_table['print_on'][id1] = str(previous_team_id) + "+"
                continue

            if not min(matched_ids) > previous_team_id:
                raise Exception('Order of control points is not correct!')
            previous_team_id = max(matched_ids)

            if len(matched_ids) == 1:
                data_table['dead_time'][id1] = datetime.timedelta()
            elif len(matched_ids) == 2:
                dead_time_start_id = matched_ids[0]
                dead_time_finish_id = matched_ids[1]
                if abs(dead_time_start_id - dead_time_finish_id) != 1:
                    warnings.warn(f'Control points for  dead time should be one after another but they are not! check!')
                if 'mrtvi' not in cp:
                    warnings.warn(
                        f'There is dead time where it was not supposed to happen (cp id = {cp[0]}, cp {cp[2]}')
                data_table['dead_time'][id1] = team_raw_table['time'][dead_time_finish_id] - team_raw_table['time'][
                    dead_time_start_id]
            else:
                warnings.warn('Did not expect more than 2 records of the card!')
            data_table['found_point'][id1] = True
            data_table['exceeded_maximum_time'][id1] = team_raw_table['time'][matched_ids[0]] > \
                                                       data_table['maximum_time'][id1]
            data_table['print_on'][id1] = matched_ids[0]
        data_table['cumulative_dead_time'][-1] = data_table['cumulative_dead_time'][-2] + data_table['dead_time'][-2]
        final_cumulative_dead_time = data_table['cumulative_dead_time'][-1]
        valid_cp = np.logical_and(data_table['found_point'], np.logical_not(data_table['exceeded_maximum_time']))
        number_of_cp = np.sum(valid_cp)
        print(print_log(data_table, team_raw_table, start_time, finish_time, final_cumulative_dead_time, number_of_cp))

        # Ekipe SIID	Start	Cilj	Skupni cas (brez mrtvega)	Mrtvi cas opisno	mrtvi cas skupno	skupni cas final
        # prekoracitve casovnic	 Vrstni red Ktjev	Stevilo Ktjev	hitrostna cas
        sheet[f'C{excel_row_index}'].value = start_time.time()  # Start
        sheet[f'D{excel_row_index}'].value = finish_time.time()  # Finish
        sheet[f'E{excel_row_index}'].value = convert_from_timedelta_to_time(finish_time - start_time)  # tot time wdt
        # description of dead time
        sheet[f'G{excel_row_index}'].value = convert_from_timedelta_to_time(final_cumulative_dead_time)  # tot time wdt
        sheet[f'H{excel_row_index}'].value = convert_from_timedelta_to_time(
            finish_time - start_time - final_cumulative_dead_time)  # tot time wdt
        # number of exceded max time
        # order of CP
        sheet[f'K{excel_row_index}'].value = number_of_cp

        for i in range(len(valid_cp)):
            cell_name = f'{openpyxl.utils.cell.get_column_letter(i + 15)}{excel_row_index}'
            sheet[cell_name].value = '+' if valid_cp[i] else '-'
        excel_row_index += 1
    workbook.save(filename=f"{folder}/results_output.xlsx")


if __name__ == "__main__":
    recalculate_results()
