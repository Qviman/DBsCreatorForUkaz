from kivy.app import App
from kivy.lang import Builder
from kivy.config import Config

#Config.set('graphics', 'width', '428')
#Config.set('graphics', 'height', '926')
#Config.set('graphics', 'resizable', '1')
#Config.set('graphics', 'borderless', '0')
#Config.set('kivy', 'exit_on_escape', '1')

from kivy.utils import get_color_from_hex
from kivy.core.window import Window
Window.clearcolor = get_color_from_hex('FFFFFF')


from kivy.clock import Clock

import os
import sqlite3
import datetime
import pickle
from kivy.animation import Animation

import openpyxl
import sqlite3

from time import sleep, time
import cProfile, pstats, io
from pstats import SortKey


class DataBase():
    def get_day(d, m):
        try: 
            conn = sqlite3.connect('db/'+str(m)+'.db')
            cur = conn.cursor()
            cur.execute('''select * from Calendar where day = ?''', (str(float(d)), ))                      
            db = cur.fetchall()[0]
            conn.commit()
            conn.close()
            return db
        except:
            return None


class Mistakes_and_solutions():
    mistakes1 = ['\xa0', '¬',   '…', ' .', '.. ', '\n ', ' )', ', -', '[У', '\n[/b]', ' ...', '\n \n']
    solved1   = [' ',     '', '...',  '.',  '. ',  '\n',  ')',  ', ',  'У',   '[/b]', '...', '\n\n']
    
    mistakes2 = ['«[font=fonts/triodion_ucs]',  '«font=fonts/triodion_ucs]', '[ref=world][color=0000ff] ', '«[font=fonts/triodion_ucs][/font]»[font=fonts/triodion_ucs]', '[ref=world][color=0000ff] ', ' [/color][/ref]']
    solved2   = ['[font=fonts/triodion_ucs]«', '[font=fonts/triodion_ucs]«',  '[ref=world][color=0000ff]', '[font=fonts/triodion_ucs]«', '[ref=world][color=0000ff]', '[/color][/ref]']

    mistakes_2 = ['[/font]...»', '[/font]»', '[/font];', '[/font].', '—[font=fonts/triodion_ucs]']
    solved_2   = ['...»[/font]', '»[/font]', ';[/font]', '.[/font]', '— [font=fonts/triodion_ucs]']

    mistakes3 = ['[', ']']
    solved3   = ['(', ')']

    mistakes4 = ['(b)', '(/b)', '(i)', '(/i)', '(font=fonts/triodion_ucs)', '(/font)', '(ref=world)', '(color=0000ff)', '(/color)', '(/ref)']
    solved4   = ['[b]', '[/b]', '[i]', '[/i]', '[font=fonts/triodion_ucs]', '[/font]', '[ref=world]', '[color=1b9bd4]', '[/color]', '[/ref]']

    mistakes5 = ['вечірн[', 'Светлої', 'Неделю', 'среду', 'жон-мироносиць', ' и ', 'віддяння', 'третья', 'пом0-лимсz', 'поми1-луй',  #'[font=fonts/triodion_ucs]заход[/font]',
                'с Крестом', 'по каждой песни', 'вечером', 'Примечание', 'Вхіду нема', 'Начало вечірні', 'вибрий', 'як На утрені', 'виголошуєся', 'зі своими', 'серпня На утрені']
    solved5   = ['вечірні[', 'Світлої', 'Неділю', 'середу', 'мироносиць', ' і ', 'віддання', 'третя', 'пом0лимсz', 'поми1луй',  #'[font=fonts/triodion_ucs]заход[/font]',
                'з Хрестом', 'після кожної пісні', 'увечері', '[b]Примітка[/b]', 'Входу немає', 'Початок вечірні', 'обраний', 'як на утрені', 'виголошується', 'зі своїми', 'серпня на утрені']
        
    all_mistakes = mistakes1 + mistakes2 + mistakes_2 + mistakes3 + mistakes4 + mistakes5
    main_mistakes = mistakes1 + mistakes2 + mistakes_2 + mistakes5

    all_solved = solved1 + solved2 + solved_2 + solved3 + solved4 + solved5
    main_solved = solved1 + solved2 + solved_2 + solved5

    spec_mistake = ['Ґллилyіа']
    spec_solution = ['Аллилyіа']

    def get_main_mistakes(self):
        return self.main_mistakes

    def get_all_mistakes(self):
        return self.all_mistakes
    
    def get_main_solutions(self):
        return self.main_solved

    def get_all_solutions(self):
        return self.all_solved


class Debugger():
    def solve_all(self, labels):
        labels = self.solve_problems_in_labels(labels)
        labels = self.add_free_line_to_labels(labels)
        return labels


    def add_old_style(label):
        try:
            old_style_range = label.find('\xa0')
            day_num = int(label[:old_style_range])
        except:
            old_style_range = label.find(' ') 
            day_num = int(label[:old_style_range])

        if day_num == 1:
            main_app.counter += 1

        label = str(day_num) + ' ' + monthes_day[main_app.counter] + ' (ст. стиль).' + label[old_style_range:]

        return label

    def remove_old_style(label):
        for day in days_name:
            text = day+'.'
            day_position = label.find(text)

            if day_position != -1:
                return label[ day_position + len(day) + 2: ]
            
        old_style_range = label.find(' ')
        if old_style_range > 4:
            old_style_range = label.find('\xa0')

        return label[old_style_range + 1 : ]


    def add_free_line_to_labels(self, labels):
        for k in range(len(labels)):
            labels[k] = self.clean_begin_and_end_of_label(labels[k])

            # Fix problem in cutted long label
            try:
                if labels[k].rfind('[font=fonts/triodion_ucs]') > labels[k].rfind('[/font]'):
                    labels[k] += '[/font]\uFEFF'
                    labels[k+1] = '\uFEFF[font=fonts/triodion_ucs]' + labels[k+1]
            except: pass
            
            # Add free line if it isnt cutted long label
            try:
                if type(labels[k]) is str:
                    if labels[k][-1] == '\uFEFF' or labels[k][-1] == ',':
                        pass
                    else:
                        labels[k] += '\n' 
            except: pass

        return labels

    def clean_begin_and_end_of_label(self, final_text):
        if type(final_text) is str:
            while final_text[0] == ' ':
                final_text = final_text[1:]
            while final_text[-1] == ' ' or final_text[-1] == '\n':
                if final_text[-1] == ' ':
                    final_text = final_text[:-1]
                if final_text[-1] == '\n':
                    final_text = final_text[:-1]
        return final_text

    def fix_alliluia(self, block):
        problem = 'Ґллилyіа'
        solution = 'Аллилyіа'

        new_block = ''

        block_parts = []
        part_start = 0
        part_end = block.find(problem)+8
        for x in range(str(block).count(problem)):
            block_parts.append(block[part_start:part_end])
            part_start = part_end
            if x != str(block).count(problem)-1:
                part_end = block[part_start:].find(problem)+8
            else: block_parts.append(block[part_start:])
            
            if block_parts[x].rfind('[font=fonts/triodion_ucs]') < block_parts[x].rfind('[/font]'):
                block_parts[x] = block_parts[x].replace(problem, solution)

        for part in block_parts:
            new_block += part

        return new_block

  
    def solve_problems_in_labels(self, old_labels):
        mistakes = Mistakes_and_solutions().get_all_mistakes()
        solutions = Mistakes_and_solutions().get_all_solutions()
        
        new_labels = []
    
        for label in old_labels:
            if label is not int or label != None:

                # Fix all problems
                for k in range(len(mistakes)):
                    while str(label).count(mistakes[k]):
                        label = str(label).replace(mistakes[k], solutions[k])
                
                # Fix problems with Alliluia
                if str(label).count('Ґллилyіа') > 0:
                    label = self.fix_alliluia(label)

            new_labels.append(label)

        return new_labels

    def count_problems_in_labels(self, labels):
        problems = Mistakes_and_solutions().get_main_mistakes()  #get_all_mistakes()
        problems_counter = []

        for problem_num in range(len(problems)):
            problems_counter.append(0)

        for label in labels:
            for problem_num in range(len(problems)):
                counter = str(label).count(problems[problem_num])
                problems_counter[problem_num] += counter
                if counter != 0:
                    print(label)

        for problem_num in range(len(problems)):
            if problems_counter[problem_num] != 0:
                print('% 8d '% problems_counter[problem_num], ' - ', [problems[problem_num]])


class Separator():
    def length_check(text):
        if len(text) < 1500:
            text = Debugger().clean_begin_and_end_of_label(text)
            return [text]
        else:
            return Separator.cut_long_label(text)

    def cut_long_label(text):
        first_end = text[:int(len(text)/3)].rfind('.') + 1
        sli = slice(0, first_end)
        text1 = text[sli]

        last_start = text[:int(len(text)*2/3)].rfind('.') + 1
        sli = slice(last_start, len(text))
        text3 = text[sli]

        sli = slice(first_end, last_start)
        text2 = text[sli]

        text1 = Debugger().clean_begin_and_end_of_label(text1)
        text2 = Debugger().clean_begin_and_end_of_label(text2)
        text3 = Debugger().clean_begin_and_end_of_label(text3)

        labels = []
        labels.append(text1)
        labels.append(text2)
        labels.append(text3)

        return labels

    def old_slice_labels(self, text, d, m, k):
        labels = []

        parts_num = int(len(text)/1000)+1
        part_len = 1000

        parts_start = 0
        parts_end = 0
        count = 0
        for t in range(parts_num+1):
            ik = t + count

            while True:

                if ik >= parts_num: 
                    parts_end = len(text)
                    break

                elif text[part_len*ik : part_len*(ik+1)].rfind('\n') == -1:
                    ik += 1

                else:
                    parts_end = text[part_len*(ik) : part_len*(ik+1)].rfind('\n') + (ik)*1000
                    break


            sli = slice(parts_start, parts_end)
            final_text = text[sli]

            parts_start = parts_end + 1

            final_text = Debugger().clean_begin_and_end_of_label(final_text)
            if final_text:
                if len(final_text) < 2000:
                    labels.append(final_text)
                else:
                    labels += Separator.cut_long_label(final_text)

        return labels

    def cut_label_into_paragraphs(text, status):
        input_text = text
        while input_text.count('\n \n'):
            input_text = input_text.replace('\n \n', '\n\n')

        if input_text.find('\n\n') == -1:
            new_labels = Separator.length_check(input_text)

            return new_labels

        labels = []
        paragraph_position = 0

        while True:
            paragraph_position = input_text.find('\n\n')
            new_labels = Separator.length_check(input_text[: paragraph_position])
            
            labels += new_labels

            input_text = input_text[paragraph_position +2 :]

            if input_text.find('\n\n') == -1:
                new_labels = input_text

                labels.append(new_labels)
                break
        
        return labels




class DBWorker():
    def create_request_for_create_db(all_rows_from_sheet):
        max_len = len(all_rows_from_sheet[0])

        #-----Create request to create calendar.db-----
        request_begin = 'CREATE TABLE IF NOT EXISTS Calendar ('
        request_main = ''
        request_end = ')'

        for column_num in range(max_len):

            #---Create names and formats of columnes
            if column_num == max_len - 1:
                column_format = ' text'
            else:
                column_format = ' text, '
            request_main += str(all_rows_from_sheet[0][column_num].value) + column_format

        request = request_begin + request_main + request_end

        return request
    
    def create_request_for_fill_db(all_rows_from_sheet):
        max_len = len(all_rows_from_sheet[0])

        request_begin = 'INSERT INTO Calendar ('
        request_column = ''
        request_variable_symb = ''
        request_end = ')'

        for column_num in range(max_len):
            
            #---Create names of columnes
            if column_num == max_len - 1:
                column_format = ') VALUES ('
            else:
                column_format = ', '
            request_column += str(all_rows_from_sheet[0][column_num].value) + column_format
            
            #---Create variable symb
            if column_num == max_len - 1:
                column_format = ''
            else:
                column_format = ', '
            request_variable_symb += '?' + column_format

        request = request_begin + request_column + request_variable_symb + request_end

        return request
    
    def create_tuple_with_all_values_of_exls_row(one_row_from_sheet):
        max_len = len(one_row_from_sheet)

        all_value = []
        for column_num in range(max_len):
            all_value.append(one_row_from_sheet[column_num].value)

        return tuple(all_value)


class PicklePacker():
    def pack_on_pickle(directory, filename, data):
        '''if os.path.exists(directory + desc_filename):
            os.remove(directory + desc_filename)'''
        os.makedirs(os.path.dirname(directory), exist_ok=True)
        with open(directory + filename, "wb") as f:
            pickle.dump(data, f)
    
class DBRangeHandler():

    # 0-5
    def first_range_processing(current_db, start, end):
        labels_to_pickle = []

        for k in range(start, end+1):
            try:
                # transform first columns from str to int: '2.0' -> 2
                labels_to_pickle.append(int(current_db[k][:-2]))
            except:
                if k == end:
                    # for last column remove day_num for short day in monthes
                    labels_to_pickle.append(Debugger.remove_old_style(current_db[k]))
                else:
                    labels_to_pickle.append(current_db[k])

        return labels_to_pickle

    # 5-...
    def second_range_processing(current_db, start, end):
        labels_text = []
        description = []

        # Add sancts and readings
        labels_text.append(Debugger.add_old_style(current_db[start]))
        labels_text.append(current_db[start+1])

        # Cut labels and add them with descr to output
        for k in range(start+2, end, 2):
            if current_db[k] != None:

                if int(current_db[0][:-2]) == 18 and int(current_db[1][:-2]) == 3 and k < 10:
                    #print([current_db[k]], '\n')
                    status = True
                else: status = False
                  
                new_labels = Separator.cut_label_into_paragraphs(current_db[k], status)
                labels_text += new_labels

            if current_db[k+1] != None: 
                description.append(current_db[k+1])

        return [labels_text, description]


class Navigator():
    def get_all_bs():
        night = 'На вечірні'
        great_night = 'На великій вечірні'
        great_afternight = 'На великому повечір\'ї'
        morning = 'На утрені'

        imagening = 'На зображальних'
        lpd = 'Літургії Передосвячених Дарів'
        hours = 'На часах'
        liturg = 'На Літургії'

        return [night, great_night, great_afternight, morning, imagening, lpd, hours, liturg]
        

    def find_points(d,m, labels):
        all_bs = Navigator.get_all_bs()
        
        # Create list [[bs, bs_label_position], ...]
        total = []
        for bs in all_bs:
            total.append([bs, 0])

        # Find all labels where bs are begining
        for position in range(len(labels)):
            for num in range(len(all_bs)):
                if total[num][1] == 0 and labels[position].find(total[num][0]) != -1:
                    total[num][1] = position

        return total

#sheet = wb['Лист1']
#print(wb.sheetnames, sheet.title)
#print(sheet['A1':'E3'])
#kk = sheet['A1'].value
#print(kk, type(kk))
#print(len(all_rows_from_sheet))

class main(App):
    def create_dbs(self):
        wb = openpyxl.load_workbook('calendar.xlsx')
        sheet = wb.active
        all_rows_from_sheet = tuple(sheet.rows) # generator -> tuple

        request_to_create_table = DBWorker.create_request_for_create_db(all_rows_from_sheet)
        request_to_fill_row = DBWorker.create_request_for_fill_db(all_rows_from_sheet)

        for mon in range(1, 13):

            #---Create folder and db's file for month
            adress = 'db/' + str(mon) + '.db'
            conn = sqlite3.connect(adress)
            cur = conn.cursor()  

            #---Create db's table for month
            cur.execute('''drop table if exists Calendar''')       
            cur.execute(f'''{request_to_create_table}''')
            
            #---Fill db's table for month
            for one_row_from_sheet in all_rows_from_sheet: #sheet['A2':'R367']             
                
                #---Take each row from exl's table with current month and put in db's row'
                if one_row_from_sheet[1].value == mon:

                    values_of_row = DBWorker.create_tuple_with_all_values_of_exls_row(one_row_from_sheet)
                    
                    cur.execute(f'''{request_to_fill_row}''', values_of_row)

                #---If exls table is end, then break
                elif one_row_from_sheet[1].value == None:
                    break

            conn.commit()
            conn.close()

        self.root.dbs.text = 'Done'

    def generate_labels(self):

        for m in range(1, 13):
            month = []

            for d in range(1, 32):

                if DataBase.get_day(d, m):
                    labels, desc = [], []

                    db = DataBase.get_day(d, m)

                    range1 = DBRangeHandler.first_range_processing(db, 0, 5)
                    month.append(range1)

                    range2 = DBRangeHandler.second_range_processing(db, 5, len(db))
                    labels = range2[0]
                    desc = range2[1]

                    labels = Debugger().solve_all(labels)
                    desc = Debugger().solve_all(desc)

                    navig = Navigator.find_points(d, m, labels)  


                    directory = "db/"+str(m)+'/'+str(d)+'/' 
                    text_filename = "text.pickle"
                    desc_filename = "desc.pickle"
                    navig_filename = "navig.pickle"
                    PicklePacker.pack_on_pickle(directory, text_filename, labels)
                    PicklePacker.pack_on_pickle(directory, desc_filename, desc)
                    PicklePacker.pack_on_pickle(directory, navig_filename, navig)
                
                else: pass

            directory = "db/"+str(m)+'/'
            month_filename = "month.pickle"
            PicklePacker.pack_on_pickle(directory, month_filename, month)
            
        self.root.gen.text = 'Done'

    def build(self):

        global main_app, d, m, screens, days_name, monthes_name, monthes_day, pr
        main_app = self  
        screens = []
        pr = cProfile.Profile()

        days_name = ['Понеділок', 'Вівторок', 'Середа', 'Четвер',
                        'П\'ятниця', 'Субота', 'Неділя']
        monthes_name = ['січень', 'лютий', 'березень',
                'квітень', 'травень', 'червень',
                'липень', 'серпень', 'вересень',
                'жовтень', 'листопад', 'грудень']
        monthes_day = ['січня', 'лютого', 'березня',
                'квітня', 'травня', 'червня',
                'липня', 'серпня', 'вересня',
                'жовтня', 'листопада', 'грудня']

        d = 18
        m = 1

        self.counter = -1   

        return super().build()   


if __name__ == '__main__':
    main().run()
