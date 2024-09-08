import openpyxl as xls
import webbrowser
import time
import datetime

'''links for modules documentation:
https://openpyxl.readthedocs.io/en/stable/tutorial.html#loading-from-a-file
https://docs.python.org/3/library/webbrowser.html'''

class Event:
    def __init__(self, name:str, date:datetime.date, description:str):
        self.name = name
        self.date = date
        self.description = description
    def get_name(self):
        return self.name
    def get_date(self):
        return self.date
    def get_description(self):
        return self.description
    
def get_events(filepath:str, sheet:str, col: int, line: int):
    '''Takes in the path of the file, the sheet containing the events, the column with the events names,
    the line with the first event.
    
    The columns MUST be in this order: event name, date, description.
    If not, the code WILL NOT work properly.'''

    wb = xls.load_workbook(filename=filepath)
    ws = wb[sheet]

    events_list = []
    while True:
        name = ws.cell(column=col, row=line).value
        date = ws.cell(column=col+1, row=line).value
        description = ws.cell(column=col+2, row=line).value

        if name == None: break
        if line > 100: break

        event = Event(name, date, description)

        events_list.append(event)
        line += 1

    return events_list

def search(filepath:str, sheet:str, col: int, line: int, search_term: str):
    '''Takes in the path of the file, the sheet containing the events, the column with the events names,
    the line with the first event, and a search term to specify the searches.
    
    The columns MUST be in this order: event name, date, description.
    If not, the code may not work properly.'''

    wb = xls.load_workbook(filename=filepath)    # open the planilha and the worksheet
    ws = wb[sheet]
    events = get_events(filepath, sheet, col, line)

    names_set = set()
    for event in events:
        event_name = event.get_name()
        names_set.add(event_name)

    for name in names_set:
        webbrowser.open(url=f'https://duckduckgo.com/?t=ffab&q={name}+{search_term}&atb=v388-1&ia=web')
        time.sleep(3)

    
#def order(filepath:str, sheet:str, col:int, line:int):
if __name__ == '__main__':
    planilha_path = 'C:/Users/USUARIO/VSCodeProjects/Planinha_Olimpiadas/pessoais/Olimpiadas_Copia.xlsx'
    print(get_events(filepath=planilha_path, sheet='2024 v2', col=2, line=3))